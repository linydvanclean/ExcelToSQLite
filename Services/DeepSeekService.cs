using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
using System.Diagnostics;
using ExcelToSQLite.Models;
using System.Collections.Generic;
using System.Threading;

namespace ExcelToSQLite.Services;

/// <summary>
/// Token使用信息
/// </summary>
public class UsageInfo
{
    [JsonPropertyName("prompt_tokens")]
    public int PromptTokens { get; set; }

    [JsonPropertyName("completion_tokens")]
    public int CompletionTokens { get; set; }

    [JsonPropertyName("total_tokens")]
    public int TotalTokens { get; set; }
}

/// <summary>
/// API使用信息
/// </summary>
public class ApiUsageInfo
{
    [JsonPropertyName("balance")]
    public decimal Balance { get; set; }

    [JsonPropertyName("total_usage")]
    public decimal TotalUsage { get; set; }

    [JsonPropertyName("available")]
    public bool IsAvailable { get; set; }
}

/// <summary>
/// DeepSeek AI API 服务 - 用于调用DeepSeek大模型生成SQL
/// </summary>
public class DeepSeekService
{
    private readonly HttpClient _httpClient;
    private readonly DeepSeekLogger _logger;
    private readonly Dictionary<string, AiSqlResponse> _cache = new();
    private readonly Dictionary<string, DateTime> _errorCache = new();
    private readonly object _cacheLock = new();
    private readonly object _errorCacheLock = new();
    private readonly SemaphoreSlim _semaphore = new(5);
    private readonly TimeSpan _errorCacheDuration = TimeSpan.FromMinutes(5);

    public DeepSeekService()
    {
        // 增加超时时间到60秒，减少超时风险
        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(60)
        };
        _logger = new DeepSeekLogger();
    }

    /// <summary>
    /// 综合性能与分析速度开展分析（带智能降级和重试）
    /// </summary>
    public async Task<AiSqlResponse?> GenerateSqlWithFallbackAsync(
        string prompt,
        string apiKey,
        string? apiEndpoint = null,
        int maxRetries = 2)
    {
        var requestId = Guid.NewGuid().ToString("N").Substring(0, 8);
        int retryCount = 0;
        int baseDelay = 1000;
        Exception? lastException = null;

        // 检查错误缓存
        if (IsInErrorCache(prompt))
        {
            _logger.LogWarning(requestId, "⚠️ 该请求最近失败过，直接使用 Pro 模型");
            return await GenerateSqlAsync(
                prompt, apiKey, apiEndpoint,
                model: "deepseek-v4-pro",
                reasoningEffort: "medium",
                maxTokens: 12288
            );
        }

        while (retryCount <= maxRetries)
        {
            try
            {
                // 第一次用 Flash，后续用 Pro
                var model = retryCount == 0 ? "deepseek-v4-flash" : "deepseek-v4-pro";
                var reasoning = retryCount == 0 ? "low" : "medium";
                var maxTokens = retryCount == 0 ? 8192 : 12288;

                _logger.LogInfo(requestId, 
                    $"🔄 尝试 {retryCount + 1}/{maxRetries + 1}: 模型={model}, reasoning={reasoning}");

                var result = await GenerateSqlAsync(
                    prompt, 
                    apiKey, 
                    apiEndpoint,
                    model: model,
                    reasoningEffort: reasoning,
                    maxTokens: maxTokens,
                    useCache: false
                );

                if (result != null && IsResultSatisfactory(result, prompt))
                {
                    _logger.LogSuccess(requestId, $"✅ 第 {retryCount + 1} 次尝试成功 (模型: {model})");
                    // 清除错误缓存
                    RemoveFromErrorCache(prompt);
                    return result;
                }

                // 结果不满足要求，继续重试
                _logger.LogWarning(requestId, $"⚠️ 第 {retryCount + 1} 次结果不满足要求");
                lastException = new Exception("结果质量不满足要求");
            }
            catch (TaskCanceledException) when (retryCount < maxRetries)
            {
                retryCount++;
                var delay = baseDelay * (int)Math.Pow(2, retryCount);
                _logger.LogWarning(requestId, 
                    $"⏳ 请求超时，{delay}ms 后重试 (第 {retryCount} 次，将使用 Pro)");
                await Task.Delay(delay);
                continue;
            }
            catch (HttpRequestException ex) when (retryCount < maxRetries)
            {
                retryCount++;
                lastException = ex;
                var delay = baseDelay * retryCount;
                _logger.LogWarning(requestId, $"⚠️ 网络请求失败: {ex.Message}，{delay}ms 后重试");
                await Task.Delay(delay);
                continue;
            }
            catch (Exception ex) when (retryCount < maxRetries)
            {
                retryCount++;
                lastException = ex;
                _logger.LogWarning(requestId, $"⚠️ 请求失败: {ex.Message}，重试中...");
                await Task.Delay(baseDelay * retryCount);
                continue;
            }

            retryCount++;
        }

        // 所有重试失败，记录错误缓存
        AddToErrorCache(prompt);
        
        var finalError = lastException?.Message ?? "所有重试均失败";
        _logger.LogError(requestId, $"❌ 所有 {maxRetries + 1} 次重试均失败: {finalError}");
        throw new Exception($"所有 {maxRetries + 1} 次重试均失败: {finalError}", lastException);
    }

    private bool IsResultSatisfactory(AiSqlResponse result, string prompt)
    {
        if (string.IsNullOrEmpty(result.SqlStatement)) 
            return false;

        // 检查 SQL 长度是否合理
        var sqlLength = result.SqlStatement.Length;
        if (sqlLength < 50) return false;

        // 检查是否包含必要的聚合函数
        var hasAggregate = new[] { "COUNT", "SUM", "AVG", "MAX", "MIN" }
            .Any(k => result.SqlStatement.Contains(k, StringComparison.OrdinalIgnoreCase));
        if (!hasAggregate) return false;

        // 检查是否包含数据清洗（CTE 或子查询）
        var hasCleaning = result.SqlStatement.Contains("WITH", StringComparison.OrdinalIgnoreCase) ||
                          result.SqlStatement.Contains("REPLACE", StringComparison.OrdinalIgnoreCase);

        // 如果 prompt 包含复杂要求但 SQL 没有清洗逻辑，可能需要升级
        var isComplex = prompt.Contains("分析") || prompt.Contains("统计") || prompt.Contains("异常");
        if (isComplex && !hasCleaning) return false;

        return true;
    }

    private bool IsInErrorCache(string prompt)
    {
        var key = prompt.GetHashCode().ToString();
        lock (_errorCacheLock)
        {
            if (_errorCache.TryGetValue(key, out var time))
            {
                if (DateTime.Now - time < _errorCacheDuration)
                    return true;
                _errorCache.Remove(key);
            }
        }
        return false;
    }

    private void AddToErrorCache(string prompt)
    {
        var key = prompt.GetHashCode().ToString();
        lock (_errorCacheLock)
        {
            _errorCache[key] = DateTime.Now;
        }
    }

    private void RemoveFromErrorCache(string prompt)
    {
        var key = prompt.GetHashCode().ToString();
        lock (_errorCacheLock)
        {
            _errorCache.Remove(key);
        }
    }

    /// <summary>
    /// 调用DeepSeek API生成SQL（带高级参数）
    /// </summary>
    public async Task<AiSqlResponse?> GenerateSqlAsync(
        string prompt,
        string apiKey,
        string? apiEndpoint = null,
        string? model = null,
        float temperature = 0.1f,
        int? maxTokens = null,
        float topP = 0.9f,
        bool isPreview = false,
        string? reasoningEffort = null,
        bool useCache = false)
    {
        // 并发控制
        await _semaphore.WaitAsync();
        try
        {
            // 缓存检查
            if (useCache && !isPreview)
            {
                var cacheKey = $"{prompt.GetHashCode()}_{apiKey.GetHashCode()}_{model}";
                lock (_cacheLock)
                {
                    if (_cache.TryGetValue(cacheKey, out var cached))
                    {
                        _logger.LogInfo("", $"✅ 使用缓存结果 (Key: {cacheKey.Substring(0, 8)})");
                        return cached;
                    }
                }
            }

            var endpoint = apiEndpoint ?? PublicEvent.DeepseekBaseUrl;
            var modelName = model ?? PublicEvent.DeepseekDefaultModel;
            var requestId = Guid.NewGuid().ToString("N").Substring(0, 8);
            
            // 开始日志记录
            _logger.LogRequestStart(requestId, modelName, endpoint, isPreview);

            // 智能参数优化
            var optimizedParams = GetOptimizedParams(prompt, modelName, temperature, maxTokens, topP, reasoningEffort);
            
            var systemPrompt = isPreview 
                ? GetPreviewSystemPrompt() 
                : GetApiSystemPrompt();
            
            _logger.LogInfo(requestId, $"参数: temp={optimizedParams.Temperature:F2}, maxTokens={optimizedParams.MaxTokens}, topP={optimizedParams.TopP:F2}, reasoning={optimizedParams.ReasoningEffort ?? "disabled"}");

            var request = new DeepSeekRequest
            {
                Model = modelName,
                Messages = new[]
                {
                    new DeepSeekMessage { Role = "system", Content = systemPrompt },
                    new DeepSeekMessage { Role = "user", Content = prompt }
                },
                ResponseFormat = isPreview ? null : new ResponseFormat { Type = "json_object" },
                Temperature = optimizedParams.Temperature,
                MaxTokens = optimizedParams.MaxTokens,
                TopP = optimizedParams.TopP,
                FrequencyPenalty = 0.1f,
                PresencePenalty = 0.1f,
                ReasoningEffort = optimizedParams.ReasoningEffort
            };

            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                WriteIndented = false
            };

            var requestJson = JsonSerializer.Serialize(request, jsonOptions);
            _logger.LogRequest(requestId, requestJson);

            var content = new StringContent(requestJson, Encoding.UTF8, "application/json");
            _httpClient.DefaultRequestHeaders.Clear();
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

            var stopwatch = Stopwatch.StartNew();
            var response = await _httpClient.PostAsync(endpoint, content);
            stopwatch.Stop();
            
            var responseBody = await response.Content.ReadAsStringAsync();
            
            // 记录响应基本信息
            _logger.LogResponse(requestId, response.StatusCode, stopwatch.ElapsedMilliseconds, responseBody);

            if (!response.IsSuccessStatusCode)
            {
                var errorMsg = $"DeepSeek API返回错误 ({(int)response.StatusCode}): {responseBody}";
                _logger.LogError(requestId, errorMsg);
                throw new Exception(errorMsg);
            }

            var deepSeekResponse = JsonSerializer.Deserialize<DeepSeekResponse>(responseBody, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower
            });

            if (deepSeekResponse?.Choices == null || deepSeekResponse.Choices.Length == 0)
            {
                var errorMsg = "DeepSeek API返回了空响应";
                _logger.LogError(requestId, errorMsg);
                throw new Exception(errorMsg);
            }

            var aiMessage = deepSeekResponse.Choices[0].Message?.Content ?? string.Empty;
            
            // 记录AI返回的原始消息
            _logger.LogAiMessage(requestId, aiMessage);
            
            // 记录token使用情况
            if (deepSeekResponse.Usage != null)
            {
                _logger.LogTokenUsage(requestId, deepSeekResponse.Usage);
                LogPerformance(
                    requestId, 
                    modelName, 
                    stopwatch.ElapsedMilliseconds, 
                    deepSeekResponse.Usage.TotalTokens,
                    false
                );
            }

            if (isPreview)
            {
                var previewResult = new AiSqlResponse
                {
                    RawContent = aiMessage,
                    IsPreview = true
                };
                _logger.LogSuccess(requestId, "预览模式完成");
                return previewResult;
            }

            // 尝试提取JSON
            var jsonContent = ExtractJson(aiMessage);
            _logger.LogExtractedJson(requestId, jsonContent);

            RawSqlResponse? rawResponse = null;
            try
            {
                rawResponse = JsonSerializer.Deserialize<RawSqlResponse>(jsonContent, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
            }
            catch (JsonException jsonEx)
            {
                _logger.LogError(requestId, $"JSON解析失败: {jsonEx.Message}");
            }

            // 创建结果对象
            var sqlResult = new AiSqlResponse
            {
                SqlStatement = string.Empty,
                SqlDetailData = string.Empty,
                Explanation = string.Empty,
                IsPreview = false
            };

            if (rawResponse != null)
            {
                sqlResult.SqlStatement = CleanSql(rawResponse.SqlStatement ?? string.Empty);
                sqlResult.SqlDetailData = CleanSql(rawResponse.SqlDetailData ?? string.Empty);
                sqlResult.Explanation = rawResponse.Explanation ?? string.Empty;
            }

            // 如果解析失败或SQL为空，尝试从文本中提取
            if (string.IsNullOrWhiteSpace(sqlResult.SqlStatement) && string.IsNullOrWhiteSpace(sqlResult.SqlDetailData))
            {
                _logger.LogWarning(requestId, "JSON解析失败，尝试从文本提取SQL");
                var extracted = ExtractSqlFromText(aiMessage);
                sqlResult.SqlStatement = CleanSql(extracted.StatSql ?? string.Empty);
                sqlResult.SqlDetailData = CleanSql(extracted.DetailSql ?? string.Empty);
                
                if (!string.IsNullOrWhiteSpace(sqlResult.SqlStatement) || !string.IsNullOrWhiteSpace(sqlResult.SqlDetailData))
                {
                    _logger.LogSuccess(requestId, $"从文本提取SQL成功 - 统计SQL长度: {sqlResult.SqlStatement.Length}, 明细SQL长度: {sqlResult.SqlDetailData.Length}");
                }
                else
                {
                    var errorMsg = "无法从AI返回内容中提取任何SQL语句";
                    _logger.LogError(requestId, errorMsg);
                    throw new Exception(errorMsg);
                }
            }
            else
            {
                _logger.LogSuccess(requestId, $"JSON解析成功 - 统计SQL长度: {sqlResult.SqlStatement.Length}, 明细SQL长度: {sqlResult.SqlDetailData.Length}");
            }

            // 缓存结果
            if (useCache && !isPreview && sqlResult != null)
            {
                var cacheKey = $"{prompt.GetHashCode()}_{apiKey.GetHashCode()}_{model}";
                lock (_cacheLock)
                {
                    _cache[cacheKey] = sqlResult;
                }
            }

            return sqlResult;
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError("", $"网络请求失败: {ex.Message}");
            throw;
        }
        catch (JsonException ex)
        {
            _logger.LogError("", $"API响应格式错误: {ex.Message}");
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError("", $"未预期错误: {ex.Message}\n堆栈: {ex.StackTrace}");
            throw;
        }
        finally
        {
            _semaphore.Release();
        }
    }
    
    /// <summary>
    /// 获取优化的参数配置（针对 v4-flash）
    /// </summary>
    private (float Temperature, int MaxTokens, float TopP, string? ReasoningEffort) 
        GetOptimizedParams(string prompt, string modelName, float temperature, int? maxTokens, float topP, string? reasoningEffort)
    {
        var isFlash = modelName.Contains("flash", StringComparison.OrdinalIgnoreCase);
    
        if (!isFlash)
        {
            var actualTemp = IsComplexQuery(prompt) ? 0.3f : 0.1f;
            var actualMaxTokens = maxTokens ?? 16384;
            return (actualTemp, actualMaxTokens, topP, reasoningEffort);
        }

        // --- v4-flash 专用优化 ---
        string? autoReasoning = reasoningEffort;
        if (autoReasoning == null)
        {
            autoReasoning = IsComplexQuery(prompt) ? "low" : null;
        }

        float optimizedTemperature = temperature;
        if (temperature == 0.1f)
        {
            optimizedTemperature = IsComplexQuery(prompt) ? 0.15f : 0.05f;
        }

        // 限制 maxTokens，避免超时
        int optimizedMaxTokens = maxTokens ?? 8192;
        if (IsComplexQuery(prompt) && prompt.Length > 500)
        {
            optimizedMaxTokens = 12288;
        }
        // 强制限制最大为 12288，避免超时
        if (optimizedMaxTokens > 12288)
        {
            optimizedMaxTokens = 12288;
            _logger.LogInfo("", $"⚠️ maxTokens 被限制为 12288 以避免超时");
        }

        float optimizedTopP = topP;
        if (topP == 0.9f)
        {
            optimizedTopP = IsComplexQuery(prompt) ? 0.85f : 0.75f;
        }

        return (optimizedTemperature, optimizedMaxTokens, optimizedTopP, autoReasoning);
    }

    /// <summary>
    /// 判断是否为复杂查询（增强版）
    /// </summary>
    private bool IsComplexQuery(string prompt)
    {
        var complexKeywords = new[] { 
            "统计", "汇总", "分析", "按", "分", "合计", "平均", "求和", "计数",
            "金额", "日期", "行业", "状态", "类型", "分类", "排名", "对比",
            "同比", "环比", "占比", "趋势", "分布", "区间", "范围"
        };
    
        var matchCount = complexKeywords.Count(k => 
            prompt.Contains(k, StringComparison.OrdinalIgnoreCase));
    
        if (matchCount >= 4) return true;
        if (matchCount >= 3 && prompt.Length > 200) return true;
        if (prompt.Length > 500) return true;
        
        return false;
    }
    
    private void LogPerformance(string requestId, string model, long elapsedMs, int tokens, bool isFallback)
    {
        var tokensPerSecond = elapsedMs > 0 ? tokens / (elapsedMs / 1000.0) : 0;
        var level = elapsedMs > 30000 ? "⚠️" : "✅";
        
        string performanceLevel;
        if (elapsedMs < 10000) performanceLevel = "🚀 极快";
        else if (elapsedMs < 20000) performanceLevel = "⚡ 快速";
        else if (elapsedMs < 30000) performanceLevel = "⏱️ 正常";
        else if (elapsedMs < 45000) performanceLevel = "🐢 较慢";
        else performanceLevel = "❌ 超时风险";

        _logger.LogInfo(requestId, 
            $"{level} 性能: {performanceLevel} | 模型={model}, 耗时={elapsedMs}ms, " +
            $"Tokens={tokens}, 速度={tokensPerSecond:F1} tokens/s, 降级={(isFallback ? "是" : "否")}");

        if (elapsedMs > 40000)
        {
            _logger.LogWarning(requestId, $"⚠️ 响应接近超时阈值 ({elapsedMs}ms)，建议优化提示词或升级模型");
        }
    }

    /// <summary>
    /// 获取预览模式的系统提示词（优化版 - 精简）
    /// </summary>
    private string GetPreviewSystemPrompt()
    {
        return @"你是SQLite专家。生成清晰规范的SQL查询。

【规范】
1. 表名/字段名用双引号包裹: ""table""
2. 金额: REPLACE去千分位, CAST转数值, ROUND保留2位
3. 日期: date()函数统一格式, COALESCE处理空值
4. NULL: 使用COALESCE或IFNULL

【输出】
返回两个SQL:
1. 统计SQL - 聚合汇总
2. 明细SQL - 完整明细

直接返回SQL语句，包含注释。";
    }

    /// <summary>
    /// 获取API模式的系统提示词（优化版 - 精简，含别名规范）
    /// </summary>
    private string GetApiSystemPrompt()
    {
        return @"你是SQLite专家。必须返回纯JSON格式。

【JSON格式】
{
    ""sql_statement"": ""统计SQL"",
    ""sql_detaildata"": ""明细SQL"",
    ""explanation"": ""说明""
}

【重要】SQLite 字段别名规范：
- 避免使用特殊字符（如 -、_ 除外）
- 如果必须使用中文别名，用反引号包裹：`1万以下_份数`
- 推荐使用英文别名：amount_less_than_10000

【规范】
1. 表名/字段名: 双引号包裹
2. 金额: REPLACE去千分位, CAST转数值, ROUND(,2)
3. 日期: date(), COALESCE处理空值
4. SQL中的双引号用反斜杠转义
5. 使用 CTE 进行数据清洗，避免多个 UNION

【示例】
{""sql_statement"":""SELECT COUNT(*) FROM \""table\"" WHERE \""date\"" >= '2024-01-01'"",""sql_detaildata"":""SELECT * FROM \""table\"" ORDER BY \""date\"" DESC"",""explanation"":""统计2024年后记录""}

只返回JSON，无其他文本。";
    }

    /// <summary>
    /// 从AI响应文本中提取JSON内容（增强版）
    /// </summary>
    private string ExtractJson(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return "{}";

        var jsonMarkers = new[] { "```json", "```JSON", "```javascript", "```" };
        foreach (var marker in jsonMarkers)
        {
            var startIndex = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (startIndex >= 0)
            {
                var start = startIndex + marker.Length;
                var end = text.IndexOf("```", start, StringComparison.Ordinal);
                if (end > start)
                {
                    var extracted = text.Substring(start, end - start).Trim();
                    if (IsValidJson(extracted))
                        return extracted;
                }
            }
        }

        var firstBrace = text.IndexOf('{');
        var lastBrace = text.LastIndexOf('}');
        if (firstBrace >= 0 && lastBrace > firstBrace)
        {
            var extracted = text.Substring(firstBrace, lastBrace - firstBrace + 1);
            if (IsValidJson(extracted))
                return extracted;
        }

        var firstBracket = text.IndexOf('[');
        var lastBracket = text.LastIndexOf(']');
        if (firstBracket >= 0 && lastBracket > firstBracket)
        {
            var extracted = text.Substring(firstBracket, lastBracket - firstBracket + 1);
            if (IsValidJson(extracted))
                return extracted;
        }

        return text;
    }

    private bool IsValidJson(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;
    
        try
        {
            JsonDocument.Parse(text);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 从文本中提取SQL（备用方案）
    /// </summary>
    private (string? StatSql, string? DetailSql) ExtractSqlFromText(string text)
    {
        string? statSql = null;
        string? detailSql = null;
        
        var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
        var currentSql = new List<string>();
        bool isInSql = false;
        bool isStat = false;
        bool isDetail = false;
        
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            
            if (trimmed.Contains("统计SQL", StringComparison.OrdinalIgnoreCase) || 
                trimmed.Contains("统计", StringComparison.OrdinalIgnoreCase))
            {
                if (currentSql.Count > 0 && isInSql)
                {
                    if (isStat) statSql = string.Join("\n", currentSql);
                    else if (isDetail) detailSql = string.Join("\n", currentSql);
                }
                currentSql.Clear();
                isStat = true;
                isDetail = false;
                isInSql = true;
                continue;
            }
            
            if (trimmed.Contains("明细SQL", StringComparison.OrdinalIgnoreCase) || 
                trimmed.Contains("明细", StringComparison.OrdinalIgnoreCase))
            {
                if (currentSql.Count > 0 && isInSql)
                {
                    if (isStat) statSql = string.Join("\n", currentSql);
                    else if (isDetail) detailSql = string.Join("\n", currentSql);
                }
                currentSql.Clear();
                isStat = false;
                isDetail = true;
                isInSql = true;
                continue;
            }
            
            if (isInSql && !string.IsNullOrWhiteSpace(trimmed))
            {
                if (!trimmed.StartsWith("--", StringComparison.Ordinal) || 
                    (trimmed.StartsWith("--", StringComparison.Ordinal) && trimmed.Contains("======")))
                {
                    currentSql.Add(trimmed);
                }
            }
        }
        
        if (currentSql.Count > 0 && isInSql)
        {
            if (isStat) statSql = string.Join("\n", currentSql);
            else if (isDetail) detailSql = string.Join("\n", currentSql);
        }
        
        return (statSql, detailSql);
    }

    /// <summary>
    /// 清理SQL语句
    /// </summary>
    private string CleanSql(string? sql)
    {
        if (string.IsNullOrWhiteSpace(sql))
            return string.Empty;

        var cleaned = sql
            .Replace("\r\n", "\n")
            .Replace("\r", "\n")
            .Trim();

        if (cleaned.StartsWith("\"") && cleaned.EndsWith("\""))
        {
            cleaned = cleaned.Substring(1, cleaned.Length - 2);
        }

        while (cleaned.Contains("\n\n\n"))
        {
            cleaned = cleaned.Replace("\n\n\n", "\n\n");
        }

        return cleaned;
    }

    /// <summary>
    /// 清理缓存
    /// </summary>
    public void ClearCache()
    {
        lock (_cacheLock)
        {
            _cache.Clear();
            _logger.LogInfo("", "缓存已清空");
        }
        lock (_errorCacheLock)
        {
            _errorCache.Clear();
            _logger.LogInfo("", "错误缓存已清空");
        }
    }

    #region API Models

    public class AiSqlResponse
    {
        [JsonPropertyName("sqlStatement")]
        public string SqlStatement { get; set; } = string.Empty;

        [JsonPropertyName("sqlDetailData")]
        public string SqlDetailData { get; set; } = string.Empty;

        [JsonPropertyName("explanation")]
        public string Explanation { get; set; } = string.Empty;

        [JsonIgnore]
        public bool IsPreview { get; set; }

        [JsonIgnore]
        public string? RawContent { get; set; }
    }

    private class RawSqlResponse
    {
        [JsonPropertyName("sql_statement")]
        public string? SqlStatement { get; set; }

        [JsonPropertyName("sql_detaildata")]
        public string? SqlDetailData { get; set; }

        [JsonPropertyName("explanation")]
        public string? Explanation { get; set; }
    }

    private class DeepSeekRequest
    {
        [JsonPropertyName("model")]
        public string Model { get; set; } = PublicEvent.DeepseekDefaultModel;

        [JsonPropertyName("messages")]
        public DeepSeekMessage[] Messages { get; set; } = Array.Empty<DeepSeekMessage>();

        [JsonPropertyName("response_format")]
        public ResponseFormat? ResponseFormat { get; set; }

        [JsonPropertyName("temperature")]
        public float Temperature { get; set; } = 0.1f;

        [JsonPropertyName("max_tokens")]
        public int? MaxTokens { get; set; }

        [JsonPropertyName("top_p")]
        public float TopP { get; set; } = 0.9f;

        [JsonPropertyName("frequency_penalty")]
        public float FrequencyPenalty { get; set; } = 0.1f;

        [JsonPropertyName("presence_penalty")]
        public float PresencePenalty { get; set; } = 0.1f;
        
        [JsonPropertyName("reasoning_effort")]
        public string? ReasoningEffort { get; set; }
    }

    private class ResponseFormat
    {
        [JsonPropertyName("type")]
        public string Type { get; set; } = "json_object";
    }

    private class DeepSeekMessage
    {
        [JsonPropertyName("role")]
        public string Role { get; set; } = "user";

        [JsonPropertyName("content")]
        public string Content { get; set; } = string.Empty;
    }

    private class DeepSeekResponse
    {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("choices")]
        public DeepSeekChoice[]? Choices { get; set; }

        [JsonPropertyName("usage")]
        public UsageInfo? Usage { get; set; }
    }

    private class DeepSeekChoice
    {
        [JsonPropertyName("index")]
        public int Index { get; set; }

        [JsonPropertyName("message")]
        public DeepSeekMessage? Message { get; set; }

        [JsonPropertyName("finish_reason")]
        public string? FinishReason { get; set; }
    }

    #endregion
}

/// <summary>
/// DeepSeek 日志记录器
/// </summary>
public class DeepSeekLogger
{
    private readonly string _logDirectory;
    private readonly object _lockObject = new object();

    public DeepSeekLogger()
    {
        _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs", "DeepSeek");
        
        if (!Directory.Exists(_logDirectory))
        {
            Directory.CreateDirectory(_logDirectory);
        }
    }

    private string GetLogFilePath()
    {
        var date = DateTime.Now.ToString("yyyy-MM-dd");
        return Path.Combine(_logDirectory, $"deepseek_{date}.log");
    }

    private void WriteLog(string message)
    {
        lock (_lockObject)
        {
            try
            {
                var logFile = GetLogFilePath();
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                var logEntry = $"[{timestamp}] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, logEntry, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"日志写入失败: {ex.Message}");
            }
        }
    }

    public void LogRequestStart(string requestId, string model, string endpoint, bool isPreview)
    {
        WriteLog($"===== 请求开始 [ID: {requestId}] =====");
        WriteLog($"模型: {model}");
        WriteLog($"端点: {endpoint}");
        WriteLog($"预览模式: {isPreview}");
    }

    public void LogRequest(string requestId, string requestJson)
    {
        WriteLog($"--- 请求JSON [ID: {requestId}] ---");
        WriteLog(requestJson);
    }

    public void LogResponse(string requestId, System.Net.HttpStatusCode statusCode, long elapsedMs, string responseBody)
    {
        WriteLog($"--- 响应 [ID: {requestId}] ---");
        WriteLog($"状态码: {(int)statusCode} {statusCode}");
        WriteLog($"耗时: {elapsedMs}ms");
        WriteLog($"原始响应内容 (前1000字符):");
        var truncated = responseBody.Length > 1000 ? responseBody.Substring(0, 1000) + "..." : responseBody;
        WriteLog(truncated);
    }

    public void LogAiMessage(string requestId, string aiMessage)
    {
        WriteLog($"--- AI返回消息 [ID: {requestId}] ---");
        WriteLog($"消息长度: {aiMessage.Length} 字符");
        WriteLog($"完整消息内容:");
        WriteLog(aiMessage);
    }

    public void LogExtractedJson(string requestId, string jsonContent)
    {
        WriteLog($"--- 提取的JSON [ID: {requestId}] ---");
        WriteLog($"JSON长度: {jsonContent.Length} 字符");
        WriteLog($"JSON内容:");
        WriteLog(jsonContent);
    }

    public void LogTokenUsage(string requestId, UsageInfo usage)
    {
        WriteLog($"--- Token使用 [ID: {requestId}] ---");
        WriteLog($"提示Token: {usage.PromptTokens}");
        WriteLog($"完成Token: {usage.CompletionTokens}");
        WriteLog($"总Token: {usage.TotalTokens}");
    }

    public void LogSuccess(string requestId, string message)
    {
        WriteLog($"✅ 成功 [ID: {requestId}]: {message}");
        WriteLog($"===== 请求结束 [ID: {requestId}] ====={Environment.NewLine}");
    }

    public void LogWarning(string requestId, string message)
    {
        WriteLog($"⚠️ 警告 [ID: {requestId}]: {message}");
    }

    public void LogError(string requestId, string errorMessage)
    {
        WriteLog($"❌ 错误 [ID: {requestId}]: {errorMessage}");
        WriteLog($"===== 请求失败 [ID: {requestId}] ====={Environment.NewLine}");
    }

    public void LogInfo(string requestId, string message)
    {
        WriteLog($"ℹ️ 信息 [ID: {requestId}]: {message}");
    }
}