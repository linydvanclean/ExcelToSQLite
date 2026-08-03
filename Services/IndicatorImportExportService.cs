using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services;

/// <summary>
/// 导入结果
/// </summary>
public class ImportResult
{
    public int TotalCount { get; set; }
    public int ImportedCount { get; set; }
    public int SkippedCount { get; set; }
    public int OverwrittenCount { get; set; }
    public int EmptySqlStatement { get; set; }
    public int EmptySqlDetailData { get; set; }
    public List<string> SkippedIndicators { get; set; } = new();
    public List<string> Errors { get; set; } = new();

    public string GetSummary()
    {
        var parts = new List<string>();
        
        // 统计信息
        if (TotalCount > 0)
        {
            parts.Add($"总计 {TotalCount} 个指标");
        }
        if (EmptySqlDetailData > 0)
        {
            parts.Add($"SQL 详细数据为空 {EmptySqlDetailData} 个");
        }
        if (EmptySqlStatement > 0)
        {
            parts.Add($"SQL 语句为空 {EmptySqlStatement} 个");
        }
        
        // 操作结果
        if (ImportedCount > 0) parts.Add($"新增 {ImportedCount} 个");
        if (OverwrittenCount > 0) parts.Add($"覆盖 {OverwrittenCount} 个");
        if (SkippedCount > 0) parts.Add($"跳过 {SkippedCount} 个");
        if (Errors.Any()) parts.Add($"失败 {Errors.Count} 个");
        
        var summary = string.Join("，", parts);
        return string.IsNullOrEmpty(summary) ? "没有导入任何指标" : summary;
    }

    public bool HasErrors => Errors.Any();
    public bool IsSuccess => !HasErrors && (ImportedCount > 0 || OverwrittenCount > 0);
}

/// <summary>
/// 导出摘要信息
/// </summary>
public class ExportSummary
{
    public int TotalCount { get; set; }
    public int EmptySqlStatement { get; set; }
    public int EmptySqlDetailData { get; set; }
    public DateTime ExportTime { get; set; }
    public string ExportedBy { get; set; } = string.Empty;
    public string Version { get; set; } = string.Empty;
    
    public string GetSummaryMessage()
    {
        var messages = new List<string>();
        messages.Add($"总共 {TotalCount} 个指标");
        
        if (EmptySqlStatement > 0)
        {
            messages.Add($"{EmptySqlStatement} 个指标 SQL 语句为空");
        }
        if (EmptySqlDetailData > 0)
        {
            messages.Add($"{EmptySqlDetailData} 个指标 SQL 详细数据为空");
        }
        
        return string.Join("，", messages);
    }
}

/// <summary>
/// 指标导入导出服务
/// </summary>
public class IndicatorImportExportService : IDisposable
{
    private readonly IndicatorService _indicatorService;
    private readonly string _fileExtension = ".van";
    private readonly JsonSerializerOptions _jsonOptions;
    private bool _disposed;

    public IndicatorImportExportService(IndicatorService indicatorService)
    {
        _indicatorService = indicatorService ?? throw new ArgumentNullException(nameof(indicatorService));
        
        _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };
    }

    public string FileExtension => _fileExtension;

    /// <summary>
    /// 导出指标
    /// </summary>
    public string ExportIndicators(List<Indicator> indicators, string exportedBy)
    {
        if (indicators == null)
            throw new ArgumentNullException(nameof(indicators));
            
        if (indicators.Count == 0)
            throw new ArgumentException("没有要导出的指标，请至少选择一个指标", nameof(indicators));

        if (string.IsNullOrEmpty(exportedBy))
            throw new ArgumentException("导出人不能为空", nameof(exportedBy));

        try
        {
            // 检查是否有空名称的指标
            var emptyNameIndicators = indicators.Where(i => string.IsNullOrEmpty(i.Name)).ToList();
            if (emptyNameIndicators.Any())
            {
                throw new Exception($"发现 {emptyNameIndicators.Count} 个指标名称为空，请先完善指标信息");
            }

            // 统计信息
            var totalCount = indicators.Count;
            var emptySqlStatement = indicators.Count(i => string.IsNullOrWhiteSpace(i.SqlStatement));
            var emptySqlDetailData = indicators.Count(i => string.IsNullOrWhiteSpace(i.SqlDetailData));

            var exportData = new IndicatorExportData
            {
                Version = "1.0",
                ExportTime = DateTime.Now,
                ExportedBy = exportedBy,
                ExportSummary = new ExportSummary
                {
                    TotalCount = totalCount,
                    EmptySqlStatement = emptySqlStatement,
                    EmptySqlDetailData = emptySqlDetailData,
                    ExportTime = DateTime.Now,
                    ExportedBy = exportedBy,
                    Version = "1.0"
                },
                Indicators = indicators.Select(i => new IndicatorExportItem
                {
                    Id = i.Id,
                    Name = i.Name ?? string.Empty,
                    SqlStatement = i.SqlStatement ?? string.Empty,
                    SqlDetailData = i.SqlDetailData ?? string.Empty,
                    Description = i.Description ?? string.Empty,
                    Category = i.Category ?? string.Empty,
                    CreatedAt = i.CreatedAt,
                    UpdatedAt = i.UpdatedAt,
                    CreatedBy = i.CreatedBy ?? "admin",
                    IsActive = i.IsActive
                }).ToList()
            };

            return JsonSerializer.Serialize(exportData, _jsonOptions);
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new Exception($"导出指标失败: {ex.Message}", ex);
        }
    }
    
    /// <summary>
    /// 导入指标（异步版本，避免 UI 线程死锁）
    /// 优化：直接追加记录，不检查重复，不覆盖已有记录
    /// </summary>
    public async Task<ImportResult> ImportIndicatorsAsync(string fileContent)
    {
        if (string.IsNullOrEmpty(fileContent))
            throw new ArgumentException("文件内容不能为空", nameof(fileContent));
    
        try
        {
            // 清理可能存在的 BOM 头
            fileContent = RemoveBom(fileContent);
            
            // 验证文件格式
            if (!ValidateVanFile(fileContent))
            {
                // 尝试更宽松的验证
                if (!TryParseAsIndicatorExportData(fileContent, out var testData))
                {
                    throw new Exception("文件格式无效，不是有效的 .van 文件格式。请确保文件是有效的 JSON 格式且包含指标数据。");
                }
            }
    
            var exportData = JsonSerializer.Deserialize<IndicatorExportData>(fileContent, _jsonOptions);
    
            if (exportData == null || exportData.Indicators == null || !exportData.Indicators.Any())
            {
                throw new Exception("文件中没有可导入的指标数据");
            }
    
            if (exportData.Version != "1.0")
            {
                throw new Exception($"不支持的文件版本: {exportData.Version}，当前仅支持版本 1.0");
            }
    
            var result = new ImportResult
            {
                TotalCount = exportData.Indicators.Count,
                EmptySqlDetailData = exportData.Indicators.Count(i => string.IsNullOrWhiteSpace(i.SqlDetailData)),
                EmptySqlStatement = exportData.Indicators.Count(i => string.IsNullOrWhiteSpace(i.SqlStatement))
            };
    
            var importedCount = 0;
            var errors = new List<string>();
    
            foreach (var item in exportData.Indicators)
            {
                try
                {
                    // 检查指标名称是否有效
                    if (string.IsNullOrEmpty(item.Name))
                    {
                        errors.Add($"指标名称不能为空 (ID: '{item.Id}')");
                        continue;
                    }
    
                    // 直接创建新指标（追加记录），不检查是否已存在
                    var indicator = new Indicator
                    {
                        // 注意：不设置 Id，让数据库自动生成
                        Name = item.Name,
                        SqlStatement = item.SqlStatement ?? string.Empty,
                        SqlDetailData = item.SqlDetailData ?? string.Empty,
                        Description = item.Description ?? string.Empty,
                        Category = string.IsNullOrEmpty(item.Category) ? "未分类" : item.Category,
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now,
                        CreatedBy = string.IsNullOrEmpty(item.CreatedBy) ? "import" : item.CreatedBy,
                        IsActive = item.IsActive
                    };
    
                    await _indicatorService.AddAsync(indicator);
                    importedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"导入指标 '{item.Name}' 失败: {ex.Message}");
                }
            }
    
            result.ImportedCount = importedCount;
            result.SkippedCount = 0; // 不再跳过任何记录
            result.OverwrittenCount = 0; // 不再覆盖任何记录
            result.SkippedIndicators = new List<string>(); // 清空跳过的指标列表
            result.Errors = errors;
    
            return result;
        }
        catch (JsonException ex)
        {
            throw new Exception($"文件格式错误，不是有效的 JSON 格式。请确保文件内容完整无误。\n详细信息: {ex.Message}", ex);
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new Exception($"导入指标失败: {ex.Message}", ex);
        }
    }
    
    /// <summary>
    /// 移除 BOM 头
    /// </summary>
    private string RemoveBom(string content)
    {
        if (string.IsNullOrEmpty(content))
            return content;

        // 检查是否包含 BOM (UTF-8 BOM: EF BB BF)
        if (content.Length >= 3 && content[0] == '\uFEFF')
        {
            return content.Substring(1);
        }
        
        // 检查是否包含其他 BOM 标记
        if (content.Length >= 2 && content[0] == '\uFFFE')
        {
            return content.Substring(1);
        }

        return content;
    }

    /// <summary>
    /// 验证 .van 文件格式
    /// </summary>
    public bool ValidateVanFile(string fileContent)
    {
        if (string.IsNullOrEmpty(fileContent))
            return false;

        try
        {
            // 清理 BOM
            var cleanContent = RemoveBom(fileContent);
            var exportData = JsonSerializer.Deserialize<IndicatorExportData>(cleanContent, _jsonOptions);
            
            return exportData != null && 
                   !string.IsNullOrEmpty(exportData.Version) && 
                   exportData.Indicators != null &&
                   exportData.Indicators.Any() &&
                   exportData.Indicators.All(i => !string.IsNullOrEmpty(i.Name));
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 尝试解析为指标导出数据（更宽松的验证）
    /// </summary>
    private bool TryParseAsIndicatorExportData(string fileContent, out IndicatorExportData? data)
    {
        data = null;
        try
        {
            var cleanContent = RemoveBom(fileContent);
            var result = JsonSerializer.Deserialize<IndicatorExportData>(cleanContent, _jsonOptions);
            
            if (result != null && result.Indicators != null && result.Indicators.Any())
            {
                data = result;
                return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取导出文件的摘要信息（用于预览）
    /// </summary>
    public ExportSummary? GetExportSummary(string fileContent)
    {
        if (string.IsNullOrEmpty(fileContent))
            return null;

        try
        {
            var cleanContent = RemoveBom(fileContent);
            var exportData = JsonSerializer.Deserialize<IndicatorExportData>(cleanContent, _jsonOptions);
            if (exportData == null || exportData.Indicators == null)
                return null;

            return new ExportSummary
            {
                TotalCount = exportData.Indicators.Count,
                EmptySqlStatement = exportData.Indicators.Count(i => string.IsNullOrWhiteSpace(i.SqlStatement)),
                EmptySqlDetailData = exportData.Indicators.Count(i => string.IsNullOrWhiteSpace(i.SqlDetailData)),
                ExportTime = exportData.ExportTime,
                ExportedBy = exportData.ExportedBy ?? "未知",
                Version = exportData.Version ?? "未知"
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 检查是否需要覆盖（用于导入前判断）
    /// </summary>
    public async Task<bool> HasConflictsAsync(string fileContent)
    {
        if (string.IsNullOrEmpty(fileContent))
            return false;

        try
        {
            var cleanContent = RemoveBom(fileContent);
            var exportData = JsonSerializer.Deserialize<IndicatorExportData>(cleanContent, _jsonOptions);
            if (exportData == null || exportData.Indicators == null || !exportData.Indicators.Any())
                return false;

            var existingIndicators = await _indicatorService.GetAllAsync();
            var existingNames = existingIndicators.Select(i => i.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

            return exportData.Indicators.Any(i => existingNames.Contains(i.Name));
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // 清理托管资源
            }
            _disposed = true;
        }
    }
}

/// <summary>
/// 导出数据模型
/// </summary>
public class IndicatorExportData
{
    public string Version { get; set; } = "1.0";
    public DateTime ExportTime { get; set; }
    public string ExportedBy { get; set; } = string.Empty;
    public ExportSummary? ExportSummary { get; set; }
    public List<IndicatorExportItem> Indicators { get; set; } = new();
}

/// <summary>
/// 导出指标项
/// </summary>
public class IndicatorExportItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string SqlStatement { get; set; } = string.Empty;
    public string SqlDetailData { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }
    public string CreatedBy { get; set; } = string.Empty;
    public bool IsActive { get; set; }
}