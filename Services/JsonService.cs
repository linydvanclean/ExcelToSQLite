using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace ExcelToSQLite.Services;

public class JsonService
{
    public async Task<List<List<object>>> ReadJsonAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            try
            {
                var jsonContent = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                
                // 尝试解析为JSON数组
                using var document = JsonDocument.Parse(jsonContent);
                var root = document.RootElement;
                
                var result = new List<List<object>>();
                
                // 判断是否为数组
                if (root.ValueKind == JsonValueKind.Array)
                {
                    return ParseJsonArray(root);
                }
                // 判断是否为对象数组
                else if (root.ValueKind == JsonValueKind.Object)
                {
                    // 尝试查找数组属性
                    foreach (var property in root.EnumerateObject())
                    {
                        if (property.Value.ValueKind == JsonValueKind.Array)
                        {
                            return ParseJsonArray(property.Value);
                        }
                    }
                    
                    // 如果是单个对象，转换为单行数据
                    return ParseSingleObject(root);
                }
                
                throw new Exception("不支持的JSON格式，请使用JSON数组或包含数组的对象");
            }
            catch (JsonException ex)
            {
                throw new Exception($"JSON解析失败: {ex.Message}");
            }
            catch (Exception ex)
            {
                throw new Exception($"读取JSON文件失败: {ex.Message}");
            }
        });
    }
    
    private List<List<object>> ParseJsonArray(JsonElement arrayElement)
    {
        var result = new List<List<object>>();
        var headers = new List<string>();
        var isFirstRow = true;
        
        foreach (var item in arrayElement.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object)
                continue;
            
            var row = new List<object>();
            var properties = new List<string>();
            
            // 收集所有属性名
            foreach (var property in item.EnumerateObject())
            {
                properties.Add(property.Name);
            }
            
            // 如果是第一行，设置表头（处理保留字段）
            if (isFirstRow)
            {
                headers = ProcessHeaders(properties);
                result.Add(new List<object>(headers));
                isFirstRow = false;
            }
            
            // 添加数据行
            var dataRow = new List<object>();
            foreach (var header in headers)
            {
                // 需要找到对应的原始属性名（因为headers可能已被修改）
                var originalPropertyName = GetOriginalPropertyName(header, properties);
                
                if (!string.IsNullOrEmpty(originalPropertyName) && 
                    item.TryGetProperty(originalPropertyName, out var value))
                {
                    dataRow.Add(GetJsonValue(value));
                }
                else
                {
                    dataRow.Add(string.Empty);
                }
            }
            result.Add(dataRow);
        }
        
        return result;
    }
    
    private List<List<object>> ParseSingleObject(JsonElement objectElement)
    {
        var result = new List<List<object>>();
        var headers = new List<string>();
        var row = new List<object>();
        var properties = new List<string>();
        
        foreach (var property in objectElement.EnumerateObject())
        {
            properties.Add(property.Name);
        }
        
        // 处理表头（保留字段）
        headers = ProcessHeaders(properties);
        result.Add(new List<object>(headers));
        
        // 添加数据行
        foreach (var header in headers)
        {
            var originalPropertyName = GetOriginalPropertyName(header, properties);
            
            if (!string.IsNullOrEmpty(originalPropertyName) && 
                objectElement.TryGetProperty(originalPropertyName, out var value))
            {
                row.Add(GetJsonValue(value));
            }
            else
            {
                row.Add(string.Empty);
            }
        }
        result.Add(row);
        
        return result;
    }
    
    /// <summary>
    /// 处理表头字段名：处理保留字段（id -> Id_1）
    /// </summary>
    /// <param name="headers">原始表头列表</param>
    /// <returns>处理后的表头列表</returns>
    private List<string> ProcessHeaders(List<string> headers)
    {
        var processedHeaders = new List<string>();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var idCounter = 1;
        
        // 定义需要检查的保留字段（不区分大小写）
        var reservedFields = new HashSet<string>(
            new[] { "id" }, 
            StringComparer.OrdinalIgnoreCase
        );
        
        foreach (var header in headers)
        {
            var headerName = header.Trim();
            
            // 检查是否为保留字段（id）
            if (reservedFields.Contains(headerName))
            {
                // 检查是否已经有同名的保留字段被处理过
                var newName = $"Id_{idCounter++}";
                processedHeaders.Add(newName);
                
                // 记录已使用的名称
                usedNames.Add(newName);
            }
            else
            {
                // 检查是否有重名冲突
                if (usedNames.Contains(headerName))
                {
                    // 如果已存在同名，添加数字后缀
                    var counter = 1;
                    var baseName = headerName;
                    string newName;
                    
                    do
                    {
                        newName = $"{baseName}_{counter++}";
                    } while (usedNames.Contains(newName));
                    
                    processedHeaders.Add(newName);
                    usedNames.Add(newName);
                }
                else
                {
                    processedHeaders.Add(headerName);
                    usedNames.Add(headerName);
                }
            }
        }
        
        return processedHeaders;
    }
    
    /// <summary>
    /// 获取原始属性名（从处理后的名称反向查找）
    /// </summary>
    /// <param name="processedName">处理后的字段名</param>
    /// <param name="originalNames">原始属性名列表</param>
    /// <returns>原始属性名，如果找不到则返回空字符串</returns>
    private string GetOriginalPropertyName(string processedName, List<string> originalNames)
    {
        // 如果处理后的名称以 "Id_" 开头，尝试匹配原始 "id"
        if (processedName.StartsWith("Id_", StringComparison.OrdinalIgnoreCase))
        {
            var idMatch = originalNames.FirstOrDefault(n => 
                n.Equals("id", StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(idMatch))
            {
                return idMatch;
            }
        }
        
        // 检查是否直接匹配
        if (originalNames.Contains(processedName))
        {
            return processedName;
        }
        
        // 尝试不区分大小写匹配
        var caseInsensitiveMatch = originalNames.FirstOrDefault(n => 
            n.Equals(processedName, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrEmpty(caseInsensitiveMatch))
        {
            return caseInsensitiveMatch;
        }
        
        // 如果处理后的名称包含下划线，尝试匹配基础名称
        var underscoreIndex = processedName.LastIndexOf('_');
        if (underscoreIndex > 0)
        {
            var baseName = processedName.Substring(0, underscoreIndex);
            var match = originalNames.FirstOrDefault(n => 
                n.Equals(baseName, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(match))
            {
                return match;
            }
        }
        
        return string.Empty;
    }
    
    private object GetJsonValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? string.Empty,
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => string.Empty,
            JsonValueKind.Object => element.GetRawText(),
            JsonValueKind.Array => element.GetRawText(),
            _ => element.GetRawText()
        };
    }
    
    public async Task<List<string>> GetColumnNamesAsync(string filePath)
    {
        var data = await ReadJsonAsync(filePath);
        if (data.Count > 0 && data[0] != null)
        {
            var columns = new List<string>();
            foreach (var item in data[0])
            {
                columns.Add(item?.ToString() ?? $"Column{columns.Count + 1}");
            }
            return columns;
        }
        return new List<string>();
    }
}