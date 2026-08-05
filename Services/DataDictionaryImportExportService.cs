using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services;

public class DataDictionaryImportExportService : IDisposable
{
    private readonly DataDictionaryService _dictionaryService;
    private readonly string _fileExtension = ".vdd"; // VanClean Data Dictionary
    private readonly JsonSerializerOptions _jsonOptions;

    public DataDictionaryImportExportService()
    {
        _dictionaryService = new DataDictionaryService();
        _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };
    }

    public string FileExtension => _fileExtension;

    public string ExportDictionaries(List<DataDictionary> dictionaries, string exportedBy)
    {
        if (dictionaries == null || dictionaries.Count == 0)
            throw new ArgumentException("没有要导出的数据字典", nameof(dictionaries));

        try
        {
            var exportData = new DataDictionaryExportData
            {
                Version = PublicEvent.Version,
                ExportTime = DateTime.Now,
                ExportedBy = exportedBy,
                Dictionaries = dictionaries.Select(d => new DataDictionaryItem
                {
                    Name = d.Name,
                    TableName = d.TableName,
                    Description = d.Description ?? string.Empty,
                    CreatedAt = d.CreatedAt,
                    CreatedBy = d.CreatedBy ?? "admin",
                    IsActive = d.IsActive
                }).ToList()
            };

            return JsonSerializer.Serialize(exportData, _jsonOptions);
        }
        catch (Exception ex)
        {
            throw new Exception($"导出数据字典失败: {ex.Message}", ex);
        }
    }

    public async Task<ImportResult> ImportDictionariesAsync(string fileContent, bool overwriteExisting = false)
    {
        if (string.IsNullOrEmpty(fileContent))
            throw new ArgumentException("文件内容不能为空", nameof(fileContent));

        try
        {
            if (!ValidateVanFile(fileContent))
            {
                throw new Exception("文件格式无效，不是有效的 .vdd 文件");
            }

            var exportData = JsonSerializer.Deserialize<DataDictionaryExportData>(fileContent, _jsonOptions);

            if (exportData == null || exportData.Dictionaries == null || !exportData.Dictionaries.Any())
            {
                throw new Exception("文件中没有可导入的数据字典数据");
            }

            if (exportData.Version != "1.0")
            {
                throw new Exception($"不支持的文件版本: {exportData.Version}");
            }

            var result = new ImportResult
            {
                TotalCount = exportData.Dictionaries.Count
            };

            // ✅ 修复: 使用 await 替代 .GetAwaiter().GetResult()
            var existingDictionaries = await _dictionaryService.GetAllAsync();
            var existingNameMap = existingDictionaries.ToDictionary(d => d.Name, d => d, StringComparer.OrdinalIgnoreCase);

            int importedCount = 0;
            int skippedCount = 0;
            int overwrittenCount = 0;
            var errors = new List<string>();

            foreach (var item in exportData.Dictionaries)
            {
                try
                {
                    if (existingNameMap.TryGetValue(item.Name, out var existing))
                    {
                        if (overwriteExisting)
                        {
                            existing.TableName = item.TableName;
                            existing.Description = item.Description ?? string.Empty;
                            existing.UpdatedAt = DateTime.Now;
                            existing.IsActive = item.IsActive;

                            await _dictionaryService.UpdateAsync(existing);
                            overwrittenCount++;
                        }
                        else
                        {
                            skippedCount++;
                            result.SkippedIndicators.Add(item.Name);
                        }
                        continue;
                    }

                    var dictionary = new DataDictionary
                    {
                        Name = item.Name,
                        TableName = item.TableName,
                        Description = item.Description ?? string.Empty,
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now,
                        CreatedBy = item.CreatedBy ?? "import",
                        IsActive = item.IsActive
                    };

                    await _dictionaryService.AddAsync(dictionary);
                    importedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"导入数据字典 '{item.Name}' 失败: {ex.Message}");
                }
            }

            result.ImportedCount = importedCount;
            result.SkippedCount = skippedCount;
            result.OverwrittenCount = overwrittenCount;
            result.Errors = errors;

            return result;
        }
        catch (JsonException ex)
        {
            throw new Exception($"文件格式错误，不是有效的 JSON 格式: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"导入数据字典失败: {ex.Message}", ex);
        }
    }

    public bool ValidateVanFile(string fileContent)
    {
        if (string.IsNullOrEmpty(fileContent))
            return false;

        try
        {
            var exportData = JsonSerializer.Deserialize<DataDictionaryExportData>(fileContent, _jsonOptions);
            return exportData != null &&
                   !string.IsNullOrEmpty(exportData.Version) &&
                   exportData.Dictionaries != null;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
    }
}