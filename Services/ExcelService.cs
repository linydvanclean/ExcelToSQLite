using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.Services;

public class ExcelService
{
    public ExcelService()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public async Task<List<List<object>>> ReadExcelAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            try
            {
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"文件不存在: {filePath}");
                }

                // 根据文件扩展名选择解析方式
                string extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".xlsx")
                {
                    return ReadXlsx(filePath);
                }
                else if (extension == ".xls")
                {
                    return ReadXls(filePath);
                }
                else
                {
                    throw new NotSupportedException($"不支持的文件格式: {extension}。请使用 .xlsx 或 .xls 格式。");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取Excel文件失败: {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    /// 读取 .xlsx 文件 (使用 EPPlus)
    /// </summary>
    private List<List<object>> ReadXlsx(string filePath)
    {
        var result = new List<List<object>>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];

        if (worksheet == null || worksheet.Dimension == null)
            return result;

        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;

        // 读取表头（第一行）并处理字段名
        var headers = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            var cell = worksheet.Cells[1, col];
            var rawValue = GetCellValue(cell)?.ToString() ?? string.Empty;  // 传递cell对象
            
            headers.Add(rawValue);
        }

        // 处理字段名冲突和未命名字段
        headers = ProcessHeaders(headers);

        // 将处理后的表头添加到结果中
        var headerRow = headers.Cast<object>().ToList();
        result.Add(headerRow);

        // 读取数据行（从第二行开始）
        for (int row = 2; row <= rowCount; row++)
        {
            var rowData = new List<object>();
            for (int col = 1; col <= colCount; col++)
            {
                // 读取数据行时
                var cell = worksheet.Cells[row, col];
                object value = GetCellValue(cell);  // 传递cell对象
                
                rowData.Add(value);
            }
            result.Add(rowData);
        }

        return result;
    }

    /// <summary>
    /// 读取 .xls 文件 (使用 NPOI)
    /// </summary>
    private List<List<object>> ReadXls(string filePath)
    {
        var result = new List<List<object>>();

        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        var workbook = new HSSFWorkbook(fs);
        var sheet = workbook.GetSheetAt(0);

        if (sheet == null)
            return result;

        int rowCount = sheet.LastRowNum + 1;
        int colCount = 0;

        // 获取最大列数
        for (int i = 0; i < rowCount; i++)
        {
            var row = sheet.GetRow(i);
            if (row != null && row.LastCellNum > colCount)
            {
                colCount = (int)row.LastCellNum;
            }
        }

        // 读取表头（第一行）并处理字段名
        var headers = new List<string>();
        var headerRow = sheet.GetRow(0);
        for (int colIndex = 0; colIndex < colCount; colIndex++)
        {
            var cell = headerRow?.GetCell(colIndex);
            var rawValue = GetNpoiCellValue(cell)?.ToString() ?? string.Empty;
            headers.Add(rawValue);
        }

        // 处理字段名冲突和未命名字段
        headers = ProcessHeaders(headers);

        // 将处理后的表头添加到结果中
        var headerRowData = headers.Cast<object>().ToList();
        result.Add(headerRowData);

        // 读取数据行（从第二行开始）
        for (int rowIndex = 1; rowIndex < rowCount; rowIndex++)
        {
            var rowData = new List<object>();
            var row = sheet.GetRow(rowIndex);
            
            for (int colIndex = 0; colIndex < colCount; colIndex++)
            {
                var cell = row?.GetCell(colIndex);
                object value = GetNpoiCellValue(cell);
                rowData.Add(value);
            }
            
            result.Add(rowData);
        }

        return result;
    }

    /// <summary>
    /// 处理表头字段名：处理保留字段和未命名字段
    /// </summary>
    /// <param name="headers">原始表头列表</param>
    /// <returns>处理后的表头列表</returns>
    private List<string> ProcessHeaders(List<string> headers)
    {
        var processedHeaders = new List<string>();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var unnamedCounter = 1;

        // 第一遍：处理未命名字段并检查保留字段
        for (int i = 0; i < headers.Count; i++)
        {
            var rawValue = headers[i] ?? string.Empty;
            var headerName = rawValue.Trim();

            // 检查是否为未命名字段（空字符串或纯空白）
            if (string.IsNullOrWhiteSpace(headerName))
            {
                // 自动命名为 "未命名_N"
                headerName = $"未命名_{unnamedCounter++}";
            }
            else
            {
                // 检查是否为保留的 "id" 字段（忽略大小写）
                // 注意：这里同时处理了单独出现或者与其他字符组合的情况
                if (headerName.Equals("id", StringComparison.OrdinalIgnoreCase))
                {
                    headerName = "Id_1";
                }
            }

            processedHeaders.Add(headerName);
        }

        // 第二遍：处理重名冲突
        var finalHeaders = new List<string>();
        var nameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var header in processedHeaders)
        {
            if (nameCounts.ContainsKey(header))
            {
                // 如果已存在同名，添加数字后缀
                nameCounts[header]++;
                var newName = $"{header}_{nameCounts[header]}";
                finalHeaders.Add(newName);
            }
            else
            {
                nameCounts[header] = 1;
                finalHeaders.Add(header);
            }
        }

        return finalHeaders;
    }

    /// <summary>
    /// 获取 EPPlus 单元格的值
    /// </summary>
    private object GetCellValue(ExcelRangeBase cell)
    {
        return ExcelHepler.GetDateCellValue(cell);
    }

    /// <summary>
    /// 获取 NPOI 单元格的值
    /// </summary>
    private object GetNpoiCellValue(ICell? cell)
    {
        if (cell == null) return string.Empty;

        try
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue.Trim();
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        var dateValue = cell.DateCellValue;
                        if (dateValue.HasValue)
                        {
                            return dateValue.Value.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        return cell.NumericCellValue.ToString();
                    }
                    
                    double numericValue = cell.NumericCellValue;
                    if (numericValue % 1 == 0)
                    {
                        return Convert.ToInt64(numericValue);
                    }
                    else
                    {
                        return numericValue;
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Formula:
                    try
                    {
                        // 尝试获取公式计算后的字符串值
                        return cell.StringCellValue.Trim();
                    }
                    catch
                    {
                        try
                        {
                            // 如果字符串获取失败，尝试获取数值
                            double formulaValue = cell.NumericCellValue;
                            if (formulaValue % 1 == 0)
                            {
                                return Convert.ToInt64(formulaValue);
                            }
                            else
                            {
                                return formulaValue;
                            }
                        }
                        catch
                        {
                            return cell.ToString() ?? string.Empty;
                        }
                    }
                case CellType.Blank:
                    return string.Empty;
                default:
                    return cell.ToString() ?? string.Empty;
            }
        }
        catch
        {
            return cell.ToString() ?? string.Empty;
        }
    }
}