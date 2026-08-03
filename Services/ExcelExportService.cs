using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Services;

public class ExcelExportService
{
    public ExcelExportService()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    /// <summary>
    /// 导出数据到 Excel
    /// </summary>
    public async Task<string> ExportToExcelAsync(List<DataRowItem> data, List<string> columnNames, string? fileName = null)
    {
        return await Task.Run(() =>
        {
            if (data == null || data.Count == 0 || columnNames == null || columnNames.Count == 0)
            {
                throw new InvalidOperationException("没有数据可导出");
            }

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");

            // 写入表头
            worksheet.Cells[1, 1].Value = "#";
            for (int i = 0; i < columnNames.Count; i++)
            {
                worksheet.Cells[1, i + 2].Value = columnNames[i];
            }

            // 写入数据
            for (int row = 0; row < data.Count; row++)
            {
                var item = data[row];
                worksheet.Cells[row + 2, 1].Value = item.Index;

                for (int col = 0; col < columnNames.Count; col++)
                {
                    var columnName = columnNames[col];
                    var value = item.Values.TryGetValue(columnName, out var v) ? v : string.Empty;
                    worksheet.Cells[row + 2, col + 2].Value = value;
                }
            }

            // 自动调整列宽
            worksheet.Cells.AutoFitColumns();

            // 保存文件
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = $"export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            }

            var filePath = Path.Combine(Path.GetTempPath(), fileName);
            package.SaveAs(new FileInfo(filePath));

            return filePath;
        });
    }

    /// <summary>
    /// 导出数据到 Excel（指定路径）
    /// </summary>
    public async Task ExportToExcelAsync(List<DataRowItem> data, List<string> columnNames, string filePath, string sheetName = "Sheet1")
    {
        await Task.Run(() =>
        {
            if (data == null || data.Count == 0 || columnNames == null || columnNames.Count == 0)
            {
                throw new InvalidOperationException("没有数据可导出");
            }

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(sheetName);

            // 写入表头
            worksheet.Cells[1, 1].Value = "#";
            for (int i = 0; i < columnNames.Count; i++)
            {
                worksheet.Cells[1, i + 2].Value = columnNames[i];
            }

            // 写入数据
            for (int row = 0; row < data.Count; row++)
            {
                var item = data[row];
                worksheet.Cells[row + 2, 1].Value = item.Index;

                for (int col = 0; col < columnNames.Count; col++)
                {
                    var columnName = columnNames[col];
                    var value = item.Values.TryGetValue(columnName, out var v) ? v : string.Empty;
                    worksheet.Cells[row + 2, col + 2].Value = value;
                }
            }

            // 自动调整列宽
            worksheet.Cells.AutoFitColumns();

            package.SaveAs(new FileInfo(filePath));
        });
    }
    
    
    public async Task<string> ExportDataGridRowToExcelAsync(
        List<DataGridRow> data,
        List<string> columnNames,
        string fileName)
    {
        return await Task.Run(() =>
        {
            var tempFile = System.IO.Path.GetTempFileName();
            tempFile = System.IO.Path.ChangeExtension(tempFile, ".xlsx");

            using var package = new OfficeOpenXml.ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("数据");

            // 写入表头
            worksheet.Cells[1, 1].Value = "#";
            for (int i = 0; i < columnNames.Count; i++)
            {
                worksheet.Cells[1, i + 2].Value = columnNames[i];
            }

            // 写入数据
            for (int i = 0; i < data.Count; i++)
            {
                var row = data[i];
                worksheet.Cells[i + 2, 1].Value = row.Index;

                for (int j = 0; j < columnNames.Count; j++)
                {
                    var columnName = columnNames[j];
                    worksheet.Cells[i + 2, j + 2].Value = row.GetValue(columnName);
                }
            }

            // 自动调整列宽
            worksheet.Cells.AutoFitColumns();

            package.SaveAs(new System.IO.FileInfo(tempFile));
            return tempFile;
        });
    }
}