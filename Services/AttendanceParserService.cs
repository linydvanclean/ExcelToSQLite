using ExcelToSQLite.Models;
using HtmlAgilityPack;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToSQLite.Services
{
    public class AttendanceParserService
    {
        public AttendanceParserService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<List<AttendanceRecord>> ParseAttendanceAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                var records = new List<AttendanceRecord>();
                
                try
                {
                    // 检查文件是否存在
                    if (!File.Exists(filePath))
                    {
                        throw new FileNotFoundException($"文件不存在: {filePath}");
                    }

                    // 根据文件扩展名选择解析方式
                    string extension = Path.GetExtension(filePath).ToLower();
                    string htmlContent;
                    
                    if (extension == ".xlsx")
                    {
                        htmlContent = ConvertXlsxToHtml(filePath);
                    }
                    else if (extension == ".xls")
                    {
                        htmlContent = ConvertXlsToHtml(filePath);
                    }
                    else
                    {
                        throw new NotSupportedException($"不支持的文件格式: {extension}。请使用 .xlsx 或 .xls 格式。");
                    }
                    
                    if (string.IsNullOrEmpty(htmlContent))
                    {
                        return records;
                    }

                    records = ParseHtmlAttendance(htmlContent);
                }
                catch (Exception ex)
                {
                    throw new Exception($"解析考勤文件失败: {ex.Message}", ex);
                }

                return records;
            });
        }

        /// <summary>
        /// 转换 .xlsx 文件为 HTML (使用 EPPlus)
        /// </summary>
        private string ConvertXlsxToHtml(string filePath)
        {
            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet == null || worksheet.Dimension == null)
                {
                    return string.Empty;
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html><head><meta charset='UTF-8'>");
                html.AppendLine("<style>");
                html.AppendLine("table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; }");
                html.AppendLine("td { border: 1px solid #ccc; padding: 4px 8px; }");
                html.AppendLine("</style>");
                html.AppendLine("</head><body>");
                html.AppendLine($"<p><strong>文件:</strong> {Path.GetFileName(filePath)}</p>");
                html.AppendLine($"<p><strong>总行数:</strong> {rowCount}, <strong>总列数:</strong> {colCount}</p>");
                html.AppendLine("<table>");

                for (int row = 1; row <= rowCount; row++)
                {
                    html.AppendLine("<tr>");
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        var value = cell.Value?.ToString()?.Trim() ?? string.Empty;
                        value = value.Replace("\n", "<br/>").Replace("\r", "");
                        html.AppendLine($"<td>{System.Net.WebUtility.HtmlEncode(value)}</td>");
                    }
                    html.AppendLine("</tr>");
                }

                html.AppendLine("</table>");
                html.AppendLine("</body></html>");
                return html.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"转换 .xlsx 文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 转换 .xls 文件为 HTML (使用 NPOI)
        /// </summary>
        private string ConvertXlsToHtml(string filePath)
        {
            try
            {
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                var workbook = new HSSFWorkbook(fs);
                var sheet = workbook.GetSheetAt(0);

                if (sheet == null)
                {
                    return string.Empty;
                }

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

                var html = new StringBuilder();
                html.AppendLine("<!DOCTYPE html>");
                html.AppendLine("<html><head><meta charset='UTF-8'>");
                html.AppendLine("<style>");
                html.AppendLine("table { border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px; }");
                html.AppendLine("td { border: 1px solid #ccc; padding: 4px 8px; }");
                html.AppendLine("</style>");
                html.AppendLine("</head><body>");
                html.AppendLine($"<p><strong>文件:</strong> {Path.GetFileName(filePath)}</p>");
                html.AppendLine($"<p><strong>总行数:</strong> {rowCount}, <strong>总列数:</strong> {colCount}</p>");
                html.AppendLine("<table>");

                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    html.AppendLine("<tr>");
                    
                    for (int colIndex = 0; colIndex < colCount; colIndex++)
                    {
                        var cell = row?.GetCell(colIndex);
                        var value = GetCellValue(cell)?.Trim() ?? string.Empty;
                        
                        
                        value = value.Replace("\n", "<br/>").Replace("\r", "");
                        html.AppendLine($"<td>{System.Net.WebUtility.HtmlEncode(value)}</td>");
                    }
                    
                    html.AppendLine("</tr>");
                }

                html.AppendLine("</table>");
                html.AppendLine("</body></html>");
                return html.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"转换 .xls 文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 获取 NPOI 单元格的值
        /// </summary>
        private string GetCellValue(ICell? cell)
        {
            if (cell == null) return string.Empty;

            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            // 修正：处理 DateTime? 类型
                            var dateValue = cell.DateCellValue;
                            if (dateValue.HasValue)
                            {
                                return dateValue.Value.ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            return cell.NumericCellValue.ToString();
                        }
                        return cell.NumericCellValue.ToString();
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Formula:
                        try
                        {
                            return cell.StringCellValue;
                        }
                        catch
                        {
                            return cell.NumericCellValue.ToString();
                        }
                    case CellType.Blank:
                        return string.Empty;
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                // 如果获取失败，返回空字符串
                return string.Empty;
            }
        }

        /// <summary>
        /// 解析 HTML 格式的考勤数据
        /// </summary>
        private List<AttendanceRecord> ParseHtmlAttendance(string htmlContent)
        {
            var records = new List<AttendanceRecord>();
            
            try
            {
                var doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);

                var table = doc.DocumentNode.SelectSingleNode("//table");
                if (table == null)
                {
                    return records;
                }

                var rows = table.SelectNodes(".//tr");
                if (rows == null || rows.Count == 0)
                {
                    return records;
                }

                DateTime monthYear = ParseMonthFromHtml(doc);

                // 第一步：收集所有员工信息行
                var employeeRows = new List<(int RowIndex, string EmployeeId, string Name, string Department)>();
                
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var currentRow = rows[rowIndex];
                    var cells = currentRow.SelectNodes(".//td");
                    
                    if (cells == null || cells.Count < 5) continue;

                    string employeeId = "";
                    string employeeName = "";
                    string department = "";

                    for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                    {
                        var text = cells[cellIndex].InnerText.Trim();
                        
                        if (text == "工号：")
                        {
                            if (cellIndex + 1 < cells.Count)
                            {
                                employeeId = cells[cellIndex + 1].InnerText.Trim();
                            }
                        }
                        else if (text == "姓名：")
                        {
                            if (cellIndex + 1 < cells.Count)
                            {
                                employeeName = cells[cellIndex + 1].InnerText.Trim();
                            }
                        }
                        else if (text == "部门：")
                        {
                            if (cellIndex + 1 < cells.Count)
                            {
                                department = cells[cellIndex + 1].InnerText.Trim();
                            }
                        }
                        else if (text.Contains("工号：") && string.IsNullOrEmpty(employeeId))
                        {
                            employeeId = text.Replace("工号：", "").Trim();
                        }
                        else if (text.Contains("姓名：") && string.IsNullOrEmpty(employeeName))
                        {
                            employeeName = text.Replace("姓名：", "").Trim();
                        }
                        else if (text.Contains("部门：") && string.IsNullOrEmpty(department))
                        {
                            department = text.Replace("部门：", "").Trim();
                        }
                    }

                    if (!string.IsNullOrEmpty(employeeId) && !string.IsNullOrEmpty(employeeName))
                    {
                        // 跳过无效员工
                        if (employeeId == "7777" && string.IsNullOrEmpty(employeeName))
                        {
                            continue;
                        }
                        
                        employeeRows.Add((rowIndex, employeeId, employeeName, department));
                    }
                }

                // 第二步：解析每个员工的数据
                foreach (var emp in employeeRows)
                {
                    // 从员工行的下一行开始查找数据行（最多向下查找8行）
                    for (int offset = 1; offset <= 8; offset++)
                    {
                        int dateRowIndex = emp.RowIndex + offset;
                        
                        if (dateRowIndex >= rows.Count) break;
                        
                        var dateRow = rows[dateRowIndex];
                        var dateCells = dateRow.SelectNodes(".//td");
                        
                        if (dateCells == null || dateCells.Count < 10) continue;
                        
                        // 检查日期行是否有效，并收集所有日期数字
                        var dateNumbers = new List<int>();
                        int firstDateColumn = -1;
                        
                        for (int colIndex = 0; colIndex < dateCells.Count; colIndex++)
                        {
                            var text = dateCells[colIndex].InnerText.Trim();
                            if (int.TryParse(text, out int day) && day >= 1 && day <= 31)
                            {
                                dateNumbers.Add(day);
                                if (firstDateColumn == -1)
                                {
                                    firstDateColumn = colIndex;
                                }
                            }
                        }
                        
                        // 如果有至少3个日期数字，认为这是日期行
                        if (dateNumbers.Count >= 3)
                        {
                            int morningRowIndex = dateRowIndex + 1;
                            int afternoonRowIndex = dateRowIndex + 2;
                            
                            if (morningRowIndex < rows.Count && afternoonRowIndex < rows.Count)
                            {
                                var morningRow = rows[morningRowIndex];
                                var afternoonRow = rows[afternoonRowIndex];
                                
                                var empRecords = ExtractRecordsFromHtmlRows(
                                    dateRow,
                                    morningRow,
                                    afternoonRow,
                                    emp.EmployeeId,
                                    emp.Name,
                                    emp.Department,
                                    monthYear,
                                    firstDateColumn,
                                    dateNumbers
                                );
                                
                                records.AddRange(empRecords);
                                break;
                            }
                        }
                    }
                }

                return records;
            }
            catch (Exception ex)
            {
                throw new Exception($"解析 HTML 考勤数据失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 从 HTML 行中提取打卡记录
        /// </summary>
        private List<AttendanceRecord> ExtractRecordsFromHtmlRows(
            HtmlNode dateRow,
            HtmlNode morningRow,
            HtmlNode afternoonRow,
            string employeeId,
            string employeeName,
            string department,
            DateTime monthYear,
            int firstDateColumn,
            List<int> dateNumbers)
        {
            var records = new List<AttendanceRecord>();

            var dateCells = dateRow.SelectNodes(".//td");
            var morningCells = morningRow.SelectNodes(".//td");
            var afternoonCells = afternoonRow.SelectNodes(".//td");

            if (dateCells == null || dateNumbers == null || dateNumbers.Count == 0) 
                return records;
            
            // 遍历所有日期数字
            for (int i = 0; i < dateNumbers.Count; i++)
            {
                int colIndex = firstDateColumn + i;
                int dayOfMonth = dateNumbers[i];
                
                // 验证日期是否有效（1-31）
                if (dayOfMonth < 1 || dayOfMonth > 31) continue;

                // 获取上午的打卡数据
                string morningValue = "";
                if (morningCells != null && colIndex < morningCells.Count)
                {
                    morningValue = morningCells[colIndex].InnerText.Trim();
                    morningValue = morningValue.Replace("&lt;br/&gt;", "\n").Replace("<br/>", "\n").Replace("<br>", "\n");
                }

                // 获取下午的打卡数据
                string afternoonValue = "";
                if (afternoonCells != null && colIndex < afternoonCells.Count)
                {
                    afternoonValue = afternoonCells[colIndex].InnerText.Trim();
                    afternoonValue = afternoonValue.Replace("&lt;br/&gt;", "\n").Replace("<br/>", "\n").Replace("<br>", "\n");
                }

                // 解析上午打卡
                if (!string.IsNullOrEmpty(morningValue))
                {
                    var morningTimes = ParseTimesFromString(morningValue);
                    foreach (var time in morningTimes)
                    {
                        if (time.HasValue)
                        {
                            records.Add(CreateRecord(
                                employeeId, employeeName, department,
                                monthYear, dayOfMonth, time.Value
                            ));
                        }
                    }
                }

                // 解析下午打卡
                if (!string.IsNullOrEmpty(afternoonValue))
                {
                    var afternoonTimes = ParseTimesFromString(afternoonValue);
                    foreach (var time in afternoonTimes)
                    {
                        if (time.HasValue)
                        {
                            records.Add(CreateRecord(
                                employeeId, employeeName, department,
                                monthYear, dayOfMonth, time.Value
                            ));
                        }
                    }
                }
            }

            return records;
        }

        /// <summary>
        /// 从字符串中解析时间
        /// </summary>
        private List<TimeSpan?> ParseTimesFromString(string cellValue)
        {
            var result = new List<TimeSpan?>();
            
            if (string.IsNullOrWhiteSpace(cellValue))
                return result;

            // 按换行符分割
            var parts = cellValue.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (string.IsNullOrEmpty(trimmed))
                    continue;

                // 尝试解析时间
                if (TimeSpan.TryParse(trimmed, out var time))
                {
                    // 只保留合理的时间范围（6:00 - 23:59）
                    if (time >= TimeSpan.FromHours(6) && time < TimeSpan.FromHours(24))
                    {
                        result.Add(time);
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 创建打卡记录
        /// </summary>
        private AttendanceRecord CreateRecord(
            string employeeId,
            string employeeName,
            string department,
            DateTime monthYear,
            int dayOfMonth,
            TimeSpan time)
        {
            int year = monthYear.Year;
            int month = monthYear.Month;
            int day = dayOfMonth;

            // 处理跨月（如果日期大于当月天数）
            int daysInMonth = DateTime.DaysInMonth(year, month);
            if (day > daysInMonth)
            {
                if (month == 12)
                {
                    year++;
                    month = 1;
                }
                else
                {
                    month++;
                }
                day = day - daysInMonth;
            }

            var checkTime = new DateTime(year, month, day, time.Hours, time.Minutes, time.Seconds);

            return new AttendanceRecord
            {
                EmployeeId = employeeId,
                EmployeeName = employeeName,
                Department = department,
                CheckTime = checkTime,
                DayOfMonth = dayOfMonth,
                CreatedAt = DateTime.Now
            };
        }

        /// <summary>
        /// 从 HTML 中解析考勤月份
        /// </summary>
        private DateTime ParseMonthFromHtml(HtmlDocument doc)
        {
            try
            {
                var allText = doc.DocumentNode.InnerText;
                var match = Regex.Match(allText, @"考勤日期：(\d{4}-\d{2}-\d{2})");
                if (match.Success && DateTime.TryParse(match.Groups[1].Value, out var date))
                {
                    return date;
                }
            }
            catch
            {
                // 忽略解析异常
            }

            // 默认返回当前月份
            return new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        }
    }
}