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
                    if (!File.Exists(filePath))
                    {
                        throw new FileNotFoundException($"文件不存在: {filePath}");
                    }

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
                html.AppendLine("td { border: 1px solid #ccc; padding: 4px 8px; white-space: pre-wrap; }");
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
                html.AppendLine("td { border: 1px solid #ccc; padding: 4px 8px; white-space: pre-wrap; }");
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
                return string.Empty;
            }
        }
        
        /// <summary>
        /// 解析 HTML 格式的考勤数据 - 自动识别样表类型（更新版）
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

                // 检测样表类型
                var tableType = DetectTableType(rows, doc);
        
                switch (tableType)
                {
                    case TableType.OldStyleKeyValue:
                        records = ParseOldStyleAttendance(rows, doc);
                        break;
                    case TableType.Template1:
                        records = ParseAttendanceRecordStyle1(rows);
                        break;
                    case TableType.Template2:
                        records = ParseAttendanceRecordStyle2(rows);
                        break;
                    case TableType.Template3:
                        records = ParseAttendanceRecordStyle3(rows);
                        break;
                    case TableType.Template4:
                        records = ParseAttendanceRecordStyle4(rows);
                        break;
                    case TableType.Template5:
                        records = ParseAttendanceRecordStyle5(rows);
                        break;
                    default:
                        records = ParseOldStyleAttendance(rows, doc);
                        break;
                }

                return records;
            }
            catch (Exception ex)
            {
                throw new Exception($"解析 HTML 考勤数据失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 检测表格类型（更新版）
        /// </summary>
        private TableType DetectTableType(HtmlNodeCollection rows, HtmlDocument doc)
        {
            // 获取前几行文本内容用于检测
            var headerText = new StringBuilder();
            for (int i = 0; i < Math.Min(10, rows.Count); i++)
            {
                headerText.Append(rows[i].InnerText);
            }
            var fullText = headerText.ToString();
            
            // 检测旧样式
            if (fullText.Contains("工号：") && fullText.Contains("姓名：") && fullText.Contains("部门："))
            {
                return TableType.OldStyleKeyValue;
            }
            // 检测样表1
            if (fullText.Contains("刷卡记录表") && fullText.Contains("考勤日期"))
            {
                return TableType.Template1;
            }
            // 检测样表2
            if (fullText.Contains("考勤记录表") && fullText.Contains("建表时间"))
            {
                return TableType.Template2;
            }
            // 检测样表3
            if (fullText.Contains("月度汇总表") || fullText.Contains("应出勤天数"))
            {
                return TableType.Template3;
            }            
            // 检测样表4
            if (fullText.Contains("考勤记录表("))
            {
                return TableType.Template4;
            }
            // 检测样表5
            if (fullText.Contains("对应时段") && fullText.Contains("上班时间") && 
                fullText.Contains("签到时间") && fullText.Contains("签退时间"))
            {
                return TableType.Template5;
            }
            
            // 检测数据行首列是否有工号（数字）- 用于样表1和样表2
            if (rows.Count > 5)
            {
                for (int i = 0; i < Math.Min(10, rows.Count); i++)
                {
                    var rowText = rows[i].InnerText;
                    if (rowText.Contains("工号") && rowText.Contains("姓名") && rowText.Contains("部门"))
                    {
                        if (i + 1 < rows.Count)
                        {
                            var nextRowText = rows[i + 1].InnerText;
                            if (nextRowText.Contains("一") || nextRowText.Contains("二") || 
                                nextRowText.Contains("三") || nextRowText.Contains("四") ||
                                nextRowText.Contains("五") || nextRowText.Contains("六") ||
                                nextRowText.Contains("日"))
                            {
                                return TableType.Template1;
                            }
                        }
                        return TableType.Template1;
                    }
                }
            }
        
            // 默认返回旧样式
            return TableType.OldStyleKeyValue;
        }

        /// <summary>
        /// 解析旧样式（键值对格式）
        /// </summary>
        private List<AttendanceRecord> ParseOldStyleAttendance(HtmlNodeCollection rows, HtmlDocument doc)
        {
            var records = new List<AttendanceRecord>();
            
            try
            {
                DateTime monthYear = ParseMonthFromHtml(doc);

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
                        if (employeeId == "7777" && string.IsNullOrEmpty(employeeName))
                        {
                            continue;
                        }
                        
                        employeeRows.Add((rowIndex, employeeId, employeeName, department));
                    }
                }

                foreach (var emp in employeeRows)
                {
                    for (int offset = 1; offset <= 8; offset++)
                    {
                        int dateRowIndex = emp.RowIndex + offset;
                        
                        if (dateRowIndex >= rows.Count) break;
                        
                        var dateRow = rows[dateRowIndex];
                        var dateCells = dateRow.SelectNodes(".//td");
                        
                        if (dateCells == null || dateCells.Count < 10) continue;
                        
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
            }
            catch (Exception ex)
            {
                throw new Exception($"解析旧样式考勤数据失败: {ex.Message}", ex);
            }

            return records;
        }

        /// <summary>
        /// 解析样表1：刷卡记录表
        /// </summary>
        private List<AttendanceRecord> ParseAttendanceRecordStyle1(HtmlNodeCollection rows)
        {
            var records = new List<AttendanceRecord>();

            try
            {
                if (rows.Count < 5) return records;

                // 解析考勤日期
                DateTime monthYear = DateTime.Now;
                for (int i = 0; i < Math.Min(5, rows.Count); i++)
                {
                    var rowText = rows[i].InnerText;
                    var match = Regex.Match(rowText, @"(\d{4})/(\d{2})/(\d{2})\s*~\s*(\d{4})/(\d{2})/(\d{2})");
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int year) &&
                            int.TryParse(match.Groups[2].Value, out int month))
                        {
                            monthYear = new DateTime(year, month, 1);
                            break;
                        }
                    }
                }

                // 查找表头行（包含"工号"、"姓名"、"部门"）
                int headerRowIndex = -1;
                for (int i = 0; i < rows.Count; i++)
                {
                    var rowText = rows[i].InnerText;
                    if (rowText.Contains("工号") && rowText.Contains("姓名") && rowText.Contains("部门"))
                    {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex == -1) return records;

                // 获取表头单元格
                var headerCells = rows[headerRowIndex].SelectNodes(".//td");
                if (headerCells == null || headerCells.Count < 3) return records;

                // 查找员工信息列的索引
                int employeeIdCol = -1, employeeNameCol = -1, departmentCol = -1;
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    if (text == "工号")
                        employeeIdCol = col;
                    else if (text == "姓名")
                        employeeNameCol = col;
                    else if (text == "部门")
                        departmentCol = col;
                }

                if (employeeIdCol == -1 || employeeNameCol == -1) return records;

                // 获取日期列映射 - 从表头行提取日期数字
                var dateColumnMap = new Dictionary<int, int>(); // columnIndex -> day
                for (int col = 0; col < headerCells.Count; col++)
                {
                    // 跳过工号、姓名、部门列
                    if (col == employeeIdCol || col == employeeNameCol || col == departmentCol) continue;
                    
                    var text = headerCells[col].InnerText.Trim();
                    if (int.TryParse(text, out int day) && day >= 1 && day <= 31)
                    {
                        dateColumnMap[col] = day;
                    }
                }

                if (dateColumnMap.Count == 0) return records;

                // 遍历数据行（从表头后2行开始，因为有一行星期行）
                for (int rowIdx = headerRowIndex + 2; rowIdx < rows.Count; rowIdx++)
                {
                    var row = rows[rowIdx];
                    var cells = row.SelectNodes(".//td");
                    if (cells == null || cells.Count <= Math.Max(employeeIdCol, employeeNameCol)) continue;

                    var employeeId = cells[employeeIdCol].InnerText.Trim();
                    var employeeName = cells[employeeNameCol].InnerText.Trim();

                    // 如果工号和姓名都为空，跳过
                    if (string.IsNullOrEmpty(employeeId) && string.IsNullOrEmpty(employeeName)) continue;

                    // 如果工号为空但姓名不为空，使用"EMP_"前缀加姓名作为工号
                    if (string.IsNullOrEmpty(employeeId) && !string.IsNullOrEmpty(employeeName))
                    {
                        employeeId = $"EMP_{employeeName}";
                    }

                    string department = departmentCol >= 0 && departmentCol < cells.Count 
                        ? cells[departmentCol].InnerText.Trim() 
                        : "";

                    // 遍历日期列
                    foreach (var kvp in dateColumnMap)
                    {
                        int colIndex = kvp.Key;
                        int day = kvp.Value;

                        if (colIndex >= cells.Count) continue;

                        var cellValue = cells[colIndex].InnerText.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        // 解析打卡时间
                        var times = ParseTimesFromHtmlCell(cellValue);
                        foreach (var time in times)
                        {
                            if (time.HasValue)
                            {
                                records.Add(CreateRecord(
                                    employeeId, employeeName, department,
                                    monthYear, day, time.Value
                                ));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"解析刷卡记录表失败: {ex.Message}", ex);
            }

            return records;
        }

        /// <summary>
        /// 解析样表2
        /// </summary>
        private List<AttendanceRecord> ParseAttendanceRecordStyle2(HtmlNodeCollection rows)
        {
            var records = new List<AttendanceRecord>();

            try
            {
                if (rows.Count < 6) return records;

                // 解析统计日期
                DateTime monthYear = DateTime.Now;
                for (int i = 0; i < Math.Min(5, rows.Count); i++)
                {
                    var rowText = rows[i].InnerText;
                    var match = Regex.Match(rowText, @"统计日期:(\d{4})/(\d{2})/(\d{2})");
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int year) &&
                            int.TryParse(match.Groups[2].Value, out int month))
                        {
                            monthYear = new DateTime(year, month, 1);
                            break;
                        }
                    }
                }

                // 查找表头行（包含工号、姓名、部门）
                int headerRowIndex = -1;
                for (int i = 0; i < rows.Count; i++)
                {
                    var rowText = rows[i].InnerText;
                    if (rowText.Contains("工号") && rowText.Contains("姓名") && rowText.Contains("部门"))
                    {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex == -1) return records;

                var headerCells = rows[headerRowIndex].SelectNodes(".//td");
                if (headerCells == null || headerCells.Count < 3) return records;

                // 获取日期列映射
                var dateColumnMap = new Dictionary<int, int>(); // columnIndex -> day
                
                // 从表头行提取日期
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    if (int.TryParse(text, out int day) && day >= 1 && day <= 31)
                    {
                        // 跳过工号、姓名、部门列
                        if (col > 2) // 假设前3列是工号、姓名、部门
                        {
                            dateColumnMap[col] = day;
                        }
                    }
                }

                // 如果表头行没有日期，尝试下一行
                if (dateColumnMap.Count == 0 && headerRowIndex + 1 < rows.Count)
                {
                    var dateRow = rows[headerRowIndex + 1];
                    var dateCells = dateRow.SelectNodes(".//td");
                    if (dateCells != null)
                    {
                        for (int col = 0; col < dateCells.Count && col < headerCells.Count; col++)
                        {
                            var text = dateCells[col].InnerText.Trim();
                            if (int.TryParse(text, out int day) && day >= 1 && day <= 31)
                            {
                                dateColumnMap[col] = day;
                            }
                        }
                    }
                }

                if (dateColumnMap.Count == 0) return records;

                // 查找员工信息列
                int employeeIdCol = -1, employeeNameCol = -1, departmentCol = -1;
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    if (text == "工号")
                        employeeIdCol = col;
                    else if (text == "姓名")
                        employeeNameCol = col;
                    else if (text == "部门")
                        departmentCol = col;
                }

                if (employeeIdCol == -1 || employeeNameCol == -1) return records;

                // 遍历数据行
                for (int rowIdx = headerRowIndex + 2; rowIdx < rows.Count; rowIdx++)
                {
                    var row = rows[rowIdx];
                    var cells = row.SelectNodes(".//td");
                    if (cells == null || cells.Count <= Math.Max(employeeIdCol, employeeNameCol)) continue;

                    var employeeId = cells[employeeIdCol].InnerText.Trim();
                    var employeeName = cells[employeeNameCol].InnerText.Trim();

                    if (string.IsNullOrEmpty(employeeId) && string.IsNullOrEmpty(employeeName)) continue;
                    
                    if (string.IsNullOrEmpty(employeeId) && !string.IsNullOrEmpty(employeeName))
                    {
                        employeeId = $"EMP_{employeeName}";
                    }

                    string department = departmentCol >= 0 && departmentCol < cells.Count 
                        ? cells[departmentCol].InnerText.Trim() 
                        : "";

                    foreach (var kvp in dateColumnMap)
                    {
                        int colIndex = kvp.Key;
                        int day = kvp.Value;

                        if (colIndex >= cells.Count) continue;

                        var cellValue = cells[colIndex].InnerText.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        var times = ParseTimesFromHtmlCell(cellValue);
                        foreach (var time in times)
                        {
                            if (time.HasValue)
                            {
                                records.Add(CreateRecord(
                                    employeeId, employeeName, department,
                                    monthYear, day, time.Value
                                ));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"解析考勤记录表失败: {ex.Message}", ex);
            }

            return records;
        }


        /// <summary>
        /// 解析样表3
        /// </summary>
        private List<AttendanceRecord> ParseAttendanceRecordStyle3(HtmlNodeCollection rows)
        {
            var records = new List<AttendanceRecord>();

            try
            {
                if (rows.Count < 5) return records;

                // 解析月份（从标题获取）
                DateTime monthYear = DateTime.Now;
                int detectedYear = DateTime.Now.Year;
                int detectedMonth = DateTime.Now.Month;
                
                for (int i = 0; i < Math.Min(3, rows.Count); i++)
                {
                    var rowText = rows[i].InnerText;
                    // 支持多种日期格式
                    var match = Regex.Match(rowText, @"(\d{4})年(\d{1,2})月");
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int year) &&
                            int.TryParse(match.Groups[2].Value, out int month))
                        {
                            monthYear = new DateTime(year, month, 1);
                            detectedYear = year;
                            detectedMonth = month;
                            break;
                        }
                    }
                    
                    // 尝试其他日期格式
                    match = Regex.Match(rowText, @"(\d{4})[/-](\d{1,2})[/-]\d{1,2}");
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int year) &&
                            int.TryParse(match.Groups[2].Value, out int month))
                        {
                            monthYear = new DateTime(year, month, 1);
                            detectedYear = year;
                            detectedMonth = month;
                            break;
                        }
                    }
                }

                // 查找表头行（包含姓名、工号、部门）
                int headerRowIndex = -1;
                for (int i = 0; i < rows.Count; i++)
                {
                    var rowText = rows[i].InnerText;
                    if (rowText.Contains("姓名") && rowText.Contains("工号") && rowText.Contains("部门"))
                    {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex == -1) return records;

                var headerCells = rows[headerRowIndex].SelectNodes(".//td");
                if (headerCells == null) return records;

                // 动态获取日期列映射（支持任意月份）
                var dateColumnMap = new Dictionary<int, int>(); // columnIndex -> day
                
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    
                    // 方法1：匹配 MM-dd 格式（如 08-01, 8-1）
                    var match1 = Regex.Match(text, @"(\d{1,2})-(\d{1,2})");
                    if (match1.Success)
                    {
                        if (int.TryParse(match1.Groups[1].Value, out int month) &&
                            int.TryParse(match1.Groups[2].Value, out int dayValue1) &&
                            month == detectedMonth && dayValue1 >= 1 && dayValue1 <= 31)
                        {
                            dateColumnMap[col] = dayValue1;
                            continue;
                        }
                    }
                    
                    // 方法2：匹配 MM/DD 格式（如 08/01）
                    var match2 = Regex.Match(text, @"(\d{1,2})/(\d{1,2})");
                    if (match2.Success)
                    {
                        if (int.TryParse(match2.Groups[1].Value, out int month) &&
                            int.TryParse(match2.Groups[2].Value, out int dayValue2) &&
                            month == detectedMonth && dayValue2 >= 1 && dayValue2 <= 31)
                        {
                            dateColumnMap[col] = dayValue2;
                            continue;
                        }
                    }
                    
                    // 方法3：匹配预设的日期模式（如 "8月1日", "8月01日"）
                    var match3 = Regex.Match(text, $@"{detectedMonth}月(\d{{1,2}})日");
                    if (match3.Success)
                    {
                        if (int.TryParse(match3.Groups[1].Value, out int dayValue3) && dayValue3 >= 1 && dayValue3 <= 31)
                        {
                            dateColumnMap[col] = dayValue3;
                            continue;
                        }
                    }
                    
                    // 方法4：直接匹配日期数字（仅在明确是日期列的情况下）
                    if (int.TryParse(text, out int dayValue4) && dayValue4 >= 1 && dayValue4 <= 31)
                    {
                        // 检查是否是日期列（通过上下文判断）
                        dateColumnMap[col] = dayValue4;
                    }
                }

                // 如果还没找到日期列，尝试使用动态生成的模式列表
                if (dateColumnMap.Count == 0)
                {
                    var datePatterns = GenerateDatePatternsForMonth(detectedYear, detectedMonth);
                    for (int col = 0; col < headerCells.Count; col++)
                    {
                        var text = headerCells[col].InnerText.Trim();
                        foreach (var pattern in datePatterns)
                        {
                            if (text.Contains(pattern))
                            {
                                var dayMatch = Regex.Match(pattern, @"(\d{2})$");
                                if (dayMatch.Success && int.TryParse(dayMatch.Groups[1].Value, out int dayValue5))
                                {
                                    dateColumnMap[col] = dayValue5;
                                    break;
                                }
                            }
                        }
                    }
                }

                if (dateColumnMap.Count == 0) return records;

                // 查找员工信息列
                int nameCol = -1, employeeIdCol = -1, departmentCol = -1;
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    if (text == "姓名")
                        nameCol = col;
                    else if (text == "工号")
                        employeeIdCol = col;
                    else if (text == "部门")
                        departmentCol = col;
                }

                if (nameCol == -1 || employeeIdCol == -1) return records;

                // 遍历数据行
                int rowIdx = headerRowIndex + 1;
                while (rowIdx < rows.Count)
                {
                    var row = rows[rowIdx];
                    var cells = row.SelectNodes(".//td");
                    if (cells == null || cells.Count < 3) { rowIdx++; continue; }

                    var name = cells[nameCol].InnerText.Trim();
                    
                    if (string.IsNullOrEmpty(name))
                    {
                        rowIdx++;
                        continue;
                    }

                    var employeeId = employeeIdCol >= 0 && employeeIdCol < cells.Count 
                        ? cells[employeeIdCol].InnerText.Trim() 
                        : "";
                    
                    if (string.IsNullOrEmpty(employeeId) || employeeId == "-")
                    {
                        employeeId = $"EMP_{name}";
                    }

                    var department = departmentCol >= 0 && departmentCol < cells.Count 
                        ? cells[departmentCol].InnerText.Trim() 
                        : "";

                    // 收集该员工的所有打卡数据（可能跨多行）
                    var employeeDataRows = new List<HtmlNode> { row };
                    int nextRowIdx = rowIdx + 1;
                    
                    while (nextRowIdx < rows.Count)
                    {
                        var nextRow = rows[nextRowIdx];
                        var nextCells = nextRow.SelectNodes(".//td");
                        if (nextCells == null || nextCells.Count < 3) { nextRowIdx++; continue; }
                        
                        var nextName = nextCells[nameCol].InnerText.Trim();
                        if (string.IsNullOrEmpty(nextName))
                        {
                            employeeDataRows.Add(nextRow);
                            nextRowIdx++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    // 合并同一员工的打卡数据
                    var mergedCellValues = new Dictionary<int, List<string>>();
                    foreach (var dataRow in employeeDataRows)
                    {
                        var dataCells = dataRow.SelectNodes(".//td");
                        if (dataCells == null) continue;
                        
                        for (int col = 0; col < dataCells.Count; col++)
                        {
                            var value = dataCells[col].InnerText.Trim();
                            if (!string.IsNullOrEmpty(value))
                            {
                                if (!mergedCellValues.ContainsKey(col))
                                    mergedCellValues[col] = new List<string>();
                                mergedCellValues[col].Add(value);
                            }
                        }
                    }

                    // 解析打卡数据
                    foreach (var kvp in dateColumnMap)
                    {
                        int colIndex = kvp.Key;
                        int dayValue = kvp.Value;

                        if (!mergedCellValues.ContainsKey(colIndex)) continue;

                        var cellValues = mergedCellValues[colIndex];
                        foreach (var cellValue in cellValues)
                        {
                            if (IsNonAttendanceText(cellValue)) continue;

                            var times = ParseTimesFromHtmlCell(cellValue);
                            foreach (var time in times)
                            {
                                if (time.HasValue)
                                {
                                    records.Add(CreateRecord(
                                        employeeId, name, department,
                                        monthYear, dayValue, time.Value
                                    ));
                                }
                            }
                        }
                    }

                    rowIdx = nextRowIdx;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"解析月度汇总表失败: {ex.Message}", ex);
            }

            return records;
        }

        /// <summary>
        /// 动态生成指定月份的日期模式列表
        /// </summary>
        private List<string> GenerateDatePatternsForMonth(int year, int month)
        {
            var patterns = new List<string>();
            int daysInMonth = DateTime.DaysInMonth(year, month);
            
            // 生成各种格式的日期模式
            for (int day = 1; day <= daysInMonth; day++)
            {
                // MM-dd 格式
                patterns.Add($"{month:00}-{day:00}");
                // M-d 格式（无前导零）
                patterns.Add($"{month}-{day}");
                // MM/dd 格式
                patterns.Add($"{month:00}/{day:00}");
                // M/d 格式（无前导零）
                patterns.Add($"{month}/{day}");
                // 月日 格式（中文）
                patterns.Add($"{month}月{day}日");
                patterns.Add($"{month}月{day:00}日");
            }
            
            return patterns;
        }

        /// <summary>
        /// 解析样表4
        /// </summary>
        private List<AttendanceRecord> ParseAttendanceRecordStyle4(HtmlNodeCollection rows)
        {
            var records = new List<AttendanceRecord>();

            try
            {
                if (rows.Count < 3) return records;

                // 解析考勤日期范围
                DateTime monthYear = DateTime.Now;
                for (int i = 0; i < Math.Min(3, rows.Count); i++)
                {
                    var rowText = rows[i].InnerText;
                    var match = Regex.Match(rowText, @"考勤记录表\((\d{4})(\d{2})(\d{2})-(\d{4})(\d{2})(\d{2})\)");
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int year) &&
                            int.TryParse(match.Groups[2].Value, out int month))
                        {
                            monthYear = new DateTime(year, month, 1);
                            break;
                        }
                    }
                }

                // 查找表头行（包含工号、姓名、部门）
                int headerRowIndex = -1;
                for (int i = 0; i < rows.Count; i++)
                {
                    var rowText = rows[i].InnerText;
                    if (rowText.Contains("工号") && rowText.Contains("姓名") && rowText.Contains("部门"))
                    {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex == -1) return records;

                var headerCells = rows[headerRowIndex].SelectNodes(".//td");
                if (headerCells == null || headerCells.Count < 3) return records;

                // 获取日期列映射（从表头提取日期，如"7-1 周三"、"8-1 周三"）
                var dateColumnMap = new Dictionary<int, int>(); // columnIndex -> day
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    // 匹配任意月份的 MM-dd 格式
                    var match = Regex.Match(text, @"(\d{1,2})-(\d{1,2})");
                    if (match.Success)
                    {
                        // 注意：这里不限制月份，因为样表4的日期可能跨月
                        if (int.TryParse(match.Groups[2].Value, out int dayNumber) && dayNumber >= 1 && dayNumber <= 31)
                        {
                            dateColumnMap[col] = dayNumber;
                        }
                    }
                }

                if (dateColumnMap.Count == 0) return records;

                // 查找员工信息列
                int employeeIdCol = -1, employeeNameCol = -1, departmentCol = -1;
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    if (text == "工号")
                        employeeIdCol = col;
                    else if (text == "姓名")
                        employeeNameCol = col;
                    else if (text == "部门")
                        departmentCol = col;
                }

                if (employeeIdCol == -1 || employeeNameCol == -1) return records;

                // 遍历数据行
                for (int rowIdx = headerRowIndex + 1; rowIdx < rows.Count; rowIdx++)
                {
                    var row = rows[rowIdx];
                    var cells = row.SelectNodes(".//td");
                    if (cells == null || cells.Count <= Math.Max(employeeIdCol, employeeNameCol)) continue;

                    var employeeId = cells[employeeIdCol].InnerText.Trim();
                    var employeeName = cells[employeeNameCol].InnerText.Trim();

                    if (string.IsNullOrEmpty(employeeId) && string.IsNullOrEmpty(employeeName)) continue;
                    
                    if (string.IsNullOrEmpty(employeeId) && !string.IsNullOrEmpty(employeeName))
                    {
                        employeeId = $"EMP_{employeeName}";
                    }

                    string department = departmentCol >= 0 && departmentCol < cells.Count 
                        ? cells[departmentCol].InnerText.Trim() 
                        : "";

                    foreach (var kvp in dateColumnMap)
                    {
                        int colIndex = kvp.Key;
                        int dayNumber = kvp.Value;

                        if (colIndex >= cells.Count) continue;

                        var cellValue = cells[colIndex].InnerText.Trim();
                        if (string.IsNullOrEmpty(cellValue)) continue;

                        var times = ParseTimesFromHtmlCell(cellValue);
                        foreach (var time in times)
                        {
                            if (time.HasValue)
                            {
                                records.Add(CreateRecord(
                                    employeeId, employeeName, department,
                                    monthYear, dayNumber, time.Value
                                ));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"解析考勤记录表V2失败: {ex.Message}", ex);
            }

            return records;
        }

        /// <summary>
        /// 从HTML单元格中解析时间列表（增强版）
        /// </summary>
        private List<TimeSpan?> ParseTimesFromHtmlCell(string cellValue)
        {
            var result = new List<TimeSpan?>();
            
            if (string.IsNullOrWhiteSpace(cellValue))
                return result;

            // 处理多种换行方式
            var text = cellValue
                .Replace("&lt;br/&gt;", "\n")
                .Replace("&lt;br&gt;", "\n")
                .Replace("<br/>", "\n")
                .Replace("<br>", "\n")
                .Replace("&lt;br /&gt;", "\n")
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

            // 按换行符分割
            var parts = text.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (string.IsNullOrEmpty(trimmed))
                    continue;

                // 尝试解析时间 (支持 HH:mm 和 HH:mm:ss 格式)
                if (TimeSpan.TryParse(trimmed, out var time))
                {
                    // 只保留合理的时间范围（5:00 - 23:59）
                    if (time >= TimeSpan.FromHours(5) && time < TimeSpan.FromHours(24))
                    {
                        result.Add(time);
                    }
                }
                else
                {
                    // 尝试用正则匹配时间格式 HH:mm 或 HH:mm:ss
                    var match = Regex.Match(trimmed, @"(\d{1,2}):(\d{2})(?::(\d{2}))?");
                    if (match.Success)
                    {
                        if (int.TryParse(match.Groups[1].Value, out int hours) &&
                            int.TryParse(match.Groups[2].Value, out int minutes))
                        {
                            int seconds = match.Groups[3].Success ? int.Parse(match.Groups[3].Value) : 0;
                            if (hours >= 5 && hours < 24 && minutes >= 0 && minutes < 60)
                            {
                                result.Add(new TimeSpan(hours, minutes, seconds));
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 判断是否为非打卡文字（如"外出"、"请假"等）
        /// </summary>
        private bool IsNonAttendanceText(string text)
        {
            if (string.IsNullOrEmpty(text)) return true;
            
            var nonAttendanceKeywords = new[] { "外出", "请假", "出差", "休息", "-", "—", "缺勤", "旷工" };
            foreach (var keyword in nonAttendanceKeywords)
            {
                if (text.Contains(keyword)) return true;
            }
            
            return false;
        }

        /// <summary>
        /// 从 HTML 行中提取打卡记录（旧样式专用）
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
            
            for (int i = 0; i < dateNumbers.Count; i++)
            {
                int colIndex = firstDateColumn + i;
                int dayOfMonth = dateNumbers[i];
                
                if (dayOfMonth < 1 || dayOfMonth > 31) continue;

                string morningValue = "";
                if (morningCells != null && colIndex < morningCells.Count)
                {
                    morningValue = morningCells[colIndex].InnerText.Trim();
                    morningValue = morningValue.Replace("&lt;br/&gt;", "\n").Replace("<br/>", "\n").Replace("<br>", "\n");
                }

                string afternoonValue = "";
                if (afternoonCells != null && colIndex < afternoonCells.Count)
                {
                    afternoonValue = afternoonCells[colIndex].InnerText.Trim();
                    afternoonValue = afternoonValue.Replace("&lt;br/&gt;", "\n").Replace("<br/>", "\n").Replace("<br>", "\n");
                }

                if (!string.IsNullOrEmpty(morningValue))
                {
                    var morningTimes = ParseTimesFromHtmlCell(morningValue);
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

                if (!string.IsNullOrEmpty(afternoonValue))
                {
                    var afternoonTimes = ParseTimesFromHtmlCell(afternoonValue);
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
        /// 从 HTML 中解析考勤月份（旧样式专用）
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

            return new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        }

        /// <summary>
        /// 表格类型枚举
        /// </summary>
        private enum TableType
        {
            OldStyleKeyValue,       // 旧样式  市局
            Template1,              // 样表1：施甸县局
            Template2,              // 样表2：隆阳区局
            Template3,              // 样表3：腾冲市局
            Template4,              // 样表4：昌宁县局
            Template5               // 样表5：龙陵县局
        }
        
        #region 样表5 支持 龙陵局样表
        
        /// <summary>
        /// 解析样表5
        /// 只读取签到时间和签退时间作为打卡记录，上班时间和下班时间是规定时间，不作为打卡记录
        /// </summary>
        private List<AttendanceRecord> ParseAttendanceRecordStyle5(HtmlNodeCollection rows)
        {
            var records = new List<AttendanceRecord>();
        
            try
            {
                if (rows.Count < 5) return records;
        
                // 查找表头行（包含"姓名"、"日期"、"对应时段"、"上班时间"等）
                int headerRowIndex = -1;
                for (int i = 0; i < rows.Count; i++)
                {
                    var rowText = rows[i].InnerText;
                    if (rowText.Contains("姓名") && rowText.Contains("日期") && 
                        rowText.Contains("对应时段") && rowText.Contains("签到时间"))
                    {
                        headerRowIndex = i;
                        break;
                    }
                }
        
                if (headerRowIndex == -1) return records;
        
                var headerCells = rows[headerRowIndex].SelectNodes(".//td");
                if (headerCells == null || headerCells.Count < 5) return records;
        
                // 获取各列索引
                int nameCol = -1, dateCol = -1, periodCol = -1, 
                    signInCol = -1, signOutCol = -1, deptCol = -1;
        
                for (int col = 0; col < headerCells.Count; col++)
                {
                    var text = headerCells[col].InnerText.Trim();
                    if (text == "姓名")
                        nameCol = col;
                    else if (text == "日期")
                        dateCol = col;
                    else if (text == "对应时段")
                        periodCol = col;
                    else if (text == "签到时间")
                        signInCol = col;
                    else if (text == "签退时间")
                        signOutCol = col;
                    else if (text == "部门")
                        deptCol = col;
                }
        
                if (nameCol == -1 || dateCol == -1) return records;
        
                // 遍历数据行
                for (int rowIdx = headerRowIndex + 1; rowIdx < rows.Count; rowIdx++)
                {
                    var row = rows[rowIdx];
                    var cells = row.SelectNodes(".//td");
                    if (cells == null || cells.Count <= Math.Max(nameCol, dateCol)) continue;
        
                    var employeeName = cells[nameCol].InnerText.Trim();
                    var dateText = cells[dateCol].InnerText.Trim();
                    var period = periodCol >= 0 && periodCol < cells.Count 
                        ? cells[periodCol].InnerText.Trim() 
                        : "";
                    var signInText = signInCol >= 0 && signInCol < cells.Count 
                        ? cells[signInCol].InnerText.Trim() 
                        : "";
                    var signOutText = signOutCol >= 0 && signOutCol < cells.Count 
                        ? cells[signOutCol].InnerText.Trim() 
                        : "";
                    var department = deptCol >= 0 && deptCol < cells.Count 
                        ? cells[deptCol].InnerText.Trim() 
                        : "";
        
                    // 跳过空行或无效数据
                    if (string.IsNullOrEmpty(employeeName)) continue;
        
                    // 解析日期
                    DateTime? recordDate = ParseDateFromText(dateText);
                    if (!recordDate.HasValue) continue;
        
                    // 生成工号（使用姓名）
                    var employeeId = GenerateEmployeeId(employeeName);
        
                    // 只记录签到时间（不为空且有效）
                    if (!string.IsNullOrEmpty(signInText))
                    {
                        var signInTime = ParseTimeFromText(signInText);
                        if (signInTime.HasValue && IsValidCheckTime(signInTime.Value))
                        {
                            records.Add(CreateRecord(
                                employeeId, employeeName, department,
                                recordDate.Value, signInTime.Value
                            ));
                        }
                    }
        
                    // 只记录签退时间（不为空且有效）
                    if (!string.IsNullOrEmpty(signOutText))
                    {
                        var signOutTime = ParseTimeFromText(signOutText);
                        if (signOutTime.HasValue && IsValidCheckTime(signOutTime.Value))
                        {
                            records.Add(CreateRecord(
                                employeeId, employeeName, department,
                                recordDate.Value, signOutTime.Value
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"解析考勤明细表失败: {ex.Message}", ex);
            }
        
            return records;
        }
        
        /// <summary>
        /// 验证打卡时间是否有效（排除明显不合理的时间）
        /// </summary>
        private bool IsValidCheckTime(TimeSpan time)
        {
            // 合理的时间范围：5:00 - 23:00
            // 排除 00:00-04:59（通常是无效数据）和 23:00以后（通常是系统错误）
            if (time >= TimeSpan.FromHours(5) && time < TimeSpan.FromHours(23))
            {
                return true;
            }
            
            // 特殊情况：如果是23:30之前，可能算加班，但这里保守处理
            // 可以调整阈值
            if (time >= TimeSpan.FromHours(23) && time <= TimeSpan.FromHours(23))
            {
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// 从文本中解析日期（支持多种格式）
        /// </summary>
        private DateTime? ParseDateFromText(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
        
            text = text.Trim();
            
            // 支持格式: 2026/6/1, 2026-6-1, 2026年6月1日
            var formats = new[] { 
                "yyyy/M/d", "yyyy-M-d", "yyyy年M月d日",
                "yyyy/MM/dd", "yyyy-MM-dd", "yyyy年MM月dd日"
            };
            
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(text, format, 
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, 
                    out var date))
                {
                    return date;
                }
            }
        
            // 尝试通用解析
            if (DateTime.TryParse(text, out var result))
            {
                return result;
            }
        
            return null;
        }
        
        /// <summary>
        /// 从文本中解析时间
        /// </summary>
        private TimeSpan? ParseTimeFromText(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
        
            text = text.Trim();
            
            // 支持格式: 07:30, 7:30, 07:30:00
            if (TimeSpan.TryParse(text, out var time))
            {
                if (time >= TimeSpan.FromHours(0) && time < TimeSpan.FromHours(24))
                {
                    return time;
                }
            }
        
            // 尝试正则匹配
            var match = Regex.Match(text, @"(\d{1,2}):(\d{2})(?::(\d{2}))?");
            if (match.Success)
            {
                if (int.TryParse(match.Groups[1].Value, out int hours) &&
                    int.TryParse(match.Groups[2].Value, out int minutes))
                {
                    int seconds = match.Groups[3].Success ? int.Parse(match.Groups[3].Value) : 0;
                    if (hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60)
                    {
                        return new TimeSpan(hours, minutes, seconds);
                    }
                }
            }
        
            return null;
        }
        
        /// <summary>
        /// 根据姓名生成工号
        /// </summary>
        private string GenerateEmployeeId(string name)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            
            // 使用姓名拼音首字母 + 随机数，或者使用固定格式
            // 这里简单使用 "EMP_" + 姓名
            return $"EMP_{name}";
        }
        
        /// <summary>
        /// 创建打卡记录（带日期）
        /// </summary>
        private AttendanceRecord CreateRecord(
            string employeeId,
            string employeeName,
            string department,
            DateTime date,
            TimeSpan time)
        {
            var checkTime = new DateTime(date.Year, date.Month, date.Day, 
                time.Hours, time.Minutes, time.Seconds);
        
            return new AttendanceRecord
            {
                EmployeeId = employeeId,
                EmployeeName = employeeName,
                Department = department,
                CheckTime = checkTime,
                DayOfMonth = date.Day,
                CreatedAt = DateTime.Now
            };
        }
        
        #endregion
    }
}