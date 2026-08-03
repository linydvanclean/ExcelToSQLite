using OfficeOpenXml;
using System;

namespace ExcelToSQLite.Helpers;

public class ExcelHepler
{
    /// <summary>
    /// 读取表格中的日期格式，确保能读取正确
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    public static object GetDateCellValue(ExcelRangeBase cell)
    {
        if (cell == null || cell.Value == null)
        {
            return string.Empty;
        }

        // 首先检查单元格样式是否为日期格式
        bool isDateFormat = false;
        if (cell.Style.Numberformat.Format != null)
        {
            string format = cell.Style.Numberformat.Format.ToLower();
            // 更精确的日期格式判断
            isDateFormat = format.Contains("yyyy") || format.Contains("yy") || 
                          format.Contains("mm") || format.Contains("dd") ||
                          format.Contains("m月") || format.Contains("d日") ||
                          format.Contains("h") || format.Contains("时") ||
                          format.Contains("分") || format.Contains("秒");
        }

        // 处理日期类型
        if (isDateFormat && cell.Value is double numericDate)
        {
            try
            {
                DateTime dateTime = DateTime.FromOADate(numericDate);
                if (dateTime.TimeOfDay.TotalSeconds == 0)
                {
                    return dateTime.ToString("yyyy-MM-dd");
                }
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            catch
            {
                // 转换失败，按数值处理
                if (numericDate % 1 == 0)
                {
                    return Convert.ToInt64(numericDate);
                }
                else
                {
                    return numericDate;
                }
            }
        }
        else if (isDateFormat && cell.Value is DateTime dateTimeVal)
        {
            if (dateTimeVal.TimeOfDay.TotalSeconds == 0)
            {
                return dateTimeVal.ToString("yyyy-MM-dd");
            }
            return dateTimeVal.ToString("yyyy-MM-dd HH:mm:ss");
        }
        
        // 处理其他类型
        if (cell.Value is DateTime dateTimeValue)
        {
            if (dateTimeValue.TimeOfDay.TotalSeconds == 0)
            {
                return dateTimeValue.ToString("yyyy-MM-dd");
            }
            return dateTimeValue.ToString("yyyy-MM-dd HH:mm:ss");
        }
        else if (cell.Value is double d)
        {
            // 整数
            if (d % 1 == 0)
            {
                return Convert.ToInt64(d);
            }
            else
            {
                return d;
            }
        }
        else
        {
            return cell.Value.ToString()?.Trim() ?? string.Empty;
        }
    }
}