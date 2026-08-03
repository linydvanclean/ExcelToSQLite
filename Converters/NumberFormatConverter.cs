// Converters/NumberFormatConverter.cs
using System;
using System.Globalization;
using Avalonia.Data.Converters;

namespace ExcelToSQLite.Converters
{
    public class NumberFormatConverter : IValueConverter
    {
        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value == null) return "0";
            
            try
            {
                if (value is int intValue)
                    return intValue.ToString("N0", CultureInfo.CurrentCulture);
                if (value is long longValue)
                    return longValue.ToString("N0", CultureInfo.CurrentCulture);
                if (value is double doubleValue)
                    return doubleValue.ToString("N0", CultureInfo.CurrentCulture);
                if (value is decimal decimalValue)
                    return decimalValue.ToString("N0", CultureInfo.CurrentCulture);
                
                return value.ToString() ?? "0";
            }
            catch
            {
                return value?.ToString() ?? "0";
            }
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}