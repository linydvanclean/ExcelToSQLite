using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;

namespace ExcelToSQLite.Converters
{
    // ============================================================
    // 1. 交替行颜色转换器
    // ============================================================
    public class AlternatingRowConverter : IValueConverter
    {
        public static readonly AlternatingRowConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is int index)
            {
                return index % 2 == 0 ? "#F8F9FA" : "White";
            }
            return "White";
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 2. 布尔值转文本转换器（▶/▼）
    // ============================================================
    public class BooleanToTextConverter : IValueConverter
    {
        public static readonly BooleanToTextConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is bool isExpanded)
            {
                // 如果有参数，使用自定义文本
                if (parameter is string param)
                {
                    var parts = param.Split('|');
                    if (parts.Length == 2)
                    {
                        return isExpanded ? parts[0] : parts[1];
                    }
                }
                return isExpanded ? "▼" : "▶";
            }
            return "▶";
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 3. 分类颜色转换器
    // ============================================================
    public class CategoryColorConverter : IValueConverter
    {
        public static readonly CategoryColorConverter Instance = new();

        private static readonly IBrush BaseDataBrush = new SolidColorBrush(Color.Parse("#2196F3"));
        private static readonly IBrush FinanceBrush = new SolidColorBrush(Color.Parse("#4CAF50"));
        private static readonly IBrush BusinessBrush = new SolidColorBrush(Color.Parse("#FF9800"));
        private static readonly IBrush PerformanceBrush = new SolidColorBrush(Color.Parse("#9C27B0"));
        private static readonly IBrush UserBrush = new SolidColorBrush(Color.Parse("#00BCD4"));
        private static readonly IBrush SystemBrush = new SolidColorBrush(Color.Parse("#607D8B"));
        private static readonly IBrush DefaultBrush = new SolidColorBrush(Color.Parse("#78909C"));

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            var category = value as string ?? string.Empty;

            // 如果参数是 "String"，返回颜色字符串而不是 Brush
            if (parameter is string param && param.Equals("String", StringComparison.OrdinalIgnoreCase))
            {
                return category switch
                {
                    "基础数据" => "#2196F3",
                    "财务分析" => "#4CAF50",
                    "业务分析" => "#FF9800",
                    "绩效分析" => "#9C27B0",
                    "用户分析" => "#00BCD4",
                    "系统分析" => "#607D8B",
                    _ => "#78909C"
                };
            }

            return category switch
            {
                "基础数据" => BaseDataBrush,
                "财务分析" => FinanceBrush,
                "业务分析" => BusinessBrush,
                "绩效分析" => PerformanceBrush,
                "用户分析" => UserBrush,
                "系统分析" => SystemBrush,
                _ => DefaultBrush
            };
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 4. 集合是否包含元素转换器
    // ============================================================
    public class CollectionHasItemsConverter : IValueConverter
    {
        public static readonly CollectionHasItemsConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value == null)
                return false;

            if (value is ICollection collection)
                return collection.Count > 0;

            if (value is IEnumerable enumerable)
            {
                var enumerator = enumerable.GetEnumerator();
                try
                {
                    return enumerator.MoveNext();
                }
                finally
                {
                    if (enumerator is IDisposable disposable)
                        disposable.Dispose();
                }
            }

            return false;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 5. 颜色转画刷转换器
    // ============================================================
    public class ColorToBrushConverter : IValueConverter
    {
        public static readonly ColorToBrushConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is ISolidColorBrush brush)
                return brush;

            if (value is IBrush brush2)
                return brush2;

            if (value is string colorString)
            {
                try
                {
                    return new SolidColorBrush(Color.Parse(colorString));
                }
                catch
                {
                    return new SolidColorBrush(Colors.Green);
                }
            }

            if (value is Color color)
                return new SolidColorBrush(color);

            return new SolidColorBrush(Colors.Green);
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 6. 画刷转颜色转换器
    // ============================================================
    public class BrushToColorConverter : IValueConverter
    {
        public static readonly BrushToColorConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
                return brush.Color;

            if (value is IBrush)
                return Colors.Green;

            if (value is string colorString)
            {
                try
                {
                    return Color.Parse(colorString);
                }
                catch
                {
                    return Colors.Green;
                }
            }

            return Colors.Green;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is Color color)
                return new SolidColorBrush(color);
            return new SolidColorBrush(Colors.Green);
        }
    }

    // ============================================================
    // 7. 字符串转画刷转换器
    // ============================================================
    public class StringToBrushConverter : IValueConverter
    {
        public static readonly StringToBrushConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is string colorString && !string.IsNullOrEmpty(colorString))
            {
                try
                {
                    return new SolidColorBrush(Color.Parse(colorString));
                }
                catch
                {
                    return new SolidColorBrush(Colors.Green);
                }
            }

            if (value is IBrush brush)
                return brush;

            return new SolidColorBrush(Colors.Green);
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
                return brush.Color.ToString();
            return "#4CAF50";
        }
    }

    // ============================================================
    // 8. 字符串转布尔值转换器
    // ============================================================
    public class StringToBoolConverter : IValueConverter
    {
        public static readonly StringToBoolConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            bool hasValue = false;

            if (value is string str)
            {
                hasValue = !string.IsNullOrWhiteSpace(str);
            }
            else if (value is int intValue)
            {
                hasValue = intValue > 0;
            }
            else if (value is long longValue)
            {
                hasValue = longValue > 0;
            }
            else if (value is bool boolValue)
            {
                hasValue = boolValue;
            }
            else if (value != null)
            {
                hasValue = true;
            }

            if (parameter is string param && param.Equals("Invert", StringComparison.OrdinalIgnoreCase))
            {
                return !hasValue;
            }

            return hasValue;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 9. 判断默认批次转换器
    // ============================================================
    public class IsDefaultBatchConverter : IValueConverter
    {
        public static readonly IsDefaultBatchConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            bool result = false;

            if (value is int id)
            {
                result = id == 1; // 默认批次 ID 为 1
            }
            else if (value is string strId && int.TryParse(strId, out int parsedId))
            {
                result = parsedId == 1;
            }

            if (parameter is string param && param.Equals("Invert", StringComparison.OrdinalIgnoreCase))
            {
                return !result;
            }

            return result;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 10. 布尔值取反转换器
    // ============================================================
    public class NotConverter : IValueConverter
    {
        public static readonly NotConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return true;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return false;
        }
    }

    // ============================================================
    // 11. 布尔值取反转换器（别名）
    // ============================================================
    public class InvertBoolConverter : IValueConverter
    {
        public static readonly InvertBoolConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return true;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return false;
        }
    }

    // ============================================================
    // 12. 对象转布尔值（是否为空）
    // ============================================================
    public class ObjectToBoolConverter : IValueConverter
    {
        public static readonly ObjectToBoolConverter Instance = new();

        public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            bool result = value != null;

            if (value is string str)
            {
                result = !string.IsNullOrWhiteSpace(str);
            }
            else if (value is ICollection collection)
            {
                result = collection.Count > 0;
            }
            else if (value is IEnumerable enumerable)
            {
                var enumerator = enumerable.GetEnumerator();
                try
                {
                    result = enumerator.MoveNext();
                }
                finally
                {
                    if (enumerator is IDisposable disposable)
                        disposable.Dispose();
                }
            }

            if (parameter is string param && param.Equals("Invert", StringComparison.OrdinalIgnoreCase))
            {
                return !result;
            }

            return result;
        }

        public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    // ============================================================
    // 13. 通用转换器集合
    // ============================================================
    public static class Converters
    {
        // 所有转换器的静态实例
        public static AlternatingRowConverter AlternatingRow => AlternatingRowConverter.Instance;
        public static BooleanToTextConverter BooleanToText => BooleanToTextConverter.Instance;
        public static CategoryColorConverter CategoryColor => CategoryColorConverter.Instance;
        public static CollectionHasItemsConverter CollectionHasItems => CollectionHasItemsConverter.Instance;
        public static ColorToBrushConverter ColorToBrush => ColorToBrushConverter.Instance;
        public static BrushToColorConverter BrushToColor => BrushToColorConverter.Instance;
        public static StringToBrushConverter StringToBrush => StringToBrushConverter.Instance;
        public static StringToBoolConverter StringToBool => StringToBoolConverter.Instance;
        public static IsDefaultBatchConverter IsDefaultBatch => IsDefaultBatchConverter.Instance;
        public static NotConverter Not => NotConverter.Instance;
        public static InvertBoolConverter InvertBool => InvertBoolConverter.Instance;
        public static ObjectToBoolConverter ObjectToBool => ObjectToBoolConverter.Instance;
    }
}