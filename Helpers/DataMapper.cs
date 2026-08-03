using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;
using System.Collections.Concurrent;

namespace ExcelToSQLite.Helpers;

/// <summary>
/// 数据映射器 - 将数据库行数据自动映射到对象
/// </summary>
public static class DataMapper
{
    // 缓存属性信息，提高性能
    private static readonly ConcurrentDictionary<Type, PropertyInfo[]> _propertyCache = new();
    private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> _propertyDictCache = new();

    /// <summary>
    /// 将行数据映射到指定类型的对象
    /// </summary>
    /// <typeparam name="T">目标类型</typeparam>
    /// <param name="row">数据行（List<object>）</param>
    /// <param name="columnNames">列名列表（可选，如果提供则按列名映射）</param>
    /// <returns>映射后的对象，失败返回 null</returns>
    public static T? MapRowToObject<T>(List<object>? row, List<string>? columnNames = null) where T : class, new()
    {
        if (row == null || row.Count == 0)
            return null;

        try
        {
            var obj = new T();
            var type = typeof(T);
            
            var propertyDict = GetPropertyDictionary(type);
            
            if (columnNames != null && columnNames.Count > 0)
            {
                for (int i = 0; i < Math.Min(row.Count, columnNames.Count); i++)
                {
                    var columnName = columnNames[i];
                    var value = row[i];
                    
                    if (string.IsNullOrEmpty(columnName) || value == null)
                        continue;

                    var key = columnName.Trim();
                    if (propertyDict.TryGetValue(key, out var property) ||
                        propertyDict.TryGetValue(key.ToLowerInvariant(), out property))
                    {
                        SetPropertyValue(property, obj, value);
                    }
                }
            }
            else
            {
                var properties = GetProperties(type);
                for (int i = 0; i < Math.Min(row.Count, properties.Length); i++)
                {
                    if (row[i] == null)
                        continue;
                    
                    SetPropertyValue(properties[i], obj, row[i]);
                }
            }

            return obj;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 将数据表映射到对象列表
    /// </summary>
    /// <typeparam name="T">目标类型</typeparam>
    /// <param name="data">查询结果（第一行为列名，后续为数据）</param>
    /// <returns>对象列表</returns>
    public static List<T> MapDataToList<T>(List<List<object>>? data) where T : class, new()
    {
        var result = new List<T>();
        
        if (data == null || data.Count <= 1)
            return result;

        var columnNames = data[0]?.Select(x => x?.ToString() ?? string.Empty).ToList() ?? new List<string>();

        for (int i = 1; i < data.Count; i++)
        {
            var row = data[i];
            if (row == null || row.Count == 0)
                continue;

            var obj = MapRowToObject<T>(row, columnNames);
            if (obj != null)
            {
                result.Add(obj);
            }
        }

        return result;
    }

    /// <summary>
    /// 将数据表映射到单个对象
    /// </summary>
    public static T? MapDataToObject<T>(List<List<object>>? data) where T : class, new()
    {
        if (data == null || data.Count <= 1)
            return null;

        var columnNames = data[0]?.Select(x => x?.ToString() ?? string.Empty).ToList() ?? new List<string>();
        return MapRowToObject<T>(data[1], columnNames);
    }

    #region 私有方法

    private static Dictionary<string, PropertyInfo> GetPropertyDictionary(Type type)
    {
        if (_propertyDictCache.TryGetValue(type, out var dict))
            return dict;

        var properties = GetProperties(type);
        dict = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
        
        foreach (var prop in properties)
        {
            dict[prop.Name] = prop;
            
            var withUnderscore = AddUnderscores(prop.Name);
            if (withUnderscore != prop.Name)
            {
                dict[withUnderscore] = prop;
            }
        }

        _propertyDictCache[type] = dict;
        return dict;
    }

    private static PropertyInfo[] GetProperties(Type type)
    {
        if (_propertyCache.TryGetValue(type, out var properties))
            return properties;

        properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                         .Where(p => p.CanWrite)
                         .ToArray();
        
        _propertyCache[type] = properties;
        return properties;
    }

    private static void SetPropertyValue<T>(PropertyInfo property, T obj, object value) where T : class
    {
        try
        {
            if (value == null || value == DBNull.Value)
                return;

            var targetType = property.PropertyType;
            var sourceValue = value.ToString();

            if (string.IsNullOrEmpty(sourceValue))
            {
                if (targetType.IsValueType && Nullable.GetUnderlyingType(targetType) == null)
                {
                    property.SetValue(obj, Activator.CreateInstance(targetType));
                }
                return;
            }

            object convertedValue;

            if (targetType == typeof(string))
            {
                convertedValue = sourceValue;
            }
            else if (targetType == typeof(int))
            {
                convertedValue = int.TryParse(sourceValue, out int intVal) ? intVal : 0;
            }
            else if (targetType == typeof(long))
            {
                convertedValue = long.TryParse(sourceValue, out long longVal) ? longVal : 0;
            }
            else if (targetType == typeof(decimal))
            {
                convertedValue = decimal.TryParse(sourceValue, out decimal decVal) ? decVal : 0;
            }
            else if (targetType == typeof(double))
            {
                convertedValue = double.TryParse(sourceValue, out double dblVal) ? dblVal : 0;
            }
            else if (targetType == typeof(bool))
            {
                convertedValue = sourceValue == "1" || sourceValue.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            else if (targetType == typeof(DateTime))
            {
                if (!DateTime.TryParse(sourceValue, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTime dtVal))
                {
                    dtVal = DateTime.Now;
                }
                convertedValue = dtVal;
            }
            else if (targetType == typeof(DateTime?))
            {
                // ✅ 使用 null! 消除 CS8600 警告
                if (DateTime.TryParse(sourceValue, null, System.Globalization.DateTimeStyles.RoundtripKind, out DateTime dtVal))
                {
                    convertedValue = dtVal;
                }
                else
                {
                    convertedValue = null!;
                }
            }
            else if (targetType.IsEnum)
            {
                try
                {
                    if (Enum.TryParse(targetType, sourceValue, true, out var enumVal))
                    {
                        convertedValue = enumVal;
                    }
                    else
                    {
                        var enumValues = Enum.GetValues(targetType);
                        if (enumValues.Length > 0)
                        {
                            var defaultEnumValue = enumValues.GetValue(0);
                            convertedValue = defaultEnumValue ?? (object)0;
                        }
                        else
                        {
                            convertedValue = (object)0;
                        }
                    }
                }
                catch
                {
                    try
                    {
                        var enumValues = Enum.GetValues(targetType);
                        if (enumValues.Length > 0)
                        {
                            var defaultEnumValue = enumValues.GetValue(0);
                            convertedValue = defaultEnumValue ?? (object)0;
                        }
                        else
                        {
                            convertedValue = (object)0;
                        }
                    }
                    catch
                    {
                        convertedValue = (object)0;
                    }
                }
            }
            else
            {
                convertedValue = Convert.ChangeType(value, targetType);
            }

            property.SetValue(obj, convertedValue);
        }
        catch
        {
        }
    }

    private static string AddUnderscores(string name)
    {
        if (string.IsNullOrEmpty(name))
            return name;

        var result = string.Empty;
        for (int i = 0; i < name.Length; i++)
        {
            if (i > 0 && char.IsUpper(name[i]) && !char.IsUpper(name[i - 1]))
            {
                result += "_";
            }
            result += name[i];
        }
        return result;
    }

    #endregion
}