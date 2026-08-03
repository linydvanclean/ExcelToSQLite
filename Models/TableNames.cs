using System.Collections.Generic;

namespace ExcelToSQLite.Models;

public static class TableNames
{
    public const string Users = "sys_Users";
    public const string AnalysisBatches = "sys_AnalysisBatches";
    public const string DataDictionaries = "sys_DataDictionaries";
    public const string Indicators = "sys_Indicators";
    public const string ScanResults = "sys_ScanResults";
    public const string Sqlite_sequence = "sqlite_sequence";
    
    // 如果需要获取所有表名列表
    public static readonly IReadOnlyList<string> All = new[]
    {
        Users,
        AnalysisBatches,
        DataDictionaries,
        AnalysisBatches,
        Indicators,
        ScanResults,
        Sqlite_sequence,
    };
    
    // 用于白名单验证的 HashSet
    public static readonly HashSet<string> AllowedSet = new(All);
}