using System;
using System.Collections.Generic;

namespace ExcelToSQLite.Models;

[Serializable]
public class DataDictionaryExportData
{
    public string Version { get; set; } = PublicEvent.Version;
    public DateTime ExportTime { get; set; } = DateTime.Now;
    public string ExportedBy { get; set; } = string.Empty;
    public List<DataDictionaryItem> Dictionaries { get; set; } = new();
}

[Serializable]
public class DataDictionaryItem
{
    public string Name { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public string CreatedBy { get; set; } = "admin";
    public bool IsActive { get; set; } = true;
}
