using System;
using System.Collections.Generic;

namespace ExcelToSQLite.Models
{
    /// <summary>
    /// 指标导出数据契约
    /// </summary>
    [Serializable]
    public class IndicatorExportData
    {
        public string Version { get; set; } = PublicEvent.Version;
        public DateTime ExportTime { get; set; } = DateTime.Now;
        public string ExportedBy { get; set; } = string.Empty;
        public List<IndicatorItem> Indicators { get; set; } = new();
    }

    /// <summary>
    /// 单个指标数据项
    /// </summary>
    [Serializable]
    public class IndicatorItem
    {
        public string Name { get; set; } = string.Empty;
        public string SqlStatement { get; set; } = string.Empty;
        public string SqlDetailData { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public string CreatedBy { get; set; } = "admin";
        public bool IsActive { get; set; } = true;
    }
}