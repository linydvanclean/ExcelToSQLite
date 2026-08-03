using System;
using System.Collections.Generic;

namespace ExcelToSQLite.Models
{
    /// <summary>
    /// 应用配置模型
    /// </summary>
    public class AppConfig
    {
        public string Version { get; set; } = PublicEvent.Version;
        public string SystemName { get; set; } = "智慧监督数据汇集分析平台";
        public string DeepSeekApiKey {get;set;} = "你的DeepSekk API Key";
        public string DeepSeekApiEndpoint { get; set; } = PublicEvent.DeepseekBaseUrl;
        public string DeepSeekModel { get; set; } = PublicEvent.DeepseekDefaultModel;
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        public List<string> Categories { get; set; } = new();
        public Dictionary<string, object> Values { get; set; } = new();
    }
}