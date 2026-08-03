using System;

namespace ExcelToSQLite.Helpers
{
    public static class DataPreviewHelper
    {
        /// <summary>
        /// 生成表数据预览 SQL（前 N 条）
        /// </summary>
        public static string BuildTablePreviewSql(string tableName, int limit = 10000)
        {
            if (string.IsNullOrWhiteSpace(tableName))
                return string.Empty;
            
            // 转义表名（防止 SQL 注入）
            var escapedTableName = tableName.Replace("\"", "\"\"");
            return $"SELECT * FROM \"{escapedTableName}\" LIMIT {limit}";
        }

        /// <summary>
        /// 生成表数据导出 SQL（全部数据）
        /// </summary>
        public static string BuildTableExportSql(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
                return string.Empty;
            
            var escapedTableName = tableName.Replace("\"", "\"\"");
            return $"SELECT * FROM \"{escapedTableName}\"";
        }
    }
}