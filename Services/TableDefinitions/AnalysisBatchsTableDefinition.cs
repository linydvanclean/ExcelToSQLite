using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services.TableDefinitions;

/// <summary>
/// 分析批次
/// </summary>
public class AnalysisBatchsTableDefinition : ITableDefinition
{
    public string TableName => TableNames.AnalysisBatches;

    public string CreateTableSql => $@"
            CREATE TABLE IF NOT EXISTS {TableName} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Name TEXT NOT NULL,
                PeriodStart TEXT NOT NULL,
                PeriodEnd TEXT NOT NULL,
                CreatedAt TEXT NOT NULL,
                TablePrefix TEXT DEFAULT ''
            )";

    public string? CreateIndexSql => $@"
        CREATE INDEX IF NOT EXISTS idx_{TableName}_Id ON {TableName}(Id);";
    
    public async Task InitializeDefaultDataAsync(DatabaseService dbService)
    {
        try
        {
            // 检查是否已有数据
            var checkSql = $"SELECT COUNT(*) FROM {TableName}";
            var result = await dbService.ExecuteQueryAsync(checkSql, new List<object>());
            
            if (result != null && result.Count > 1 && Convert.ToInt32(result[1][0]) > 0)
            {
                return;
            }
            
            var insertSql = $@"
                INSERT INTO {TableName} (Name, PeriodStart, PeriodEnd, CreatedAt, TablePrefix)
                VALUES (@p0, @p1, @p2, @p3, @p4)";
            
            var now = DateTime.Now;
            var yearStart = new DateTime(now.Year, 1, 1);
            var yearEnd = new DateTime(now.Year, 12, 31);
            var parameters = new List<object>
            {
                "默认分析",
                yearStart.ToString("o"),
                yearEnd.ToString("o"),
                now.ToString("o"),
                string.Empty
            };
            await dbService.ExecuteNonQueryAsync(insertSql, parameters);
        }
        catch
        {
            throw;
        }
    }
}