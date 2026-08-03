using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services.TableDefinitions;

/// <summary>
/// 扫描结果表定义
/// </summary>
public class ScanResultsTableDefinition : ITableDefinition
{
    public string TableName => TableNames.ScanResults;

    public string CreateTableSql => $@"
            CREATE TABLE IF NOT EXISTS {TableName} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                BatchId INTEGER NOT NULL,
                BatchName TEXT NOT NULL,
                IndicatorId TEXT NOT NULL,
                IndicatorName TEXT NOT NULL,
                RowCount INTEGER DEFAULT 0,
                Status TEXT NOT NULL,
                ErrorMessage TEXT,
                SqlStatement TEXT NOT NULL,
                ScanTime TEXT NOT NULL,
                Duration TEXT,
                FOREIGN KEY (BatchId) REFERENCES {TableNames.AnalysisBatches}(Id)
            )";

    public string? CreateIndexSql => $@"
        CREATE INDEX IF NOT EXISTS idx_{TableName}_batchid ON {TableName} (BatchId);
        CREATE INDEX IF NOT EXISTS idx_{TableName}_indicatorid ON {TableName} (IndicatorId);";

    public Task InitializeDefaultDataAsync(DatabaseService dbService)
    {
        return Task.CompletedTask;
    }
}