using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services.TableDefinitions;

/// <summary>
/// 指标表定义
/// </summary>
public class IndicatorTableDefinition : ITableDefinition
{
    public string TableName => TableNames.Indicators;

    public string CreateTableSql => $@"
        CREATE TABLE IF NOT EXISTS {TableName} (
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Name TEXT NOT NULL,
            SqlStatement TEXT NOT NULL,
            SqlDetailData TEXT,
            Description TEXT,
            Category TEXT,
            CreatedAt TEXT NOT NULL,
            UpdatedAt TEXT NOT NULL,
            CreatedBy TEXT DEFAULT 'admin',
            IsActive INTEGER DEFAULT 1
        )";

    public string? CreateIndexSql => $@"
        CREATE INDEX IF NOT EXISTS idx_{TableName}_Name ON {TableName}(Name);";

    public Task InitializeDefaultDataAsync(DatabaseService dbService)
    {
        return Task.CompletedTask;
    }
}