using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services.TableDefinitions;

/// <summary>
/// 数据字典表定义
/// </summary>
public class DataDictionaryTableDefinition : ITableDefinition
{
    public string TableName => TableNames.DataDictionaries;

    public string CreateTableSql => $@"
                CREATE TABLE IF NOT EXISTS {TableName} (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL,
                    TableName TEXT,
                    Description TEXT,
                    CreatedAt TEXT NOT NULL,
                    UpdatedAt TEXT NOT NULL,
                    CreatedBy TEXT DEFAULT 'admin',
                    IsActive INTEGER DEFAULT 1
                )";

    public string? CreateIndexSql =>  $@"
                CREATE INDEX IF NOT EXISTS idx_{TableName}_Name ON {TableName}(Name);
                CREATE INDEX IF NOT EXISTS idx_{TableName}_TableName ON {TableName}(TableName);";

    public async Task InitializeDefaultDataAsync(DatabaseService dbService)
    {
        
        // 检查是否已有数据
        var checkSql = $"SELECT COUNT(*) FROM {TableName}";
        var result = await dbService.ExecuteQueryAsync(checkSql, new List<object>());
        
        if (result != null && result.Count > 1 && Convert.ToInt32(result[1][0]) > 0)
        {
            return;
        }
        
        var insertSql = $@"
                    INSERT INTO {TableName} 
                        (Name, TableName, Description, CreatedAt, UpdatedAt, CreatedBy, IsActive)
                    VALUES (@p0, @p1, @p2, @p3, @p4, @p5, @p6)";

        var defaultItems = new[]
        {
            new { 
                Name = "默认表名",
                TableName = string.Empty,
                Description = "当表名为空时，使用表格本身名称",
                CreatedBy = "admin" ,
                IsActive = 1
            }
        };

        foreach (var item in defaultItems)
        {
            var parameters = new List<object>
            {
                item.Name,
                item.TableName,
                item.Description,
                DateTime.Now.ToString("o"),
                DateTime.Now.ToString("o"),
                item.CreatedBy,
                item.IsActive
            };
            
            await dbService.ExecuteNonQueryAsync(insertSql, parameters);
        }
    }
}