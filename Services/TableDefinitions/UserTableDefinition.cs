using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services.TableDefinitions;

/// <summary>
/// 用户表定义
/// </summary>
public class UserTableDefinition : ITableDefinition
{
    public string TableName => TableNames.Users;
    private const string DefaultAdminUsername = "admin";
    private const string DefaultAdminPassword = "123456";

    public string CreateTableSql => $@"
            CREATE TABLE IF NOT EXISTS {TableName} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Username TEXT NOT NULL UNIQUE,
                PasswordHash TEXT NOT NULL,
                Role TEXT NOT NULL DEFAULT 'User',
                CreatedAt TEXT NOT NULL,
                LastLoginAt TEXT
            )";

    public string? CreateIndexSql => $@"
        CREATE INDEX IF NOT EXISTS idx_{TableName}_Username ON {TableName}(Username);";

    public async Task InitializeDefaultDataAsync(DatabaseService dbService)
    {
        // 检查是否已有数据
        var checkSql = $"SELECT COUNT(*) FROM {TableName}";
        var result = await dbService.ExecuteQueryAsync(checkSql, new List<object>());
        
        var count = 0;
        if (result != null && result.Count > 1 && Convert.ToInt32(result[1][0]) > 0)
        {
            count = Convert.ToInt32(result[1][0]);
        }
        
        if (count == 0)
        {
            string insertSql = $@"
                    INSERT INTO {TableName} (Username, PasswordHash, Role, CreatedAt)
                    VALUES (@p0, @p1, 'Admin', @p2)";

            var insertParams = new List<object> 
            { 
                DefaultAdminUsername,
                PublicEvent.HashString(DefaultAdminPassword),
                PublicEvent.GetCurrentTimestamp()
            };

            await dbService.ExecuteNonQueryAsync(insertSql, insertParams);
        }
        else
        {
            // 管理员用户已存在
        }
    }
}