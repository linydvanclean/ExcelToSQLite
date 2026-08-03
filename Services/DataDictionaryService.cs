using ExcelToSQLite.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelToSQLite.Services.TableDefinitions;

namespace ExcelToSQLite.Services;

public class DataDictionaryService : IDisposable
{
    private readonly DatabaseService _databaseService;
    private readonly string _tableName = new DataDictionaryTableDefinition().TableName;
    private bool _isInitialized = false;
    private readonly SemaphoreSlim _initSemaphore = new SemaphoreSlim(1, 1);
    private bool _disposed = false;

    // 列名常量
    private static readonly List<string> _columns = new()
    {
        "Name",
        "TableName",
        "Description",
        "CreatedAt",
        "UpdatedAt",
        "CreatedBy",
        "IsActive"
    };

    public DataDictionaryService()
    {
        _databaseService = DatabaseService.Instance;
    }

    /// <summary>
    /// 初始化表结构（使用 SemaphoreSlim 异步锁，避免 Linux 死锁）
    /// </summary>
    public async Task InitializeAsync()
    {
        if (_isInitialized) return;

        await _initSemaphore.WaitAsync();
        try
        {
            if (_isInitialized) return;



            // ✅ 修复: 使用 await 而不是 .Wait()，避免在 Linux 上死锁
            await CreateTableAsync();
            _isInitialized = true;
        }
        catch
        {
            throw;
        }
        finally
        {
            _initSemaphore.Release();
        }

        // ✅ 在锁外部初始化默认数据
        await EnsureDefaultDataAsync();
    }

    /// <summary>
    /// 创建表（参考 AnalysisBatchService.CreateTableAsync）
    /// </summary>
    private async Task CreateTableAsync()
    {
        try
        {

            // ✅ 使用 ExecuteNonQueryAsync 直接创建表（与 AnalysisBatchService 一致）
            var createTableSql = $@"
                CREATE TABLE IF NOT EXISTS {_tableName} (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL,
                    TableName TEXT,
                    Description TEXT,
                    CreatedAt TEXT NOT NULL,
                    UpdatedAt TEXT NOT NULL,
                    CreatedBy TEXT DEFAULT 'admin',
                    IsActive INTEGER DEFAULT 1
                )";

            await _databaseService.ExecuteNonQueryAsync(createTableSql, new List<object>());

            // 创建索引
            var indexSql = $@"
                CREATE INDEX IF NOT EXISTS idx_{_tableName}_Name ON {_tableName}(Name);
                CREATE INDEX IF NOT EXISTS idx_{_tableName}_TableName ON {_tableName}(TableName);";

            await _databaseService.ExecuteNonQueryAsync(indexSql, new List<object>());

        }
        catch
        {
            throw;
        }
    }

    /// <summary>
    /// 确保默认数据（参考 AnalysisBatchService.EnsureDefaultDataAsync）
    /// </summary>
    private async Task EnsureDefaultDataAsync()
    {
        try
        {

            // ✅ 直接使用 ExecuteQueryAsync 检查数据（与 AnalysisBatchService 一致）
            var countSql = $"SELECT COUNT(*) FROM {_tableName}";
            var result = await _databaseService.ExecuteQueryAsync(countSql, new List<object>());

            int count = 0;
            if (result != null && result.Count > 1)
            {
                int.TryParse(result[1][0]?.ToString(), out count);
            }

            if (count == 0)
            {

                // ✅ 直接使用参数化插入（与 AnalysisBatchService 一致）
                var insertSql = $@"
                    INSERT INTO {_tableName} (Name, TableName, Description, CreatedAt, UpdatedAt, CreatedBy, IsActive)
                    VALUES (@p0, @p1, @p2, @p3, @p4, @p5, @p6)";

                var parameters = new List<SqliteParameter>
                {
                    new SqliteParameter("@p0", "默认表名"),
                    new SqliteParameter("@p1", string.Empty),
                    new SqliteParameter("@p2", "当表名为空时，使用表格本身名称"),
                    new SqliteParameter("@p3", DateTime.Now.ToString("o")),
                    new SqliteParameter("@p4", DateTime.Now.ToString("o")),
                    new SqliteParameter("@p5", "admin"),
                    new SqliteParameter("@p6", 1)
                };

                await _databaseService.ExecuteNonQueryAsync(insertSql, parameters);
            }
            else
            {
            }
        }
        catch
        {
          
        }
    }

    /// <summary>
    /// 确保表已初始化
    /// </summary>
    private async Task EnsureInitializedAsync()
    {
        if (!_isInitialized)
        {
            await InitializeAsync();
        }
    }

    /// <summary>
    /// 添加数据字典
    /// </summary>
    public async Task AddAsync(DataDictionary dictionary)
    {
        if (dictionary == null)
            throw new ArgumentNullException(nameof(dictionary));

        ValidateDictionary(dictionary);
        await EnsureInitializedAsync();



        var sql = $@"
            INSERT INTO {_tableName} (Name, TableName, Description, CreatedAt, UpdatedAt, CreatedBy, IsActive)
            VALUES (@p0, @p1, @p2, @p3, @p4, @p5, @p6)";

        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", dictionary.Name ?? string.Empty),
            new SqliteParameter("@p1", dictionary.TableName ?? string.Empty),
            new SqliteParameter("@p2", dictionary.Description ?? string.Empty),
            new SqliteParameter("@p3", dictionary.CreatedAt.ToString("o")),
            new SqliteParameter("@p4", dictionary.UpdatedAt.ToString("o")),
            new SqliteParameter("@p5", dictionary.CreatedBy ?? "admin"),
            new SqliteParameter("@p6", dictionary.IsActive ? 1 : 0)
        };

        await _databaseService.ExecuteNonQueryAsync(sql, parameters);
    }

    /// <summary>
    /// 更新数据字典
    /// </summary>
    public async Task UpdateAsync(DataDictionary dictionary)
    {
        if (dictionary == null)
            throw new ArgumentNullException(nameof(dictionary));

        if (string.IsNullOrEmpty(dictionary.Id))
            throw new ArgumentException("数据字典ID不能为空", nameof(dictionary));

        ValidateDictionary(dictionary);
        await EnsureInitializedAsync();



        var sql = $@"
            UPDATE {_tableName}
            SET Name = @p0,
                TableName = @p1,
                Description = @p2,
                UpdatedAt = @p3,
                IsActive = @p4
            WHERE Id = @p5";

        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", dictionary.Name ?? string.Empty),
            new SqliteParameter("@p1", dictionary.TableName ?? string.Empty),
            new SqliteParameter("@p2", dictionary.Description ?? string.Empty),
            new SqliteParameter("@p3", DateTime.Now.ToString("o")),
            new SqliteParameter("@p4", dictionary.IsActive ? 1 : 0),
            new SqliteParameter("@p5", int.Parse(dictionary.Id))
        };

        var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);

        if (rowsAffected == 0)
        {
            throw new KeyNotFoundException($"未找到要更新的数据字典: {dictionary.Id}");
        }

    }

    /// <summary>
    /// 删除数据字典
    /// </summary>
    public async Task DeleteAsync(string id)
    {
        if (string.IsNullOrEmpty(id))
            throw new ArgumentException("数据字典ID不能为空", nameof(id));

        await EnsureInitializedAsync();



        var sql = $"DELETE FROM {_tableName} WHERE Id = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", int.Parse(id))
        };

        var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);

        if (rowsAffected == 0)
        {
            throw new KeyNotFoundException($"未找到要删除的数据字典: {id}");
        }

    }

    /// <summary>
    /// 获取所有数据字典
    /// </summary>
    public async Task<List<DataDictionary>> GetAllAsync()
    {
        await EnsureInitializedAsync();



        var sql = $"SELECT * FROM {_tableName} ORDER BY Id ASC";
        var data = await _databaseService.ExecuteQueryAsync(sql, new List<object>());

        var result = new List<DataDictionary>();

        if (data == null || data.Count <= 1)
            return result;

        for (int i = 1; i < data.Count; i++)
        {
            try
            {
                var dictionary = ParseDictionary(data[i]);
                if (dictionary != null)
                {
                    result.Add(dictionary);
                }
            }
            catch
            {
            }
        }

        return result;
    }

    /// <summary>
    /// 根据 ID 获取数据字典
    /// </summary>
    public async Task<DataDictionary?> GetByIdAsync(string id)
    {
        if (string.IsNullOrEmpty(id))
            return null;

        await EnsureInitializedAsync();

        var sql = $"SELECT * FROM {_tableName} WHERE Id = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", int.Parse(id))
        };

        var data = await _databaseService.ExecuteQueryAsync(sql, parameters.Cast<object>().ToList());

        if (data == null || data.Count <= 1)
            return null;

        return ParseDictionary(data[1]);
    }

    /// <summary>
    /// 根据表名获取数据字典
    /// </summary>
    public async Task<DataDictionary?> GetByTableNameAsync(string tableName)
    {
        if (string.IsNullOrEmpty(tableName))
            return null;

        await EnsureInitializedAsync();

        var sql = $"SELECT * FROM {_tableName} WHERE TableName = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", tableName)
        };

        var data = await _databaseService.ExecuteQueryAsync(sql, parameters.Cast<object>().ToList());

        if (data == null || data.Count <= 1)
            return null;

        return ParseDictionary(data[1]);
    }

    /// <summary>
    /// 检查名称是否存在
    /// </summary>
    public async Task<bool> ExistsAsync(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;

        await EnsureInitializedAsync();

        var sql = $"SELECT COUNT(*) FROM {_tableName} WHERE Name = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", name)
        };

        var result = await _databaseService.ExecuteQueryAsync(sql, parameters.Cast<object>().ToList());

        if (result != null && result.Count > 1 && result[1].Count > 0)
        {
            return Convert.ToInt32(result[1][0]) > 0;
        }

        return false;
    }

    /// <summary>
    /// 检查表名是否存在
    /// </summary>
    public async Task<bool> TableNameExistsAsync(string tableName)
    {
        if (string.IsNullOrEmpty(tableName))
            return false;

        await EnsureInitializedAsync();

        var sql = $"SELECT COUNT(*) FROM {_tableName} WHERE TableName = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", tableName)
        };

        var result = await _databaseService.ExecuteQueryAsync(sql, parameters.Cast<object>().ToList());

        if (result != null && result.Count > 1 && result[1].Count > 0)
        {
            return Convert.ToInt32(result[1][0]) > 0;
        }

        return false;
    }

    /// <summary>
    /// 获取数据字典总数
    /// </summary>
    public async Task<int> GetCountAsync()
    {
        await EnsureInitializedAsync();

        var sql = $"SELECT COUNT(*) FROM {_tableName}";
        var result = await _databaseService.ExecuteQueryAsync(sql, new List<object>());

        if (result != null && result.Count > 1 && result[1].Count > 0)
        {
            return Convert.ToInt32(result[1][0]);
        }

        return 0;
    }

    #region 私有方法

    /// <summary>
    /// 解析行数据到 DataDictionary 对象
    /// </summary>
    private DataDictionary? ParseDictionary(List<object> row)
    {
        try
        {
            if (row == null || row.Count < 7)
                return null;

            var createdAt = DateTime.Now;
            var updatedAt = DateTime.Now;

            if (row[4] != null)
            {
                DateTime.TryParse(row[4].ToString(), null, System.Globalization.DateTimeStyles.RoundtripKind, out createdAt);
            }

            if (row[5] != null)
            {
                DateTime.TryParse(row[5].ToString(), null, System.Globalization.DateTimeStyles.RoundtripKind, out updatedAt);
            }

            return new DataDictionary
            {
                Id = row[0]?.ToString() ?? string.Empty,
                Name = row[1]?.ToString() ?? string.Empty,
                TableName = row[2]?.ToString() ?? string.Empty,
                Description = row[3]?.ToString() ?? string.Empty,
                CreatedAt = createdAt,
                UpdatedAt = updatedAt,
                CreatedBy = row[6]?.ToString() ?? "admin",
                IsActive = row.Count > 7 ? row[7]?.ToString() == "1" : true,
                Index = 0
            };
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 验证数据字典
    /// </summary>
    private void ValidateDictionary(DataDictionary dictionary)
    {
        if (dictionary == null)
            throw new ArgumentNullException(nameof(dictionary));

        if (string.IsNullOrEmpty(dictionary.Name))
            throw new ArgumentException("数据字典名称不能为空", nameof(dictionary));

        if (!string.IsNullOrEmpty(dictionary.TableName) &&
            !Regex.IsMatch(dictionary.TableName, @"^[a-zA-Z0-9_]+$"))
        {
            throw new ArgumentException("数据表名只能包含字母、数字和下划线", nameof(dictionary));
        }
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _initSemaphore?.Dispose();
            }
            _disposed = true;
        }
    }

    #endregion
}
