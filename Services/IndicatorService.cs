using ExcelToSQLite.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.Services;

/// <summary>
/// 指标服务 - 管理指标的 CRUD 操作
/// </summary>
public class IndicatorService : DisposableBase
{
    private readonly DatabaseService _databaseService;
    private readonly string _tableName = TableNames.Indicators;
    private readonly SemaphoreSlim _initSemaphore = new SemaphoreSlim(1, 1);
    private readonly bool _debugMode = true;

    public IndicatorService()
    {
        _databaseService = DatabaseService.Instance;
    }

    protected override void DisposeManagedResources()
    {
        _initSemaphore?.Dispose();
    }

    #region 基础 CRUD 操作

    /// <summary>
    /// 添加指标
    /// </summary>
    public async Task<bool> AddAsync(Indicator indicator)
    {
        try
        {
            if (indicator == null)
                throw new ArgumentNullException(nameof(indicator));

            if (string.IsNullOrEmpty(indicator.Name))
                throw new ArgumentException("指标名称不能为空", nameof(indicator));

            if (string.IsNullOrEmpty(indicator.SqlStatement))
                throw new ArgumentException("SQL语句不能为空", nameof(indicator));

            var sql = $@"
            INSERT INTO {_tableName} (Name, SqlStatement, SqlDetailData, Description, Category, CreatedAt, UpdatedAt, CreatedBy, IsActive)
            VALUES (@p0, @p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8)";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@p0", indicator.Name),
                new SqliteParameter("@p1", indicator.SqlStatement ?? string.Empty),
                new SqliteParameter("@p2", indicator.SqlDetailData ?? string.Empty),
                new SqliteParameter("@p3", indicator.Description ?? string.Empty),
                new SqliteParameter("@p4", indicator.Category ?? string.Empty),
                new SqliteParameter("@p5", DateTime.Now.ToString("o")),
                new SqliteParameter("@p6", DateTime.Now.ToString("o")),
                new SqliteParameter("@p7", indicator.CreatedBy ?? "admin"),
                new SqliteParameter("@p8", indicator.IsActive ? 1 : 0)
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);
            return rowsAffected > 0;
        }
        catch (Exception ex)
        {
            LogError($"添加指标失败: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 批量添加指标
    /// </summary>
    public async Task<int> AddRangeAsync(IEnumerable<Indicator> indicators)
    {
        try
        {
            if (indicators == null)
                throw new ArgumentNullException(nameof(indicators));

            var indicatorList = indicators.ToList();
            if (!indicatorList.Any())
                return 0;

            var sql = $@"
            INSERT INTO {_tableName} (Name, SqlStatement, SqlDetailData, Description, Category, CreatedAt, UpdatedAt, CreatedBy, IsActive)
            VALUES (@p0, @p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8)";

            var successCount = 0;
            foreach (var indicator in indicatorList)
            {
                try
                {
                    if (string.IsNullOrEmpty(indicator.Name))
                        continue;

                    var parameters = new List<SqliteParameter>
                    {
                        new SqliteParameter("@p0", indicator.Name),
                        new SqliteParameter("@p1", indicator.SqlStatement ?? string.Empty),
                        new SqliteParameter("@p2", indicator.SqlDetailData ?? string.Empty),
                        new SqliteParameter("@p3", indicator.Description ?? string.Empty),
                        new SqliteParameter("@p4", indicator.Category ?? string.Empty),
                        new SqliteParameter("@p5", DateTime.Now.ToString("o")),
                        new SqliteParameter("@p6", DateTime.Now.ToString("o")),
                        new SqliteParameter("@p7", indicator.CreatedBy ?? "admin"),
                        new SqliteParameter("@p8", indicator.IsActive ? 1 : 0)
                    };

                    await _databaseService.ExecuteNonQueryAsync(sql, parameters);
                    successCount++;
                }
                catch (Exception ex)
                {
                    LogError($"批量添加指标 '{indicator.Name}' 失败: {ex.Message}");
                }
            }
            return successCount;
        }
        catch (Exception ex)
        {
            LogError($"批量添加失败: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// 更新指标
    /// </summary>
    public async Task<bool> UpdateAsync(Indicator indicator)
    {
        try
        {
            if (indicator == null)
                throw new ArgumentNullException(nameof(indicator));

            if (string.IsNullOrEmpty(indicator.Id))
                throw new ArgumentException("指标ID不能为空", nameof(indicator));

            if (string.IsNullOrEmpty(indicator.Name))
                throw new ArgumentException("指标名称不能为空", nameof(indicator));

            if (string.IsNullOrEmpty(indicator.SqlStatement))
                throw new ArgumentException("SQL语句不能为空", nameof(indicator));

            var sql = $@"
            UPDATE {_tableName}
            SET Name = @p0,
                SqlStatement = @p1,
                SqlDetailData = @p2,
                Description = @p3,
                Category = @p4,
                UpdatedAt = @p5,
                IsActive = @p6
            WHERE Id = @p7";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@p0", indicator.Name),
                new SqliteParameter("@p1", indicator.SqlStatement ?? string.Empty),
                new SqliteParameter("@p2", indicator.SqlDetailData ?? string.Empty),
                new SqliteParameter("@p3", indicator.Description ?? string.Empty),
                new SqliteParameter("@p4", indicator.Category ?? string.Empty),
                new SqliteParameter("@p5", DateTime.Now.ToString("o")),
                new SqliteParameter("@p6", indicator.IsActive ? 1 : 0),
                new SqliteParameter("@p7", ParseId(indicator.Id))
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);

            if (rowsAffected == 0)
            {
                LogWarning($"未找到要更新的指标: {indicator.Id}");
                return false;
            }
            
            return true;
        }
        catch (Exception ex)
        {
            LogError($"更新指标失败: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 删除指标
    /// </summary>
    public async Task<bool> DeleteAsync(string id)
    {
        try
        {
            if (string.IsNullOrEmpty(id))
                throw new ArgumentException("指标ID不能为空", nameof(id));

            var sql = $"DELETE FROM {_tableName} WHERE Id = @p0";
            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@p0", ParseId(id))
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);

            if (rowsAffected == 0)
            {
                LogWarning($"未找到要删除的指标: {id}");
                return false;
            }
            
            return true;
        }
        catch (Exception ex)
        {
            LogError($"删除指标失败: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// 批量删除指标
    /// </summary>
    public async Task<int> DeleteRangeAsync(IEnumerable<string> ids)
    {
        try
        {
            if (ids == null)
                throw new ArgumentNullException(nameof(ids));

            var idList = ids.Where(id => !string.IsNullOrEmpty(id))
                           .Select(id => ParseId(id))
                           .ToList();
            
            if (!idList.Any())
                return 0;

            var parameters = new List<SqliteParameter>();
            var placeholders = new List<string>();

            for (int i = 0; i < idList.Count; i++)
            {
                placeholders.Add($"@p{i}");
                parameters.Add(new SqliteParameter($"@p{i}", idList[i]));
            }

            var sql = $"DELETE FROM {_tableName} WHERE Id IN ({string.Join(", ", placeholders)})";

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);
            return rowsAffected;
        }
        catch (Exception ex)
        {
            LogError($"批量删除失败: {ex.Message}");
            // ✅ 修复: 不使用 ShowErrorSync，调用方负责显示错误
            return 0;
        }
    }

    #endregion

    #region 查询操作

    /// <summary>
    /// 获取所有指标（使用泛型映射）
    /// </summary>
    public async Task<List<Indicator>> GetAllAsync()
    {
        try
        {
            var sql = $"SELECT * FROM {_tableName} ORDER BY Id ASC";
            var data = await _databaseService.ExecuteQueryAsync(sql, new List<object>());
            var result = DataMapper.MapDataToList<Indicator>(data);
            return result;
        }
        catch (Exception ex)
        {
            LogError($"获取所有指标失败: {ex.Message}");
            return new List<Indicator>();
        }
    }

    /// <summary>
    /// 根据 ID 获取指标
    /// </summary>
public async Task<Indicator?> GetByIdAsync(string id)
{
    try
    {
        if (string.IsNullOrEmpty(id))
        {
            LogWarning("GetByIdAsync: ID 为空");
            return null;
        }

        if (!int.TryParse(id, out int intId))
        {
            LogWarning($"GetByIdAsync: ID 格式无效: {id}");
            return null;
        }

        var sql = $"SELECT * FROM {_tableName} WHERE Id = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", intId)
        };
        
        var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

        // 检查是否有数据（现在无数据时返回空列表）
        if (data == null || data.Count == 0)
        {
            LogWarning($"查询结果为空");
            return null;
        }
        
        var indicator = DataMapper.MapDataToObject<Indicator>(data);
        
        if (indicator != null)
        {
            return indicator;
        }
        else
        {
            // 手动映射（备用方案）
            if (data.Count >= 2)
            {
                var row = data[1];
                try
                {
                    indicator = new Indicator
                    {
                        Id = row[0]?.ToString() ?? string.Empty,
                        Name = row[1]?.ToString() ?? string.Empty,
                        SqlStatement = row[2]?.ToString() ?? string.Empty,
                        SqlDetailData = row[3]?.ToString() ?? string.Empty,
                        Description = row[4]?.ToString() ?? string.Empty,
                        Category = row[5]?.ToString() ?? string.Empty,
                        CreatedBy = row.Count > 8 ? row[8]?.ToString() ?? "admin" : "admin",
                        IsActive = row.Count > 9 && row[9]?.ToString() == "1"
                    };

                    // 解析日期
                    if (row.Count > 6 && row[6] != null)
                    {
                        DateTime.TryParse(row[6].ToString(), out var createdAt);
                        indicator.CreatedAt = createdAt;
                    }

                    if (row.Count > 7 && row[7] != null)
                    {
                        DateTime.TryParse(row[7].ToString(), out var updatedAt);
                        indicator.UpdatedAt = updatedAt;
                    }
                    return indicator;
                }
                catch
                {
                }
            }
        }
        LogWarning($"❌ 所有映射方式都失败，ID: {id}");
        return null;
    }
    catch (Exception ex)
    {
        LogError($"GetByIdAsync 异常: {ex.Message}");
        return null;
    }
}

    /// <summary>
    /// 根据名称获取指标（使用泛型映射）
    /// </summary>
    public async Task<Indicator?> GetByNameAsync(string name)
    {
        try
        {
            if (string.IsNullOrEmpty(name))
            {
                LogWarning("GetByNameAsync: 名称为空");
                return null;
            }

            var sql = $"SELECT * FROM {_tableName} WHERE Name = @p0";
            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@p0", name)
            };
            
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);
            
            // ✅ 现在无数据时 data 为空列表或 null
            if (data == null || data.Count == 0)
            {
                LogWarning($"未找到名称为 '{name}' 的数据");
                return null;
            }
            var indicator = DataMapper.MapDataToObject<Indicator>(data);
            return indicator;
        }
        catch (Exception ex)
        {
            LogError($"GetByNameAsync 异常: {ex.Message}");
            // ✅ 修复: 不使用 ShowErrorSync，返回 null 让调用方处理
            return null;
        }
    }

    /// <summary>
    /// 获取指标总数
    /// </summary>
    public async Task<int> GetCountAsync()
    {
        try
        {
            var sql = $"SELECT COUNT(*) FROM {_tableName}";
            var result = await _databaseService.ExecuteQueryAsync(sql, new List<object>());

            // ✅ 现在无数据时 result 为空列表
            if (result == null || result.Count == 0)
            {
                return 0;
            }

            // result[0] 是表头，result[1] 是数据
            if (result.Count > 1 && result[1].Count > 0)
            {
                return Convert.ToInt32(result[1][0]);
            }

            return 0;
        }
        catch (Exception ex)
        {
            LogError($"获取指标总数失败: {ex.Message}");
            return 0;
        }
    }

    /// <summary>
    /// 检查指标名称是否存在
    /// </summary>
    public async Task<bool> ExistsByNameAsync(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;

        var indicator = await GetByNameAsync(name);
        return indicator != null;
    }

    /// <summary>
    /// 检查指标 ID 是否存在
    /// </summary>
    public async Task<bool> ExistsByIdAsync(string id)
    {
        if (string.IsNullOrEmpty(id))
            return false;

        var indicator = await GetByIdAsync(id);
        return indicator != null;
    }

    #endregion

    #region 执行操作

    /// <summary>
    /// 执行指标 SQL
    /// </summary>
    public async Task<object?> ExecuteIndicatorAsync(Indicator indicator)
    {
        if (indicator == null)
            throw new ArgumentNullException(nameof(indicator));

        if (string.IsNullOrEmpty(indicator.SqlStatement))
            throw new InvalidOperationException($"指标 '{indicator.Name}' 的 SQL 语句为空");

        try
        {
            var result = await _databaseService.ExecuteQueryAsync(indicator.SqlStatement, new List<object>());
            return result;
        }
        catch (Exception ex)
        {
            LogError($"执行指标 '{indicator.Name}' 失败: {ex.Message}");
            throw new Exception($"执行指标 '{indicator.Name}' 失败: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// 验证 SQL 语法（使用 EXPLAIN）
    /// </summary>
    public async Task<bool> ValidateSqlAsync(string sql)
    {
        if (string.IsNullOrEmpty(sql))
            return false;

        try
        {
            var testSql = $"EXPLAIN {sql}";
            var result = await _databaseService.ExecuteQueryAsync(testSql, new List<object>());
            // 只要没有异常，且返回了结果（即使只有表头），就说明语法正确
            return result != null && result.Count > 0;
        }
        catch (Exception ex)
        {
            LogError($"SQL 验证失败: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region 私有方法

    /// <summary>
    /// 解析 ID（带错误处理）
    /// </summary>
    private int ParseId(string id)
    {
        if (string.IsNullOrEmpty(id))
            throw new ArgumentException("ID 不能为空", nameof(id));

        if (!int.TryParse(id, out int result))
            throw new FormatException($"ID 格式无效: {id}");

        return result;
    }
    /// <summary>
    /// 日志工具 - 警告
    /// </summary>
    private void LogWarning(string message)
    {
        if (_debugMode)
        {
            }
    }

    /// <summary>
    /// 日志工具 - 错误
    /// </summary>
    private void LogError(string message)
    {
    }

    #endregion
}