using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelToSQLite.Models;
using Microsoft.Data.Sqlite;
using System.Linq;
using System.Threading;
using ExcelToSQLite.Services.TableDefinitions;

namespace ExcelToSQLite.Services;

public class AnalysisBatchService : DisposableBase
{
    private readonly DatabaseService _databaseService;
    private readonly string _tableName = TableNames.AnalysisBatches;
    private readonly SemaphoreSlim _initSemaphore = new SemaphoreSlim(1, 1);

    public AnalysisBatchService()
    {
        _databaseService = DatabaseService.Instance;
    }
    protected override void DisposeManagedResources()
    {
        // 释放托管资源
        _initSemaphore?.Dispose();
    }
    public async Task<List<AnalysisBatch>> GetAllAsync(int limit = 1000)
    {
        var result = new List<AnalysisBatch>();

        var sql = $"SELECT * FROM {_tableName} ORDER BY CreatedAt DESC LIMIT {limit}";
        var data = await _databaseService.ExecuteQueryAsync(sql, new List<object>());

        if (data == null || data.Count <= 1)
            return result;

        for (int i = 1; i < data.Count; i++)
        {
            result.Add(ParseRow(data[i]));
        }

        return result;
    }

    private AnalysisBatch ParseRow(List<object> row)
    {
        int id = Convert.ToInt32(row[0]);
        string name = row[1]?.ToString() ?? string.Empty;
        string periodStartStr = row[2]?.ToString() ?? string.Empty;
        string periodEndStr = row[3]?.ToString() ?? string.Empty;
        string createdAtStr = row[4]?.ToString() ?? string.Empty;
        string tablePrefix = row.Count > 5 ? row[5]?.ToString() ?? string.Empty : string.Empty;

        DateTime.TryParse(periodStartStr, null, System.Globalization.DateTimeStyles.RoundtripKind, out var ps);
        DateTime.TryParse(periodEndStr, null, System.Globalization.DateTimeStyles.RoundtripKind, out var pe);
        DateTime.TryParse(createdAtStr, null, System.Globalization.DateTimeStyles.RoundtripKind, out var ca);

        return new AnalysisBatch
        {
            Id = id,
            Name = name,
            PeriodStart = ps,
            PeriodEnd = pe,
            CreatedAt = ca == default ? DateTime.Now : ca,
            TablePrefix = tablePrefix
        };
    }

    public async Task AddAsync(AnalysisBatch batch)
    {
        var sql = $@"
            INSERT INTO {_tableName} (Name, PeriodStart, PeriodEnd, CreatedAt, TablePrefix)
            VALUES (@p0, @p1, @p2, @p3, @p4)";

        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", batch.Name),
            new SqliteParameter("@p1", batch.PeriodStart.ToString("o")),
            new SqliteParameter("@p2", batch.PeriodEnd.ToString("o")),
            new SqliteParameter("@p3", batch.CreatedAt.ToString("o")),
            new SqliteParameter("@p4", batch.TablePrefix ?? string.Empty)
        };

        await _databaseService.ExecuteNonQueryAsync(sql, parameters);
    }

    public async Task UpdateAsync(AnalysisBatch batch)
    {
        var sql = $@"
            UPDATE {_tableName}
            SET Name = @p0,
                PeriodStart = @p1,
                PeriodEnd = @p2,
                CreatedAt = @p3,
                TablePrefix = @p4
            WHERE Id = @p5";

        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", batch.Name),
            new SqliteParameter("@p1", batch.PeriodStart.ToString("o")),
            new SqliteParameter("@p2", batch.PeriodEnd.ToString("o")),
            new SqliteParameter("@p3", DateTime.Now.ToString("o")),
            new SqliteParameter("@p4", batch.TablePrefix ?? string.Empty),
            new SqliteParameter("@p5", batch.Id)
        };

        await _databaseService.ExecuteNonQueryAsync(sql, parameters);
    }

    public async Task DeleteAsync(int id)
    {
        // ✅ 修复: 使用参数化查询
        var sql = $"DELETE FROM {_tableName} WHERE Id = @p0";
        var parameters = new List<SqliteParameter>
        {
            new SqliteParameter("@p0", id)
        };
        await _databaseService.ExecuteNonQueryAsync(sql, parameters);
    }

    /// <summary>
    /// 替换SQL语句中的参数
    /// </summary>
    /// <param name="sql">原始SQL语句</param>
    /// <param name="batch">分析批次信息</param>
    /// <returns>替换后的SQL语句</returns>
    public string ReplaceSqlParameters(string sql, AnalysisBatch batch)
    {
        if (string.IsNullOrWhiteSpace(sql))
            return sql;

        var result = sql;
        
        // 替换表名前缀 @FXPC
        if (!string.IsNullOrWhiteSpace(batch.TablePrefix))
        {
            result = result.Replace("@FXPC", batch.TablePrefix);
        }
        
        // 替换期间开始 @FXQQ
        result = result.Replace("@FXQQ", $"'{batch.PeriodStart:yyyy-MM-dd}'");
        
        // 替换期间结束 @FXQZ
        // 注意，这里给截止日志增加1天，确保可以正确扫描到截止当天的数据
        result = result.Replace("@FXQZ", $"'{batch.PeriodEnd.AddDays(1):yyyy-MM-dd}'");
        
        return result;
    }

    /// <summary>
    /// 执行扫描：对选中的指标执行SQL扫描
    /// </summary>
    public async Task<ScanResult> ScanAsync(
        AnalysisBatch batch, 
        List<Indicator> indicators, 
        Action<int, string>? progressCallback = null,
        bool saveResult = true)
    {
        var result = new ScanResult
        {
            BatchId = batch.Id,
            BatchName = batch.Name,
            StartTime = DateTime.Now
        };

        if (indicators == null || indicators.Count == 0)
        {
            result.ErrorMessage = "没有选择任何指标";
            result.EndTime = DateTime.Now;
            return result;
        }

        // 在扫描前，删除该批次的旧扫描记录
        if (saveResult)
        {
            try
            {
                var scanResultService = new ScanResultService();
                var deletedCount = await scanResultService.DeleteResultsByBatchIdAsync(batch.Id);
            }
            catch
            {
                
            }
        }

        int total = indicators.Count;
        int completed = 0;

        foreach (var indicator in indicators)
        {
            try
            {
                // 替换SQL参数
                var sql = ReplaceSqlParameters(indicator.SqlStatement, batch);
                
                // 执行SQL
                var data = await _databaseService.ExecuteQueryAsync(sql, new List<object>());
                
                // 记录结果
                result.Results.Add(new ScanItemResult
                {
                    IndicatorId = indicator.Id,
                    IndicatorName = indicator.Name,
                    SqlStatement = sql,
                    RowCount = data?.Count > 1 ? data.Count - 1 : 0,
                    Data = data,
                    IsSuccess = true
                });

                completed++;
                progressCallback?.Invoke(
                    (int)((double)completed / total * 100), 
                    $"已完成: {indicator.Name} ({completed}/{total})"
                );
            }
            catch (Exception ex)
            {
                result.Results.Add(new ScanItemResult
                {
                    IndicatorId = indicator.Id,
                    IndicatorName = indicator.Name,
                    SqlStatement = indicator.SqlStatement,
                    ErrorMessage = ex.Message,
                    IsSuccess = false
                });

                progressCallback?.Invoke(
                    (int)((double)completed / total * 100), 
                    $"执行失败: {indicator.Name} - {ex.Message}"
                );
            }
        }

        result.EndTime = DateTime.Now;
        result.SuccessCount = result.Results.Count(r => r.IsSuccess);
        result.FailCount = result.Results.Count(r => !r.IsSuccess);

        // 保存扫描结果到数据库（包括成功和失败的，确保用户能看到完整信息）
        if (saveResult && result.Results.Count > 0)
        {
            try
            {
                var scanResultService = new ScanResultService();
                var savedCount = await scanResultService.SaveScanResultAsync(result);
                result.SavedCount = savedCount;
                result.SaveSuccess = savedCount >= result.Results.Count;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"保存结果失败: {ex.Message}";
            }
        }
        else if (result.Results.Count == 0)
        {
            result.ErrorMessage = "没有可保存的扫描结果（所有指标均执行失败或数据为空）";
        }

        return result;
    }

    public static bool IsDefaultBatch(int id)
    {
        return id == 1;
    }

    public static bool IsValidTablePrefix(string prefix)
    {
        if (string.IsNullOrWhiteSpace(prefix))
            return true;

        if (!char.IsLetter(prefix[0]))
            return false;

        foreach (char c in prefix)
        {
            if (!char.IsLetterOrDigit(c) && c != '_')
                return false;
        }

        return true;
    }
}
    /// <summary>
    /// 扫描结果
    /// </summary>
    public class ScanResult
    {
        public int BatchId { get; set; }
        public string BatchName { get; set; } = string.Empty;
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public int SuccessCount { get; set; }
        public int FailCount { get; set; }
        public int SavedCount { get; set; }
        public bool SaveSuccess { get; set; }
        public string? ErrorMessage { get; set; }
        public List<ScanItemResult> Results { get; set; } = new();

        public TimeSpan Duration => EndTime - StartTime;
        public int TotalCount => SuccessCount + FailCount;
        public bool IsSuccess => string.IsNullOrEmpty(ErrorMessage) && FailCount == 0;
    }

    /// <summary>
    /// 扫描项结果
    /// </summary>
    public class ScanItemResult
    {
        public string IndicatorId { get; set; } = string.Empty;
        public string IndicatorName { get; set; } = string.Empty;
        public string SqlStatement { get; set; } = string.Empty;
        public int RowCount { get; set; }
        public List<List<object>>? Data { get; set; } // 临时存储，不保存到数据库
        public string? ErrorMessage { get; set; }
        public bool IsSuccess { get; set; }
    }
