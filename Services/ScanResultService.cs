using ExcelToSQLite.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelToSQLite.Services;

public class ScanResultService
{
    private readonly DatabaseService _databaseService;
    private readonly string _tableName;

    public ScanResultService()
    {
        _databaseService = DatabaseService.Instance;
        _tableName = TableNames.ScanResults;
    }

    /// <summary>
    /// 保存扫描结果（保存元数据，包括异常记录）
    /// </summary>
    public async Task<int> SaveScanResultAsync(ScanResult scanResult)
    {
        if (scanResult == null)
            return 0;

        // 如果 Results 为空，检查是否有错误信息需要保存
        if (scanResult.Results == null || scanResult.Results.Count == 0)
        {
            if (!string.IsNullOrEmpty(scanResult.ErrorMessage))
            {
                return await SaveErrorRecordAsync(scanResult);
            }
            return 0;
        }

        int savedCount = 0;

        foreach (var item in scanResult.Results)
        {
            try
            {
                var insertSql = $@"
                    INSERT INTO {_tableName} 
                    (BatchId, BatchName, IndicatorId, IndicatorName, RowCount, Status, ErrorMessage, SqlStatement, ScanTime, Duration)
                    VALUES 
                    (@BatchId, @BatchName, @IndicatorId, @IndicatorName, @RowCount, @Status, @ErrorMessage, @SqlStatement, @ScanTime, @Duration)";

                var parameters = new List<SqliteParameter>
                {
                    new SqliteParameter("@BatchId", scanResult.BatchId),
                    new SqliteParameter("@BatchName", scanResult.BatchName ?? string.Empty),
                    new SqliteParameter("@IndicatorId", item.IndicatorId ?? string.Empty),
                    new SqliteParameter("@IndicatorName", item.IndicatorName ?? string.Empty),
                    new SqliteParameter("@RowCount", item.RowCount),
                    new SqliteParameter("@Status", item.IsSuccess ? "Success" : "Failed"),
                    new SqliteParameter("@ErrorMessage", item.ErrorMessage ?? string.Empty),
                    new SqliteParameter("@SqlStatement", item.SqlStatement ?? string.Empty),
                    new SqliteParameter("@ScanTime", scanResult.EndTime.ToString("o")),
                    new SqliteParameter("@Duration", scanResult.Duration.TotalSeconds.ToString("F2") + "s")
                };

                var rowsAffected = await _databaseService.ExecuteNonQueryAsync(insertSql, parameters);
                if (rowsAffected > 0)
                {
                    savedCount++;
                }
            }
            catch
            {
            }
        }

        return savedCount;
    }

    /// <summary>
    /// 保存错误记录（当整个扫描失败时）
    /// </summary>
    private async Task<int> SaveErrorRecordAsync(ScanResult scanResult)
    {
        try
        {
            var insertSql = $@"
                INSERT INTO {_tableName} 
                (BatchId, BatchName, IndicatorId, IndicatorName, RowCount, Status, ErrorMessage, SqlStatement, ScanTime, Duration)
                VALUES 
                (@BatchId, @BatchName, @IndicatorId, @IndicatorName, @RowCount, @Status, @ErrorMessage, @SqlStatement, @ScanTime, @Duration)";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", scanResult.BatchId),
                new SqliteParameter("@BatchName", scanResult.BatchName ?? string.Empty),
                new SqliteParameter("@IndicatorId", string.Empty),
                new SqliteParameter("@IndicatorName", "扫描失败"),
                new SqliteParameter("@RowCount", 0),
                new SqliteParameter("@Status", "Failed"),
                new SqliteParameter("@ErrorMessage", scanResult.ErrorMessage ?? "未知错误"),
                new SqliteParameter("@SqlStatement", string.Empty),
                new SqliteParameter("@ScanTime", scanResult.EndTime.ToString("o")),
                new SqliteParameter("@Duration", scanResult.Duration.TotalSeconds.ToString("F2") + "s")
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(insertSql, parameters);
            return rowsAffected > 0 ? 1 : 0;
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// 获取批次的所有扫描结果
    /// </summary>
    public async Task<List<ScanResultRecord>> GetResultsByBatchIdAsync(int batchId)
    {
        var result = new List<ScanResultRecord>();

        try
        {
            var sql = $@"
                SELECT Id, BatchId, BatchName, IndicatorId, IndicatorName, 
                    RowCount, Status, ErrorMessage, SqlStatement, ScanTime, Duration
                FROM {_tableName}
                WHERE BatchId = @BatchId
                ORDER BY ScanTime DESC, Id DESC";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", batchId)
            };
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (data == null || data.Count <= 1)
                return result;

            for (int i = 1; i < data.Count; i++)
            {
                var record = MapRowToScanResultRecord(data[i]);
                if (record != null)
                {
                    result.Add(record);
                }
            }
        }
        catch
        {
        }

        return result;
    }

    /// <summary>
    /// 获取指标的所有扫描结果
    /// </summary>
    public async Task<List<ScanResultRecord>> GetResultsByIndicatorIdAsync(string indicatorId)
    {
        var result = new List<ScanResultRecord>();

        try
        {
            var sql = $@"
                SELECT Id, BatchId, BatchName, IndicatorId, IndicatorName, 
                    RowCount, Status, ErrorMessage, SqlStatement, ScanTime, Duration
                FROM {_tableName}
                WHERE IndicatorId = @IndicatorId
                ORDER BY ScanTime DESC, Id DESC";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@IndicatorId", indicatorId)
            };
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (data == null || data.Count <= 1)
                return result;

            for (int i = 1; i < data.Count; i++)
            {
                var record = MapRowToScanResultRecord(data[i]);
                if (record != null)
                {
                    result.Add(record);
                }
            }
        }
        catch
        {
        }

        return result;
    }

    /// <summary>
    /// 根据记录ID获取扫描结果
    /// </summary>
    public async Task<ScanResultRecord?> GetResultByIdAsync(int recordId)
    {
        try
        {
            var sql = $@"
                SELECT Id, BatchId, BatchName, IndicatorId, IndicatorName, 
                    RowCount, Status, ErrorMessage, SqlStatement, ScanTime, Duration
                FROM {_tableName}
                WHERE Id = @RecordId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@RecordId", recordId)
            };
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (data == null || data.Count <= 1)
                return null;

            return MapRowToScanResultRecord(data[1]);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 执行扫描记录的SQL，获取详细数据（用于预览）
    /// </summary>
    public async Task<object?> ExecuteResultSqlAsync(int recordId)
    {
        var record = await GetResultByIdAsync(recordId);
        if (record == null || string.IsNullOrWhiteSpace(record.SqlStatement))
            return null;

        try
        {
            return await _databaseService.ExecuteQueryAsync(record.SqlStatement, new List<object>());
        }
        catch (Exception ex)
        {
            throw new Exception($"执行SQL失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 删除批次的扫描结果
    /// </summary>
    public async Task<int> DeleteResultsByBatchIdAsync(int batchId)
    {
        try
        {
            var sql = $@"
                DELETE FROM {_tableName}
                WHERE BatchId = @BatchId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", batchId)
            };

            return await _databaseService.ExecuteNonQueryAsync(sql, parameters);
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// 检查批次是否有扫描记录
    /// </summary>
    public async Task<bool> HasScanResultsAsync(int batchId)
    {
        try
        {
            var sql = $@"
                SELECT COUNT(*) FROM {_tableName}
                WHERE BatchId = @BatchId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", batchId)
            };
            var result = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (result != null && result.Count > 1 && result[1].Count > 0)
            {
                return Convert.ToInt32(result[1][0]) > 0;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 删除特定的扫描记录
    /// </summary>
    public async Task<bool> DeleteResultByIdAsync(int recordId)
    {
        try
        {
            var sql = $@"
                DELETE FROM {_tableName}
                WHERE Id = @RecordId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@RecordId", recordId)
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);
            return rowsAffected > 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 获取批次的最新扫描时间
    /// </summary>
    public async Task<DateTime?> GetLastScanTimeAsync(int batchId)
    {
        try
        {
            var sql = $@"
                SELECT MAX(ScanTime) FROM {_tableName}
                WHERE BatchId = @BatchId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", batchId)
            };
            var result = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (result != null && result.Count > 1 && result[1].Count > 0)
            {
                var value = result[1][0]?.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    return DateTime.Parse(value);
                }
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 获取扫描结果统计信息
    /// </summary>
    public async Task<ScanResultStatistics> GetStatisticsAsync(int batchId)
    {
        var stats = new ScanResultStatistics { BatchId = batchId };

        try
        {
            var sql = $@"
                SELECT 
                    COUNT(*) as TotalCount,
                    SUM(CASE WHEN Status = 'Success' THEN 1 ELSE 0 END) as SuccessCount,
                    SUM(CASE WHEN Status = 'Failed' THEN 1 ELSE 0 END) as FailedCount,
                    AVG(RowCount) as AvgRowCount,
                    MAX(RowCount) as MaxRowCount,
                    MIN(RowCount) as MinRowCount
                FROM {_tableName}
                WHERE BatchId = @BatchId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", batchId)
            };
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (data != null && data.Count > 1 && data[1].Count > 0)
            {
                var row = data[1];
                int idx = 0;
                stats.TotalCount = Convert.ToInt32(row[idx++]);
                stats.SuccessCount = Convert.ToInt32(row[idx++]);
                stats.FailedCount = Convert.ToInt32(row[idx++]);
                stats.AvgRowCount = row[idx] != null ? Convert.ToInt32(row[idx]) : 0; idx++;
                stats.MaxRowCount = row[idx] != null ? Convert.ToInt32(row[idx]) : 0; idx++;
                stats.MinRowCount = row[idx] != null ? Convert.ToInt32(row[idx]) : 0;
            }
        }
        catch
        {
        }

        return stats;
    }

    /// <summary>
    /// 获取批次的扫描结果摘要
    /// </summary>
    public async Task<BatchScanSummary> GetBatchSummaryAsync(int batchId)
    {
        var summary = new BatchScanSummary
        {
            BatchId = batchId
        };

        try
        {
            var sql = $@"
                SELECT 
                    COUNT(*) as TotalCount,
                    SUM(CASE WHEN Status = 'Success' THEN 1 ELSE 0 END) as SuccessCount,
                    SUM(CASE WHEN Status = 'Failed' THEN 1 ELSE 0 END) as FailedCount,
                    MAX(ScanTime) as LastScanTime
                FROM {_tableName}
                WHERE BatchId = @BatchId";

            var parameters = new List<SqliteParameter>
            {
                new SqliteParameter("@BatchId", batchId)
            };
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (data != null && data.Count > 1 && data[1].Count > 0)
            {
                var row = data[1];
                int idx = 0;
                summary.TotalIndicators = Convert.ToInt32(row[idx++]);
                summary.SuccessCount = Convert.ToInt32(row[idx++]);
                summary.FailedCount = Convert.ToInt32(row[idx++]);
                var lastScan = row[idx]?.ToString();
                if (!string.IsNullOrEmpty(lastScan))
                {
                    summary.LastScanTime = DateTime.Parse(lastScan);
                }
            }
        }
        catch
        {
        }

        return summary;
    }

    #region 私有方法

    private ScanResultRecord? MapRowToScanResultRecord(List<object> row)
    {
        try
        {
            int idx = 0;
            return new ScanResultRecord
            {
                Id = Convert.ToInt32(row[idx++]),
                BatchId = Convert.ToInt32(row[idx++]),
                BatchName = row[idx++]?.ToString() ?? string.Empty,
                IndicatorId = row[idx++]?.ToString() ?? string.Empty,
                IndicatorName = row[idx++]?.ToString() ?? string.Empty,
                RowCount = Convert.ToInt32(row[idx++]),
                Status = row[idx++]?.ToString() ?? string.Empty,
                ErrorMessage = row[idx++]?.ToString() ?? string.Empty,
                SqlStatement = row[idx++]?.ToString() ?? string.Empty,
                ScanTime = DateTime.Parse(row[idx++]?.ToString() ?? DateTime.Now.ToString("o")),
                Duration = row[idx++]?.ToString() ?? string.Empty
            };
        }
        catch
        {
            return null;
        }
    }

    #endregion
}

/// <summary>
/// 扫描结果统计
/// </summary>
public class ScanResultStatistics
{
    public int BatchId { get; set; }
    public int TotalCount { get; set; }
    public int SuccessCount { get; set; }
    public int FailedCount { get; set; }
    public int AvgRowCount { get; set; }
    public int MaxRowCount { get; set; }
    public int MinRowCount { get; set; }

    public double SuccessRate => TotalCount > 0 ? (double)SuccessCount / TotalCount * 100 : 0;
}

/// <summary>
/// 批次扫描摘要
/// </summary>
public class BatchScanSummary
{
    public int BatchId { get; set; }
    public int TotalIndicators { get; set; }
    public int SuccessCount { get; set; }
    public int FailedCount { get; set; }
    public DateTime? LastScanTime { get; set; }

    public double SuccessRate => TotalIndicators > 0 ? (double)SuccessCount / TotalIndicators * 100 : 0;
    public string LastScanTimeDisplay => LastScanTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "从未扫描";
}

/// <summary>
/// 扫描结果记录（用于显示）
/// </summary>
public class ScanResultRecord
{
    public int Id { get; set; }
    public int BatchId { get; set; }
    public string BatchName { get; set; } = string.Empty;
    public string IndicatorId { get; set; } = string.Empty;
    public string IndicatorName { get; set; } = string.Empty;
    public int RowCount { get; set; }
    public string Status { get; set; } = string.Empty;
    public string ErrorMessage { get; set; } = string.Empty;
    public string SqlStatement { get; set; } = string.Empty;
    public DateTime ScanTime { get; set; }
    public string Duration { get; set; } = string.Empty;

    public bool IsSuccess => Status == "Success";
    public string StatusDisplay => IsSuccess ? "✅ 成功" : "❌ 失败";
    public string RowCountDisplay => IsSuccess ? RowCount.ToString() : "-";
}