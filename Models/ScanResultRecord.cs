using ReactiveUI;
using System;

namespace ExcelToSQLite.Models;

/// <summary>
/// 扫描结果记录
/// </summary>
public class ScanResultRecord : ReactiveObject
{
    private int _id;
    private int _batchId;
    private string _batchName = string.Empty;
    private string _indicatorId = string.Empty;
    private string _indicatorName = string.Empty;
    private int _rowCount;
    private string _status = string.Empty; // Success, Failed, Error
    private string _errorMessage = string.Empty;
    private string _sqlStatement = string.Empty;
    private DateTime _scanTime;
    private string _duration = string.Empty;
    private int _index;

    public int Id
    {
        get => _id;
        set => this.RaiseAndSetIfChanged(ref _id, value);
    }

    public int BatchId
    {
        get => _batchId;
        set => this.RaiseAndSetIfChanged(ref _batchId, value);
    }

    public string BatchName
    {
        get => _batchName;
        set => this.RaiseAndSetIfChanged(ref _batchName, value);
    }

    public string IndicatorId
    {
        get => _indicatorId;
        set => this.RaiseAndSetIfChanged(ref _indicatorId, value);
    }

    public string IndicatorName
    {
        get => _indicatorName;
        set => this.RaiseAndSetIfChanged(ref _indicatorName, value);
    }

    public int RowCount
    {
        get => _rowCount;
        set => this.RaiseAndSetIfChanged(ref _rowCount, value);
    }

    public string Status
    {
        get => _status;
        set => this.RaiseAndSetIfChanged(ref _status, value);
    }

    public string ErrorMessage
    {
        get => _errorMessage;
        set => this.RaiseAndSetIfChanged(ref _errorMessage, value);
    }

    public string SqlStatement
    {
        get => _sqlStatement;
        set => this.RaiseAndSetIfChanged(ref _sqlStatement, value);
    }

    public DateTime ScanTime
    {
        get => _scanTime;
        set => this.RaiseAndSetIfChanged(ref _scanTime, value);
    }

    public string Duration
    {
        get => _duration;
        set => this.RaiseAndSetIfChanged(ref _duration, value);
    }

    public int Index
    {
        get => _index;
        set => this.RaiseAndSetIfChanged(ref _index, value);
    }

    public string StatusDisplay
    {
        get
        {
            return Status switch
            {
                "Success" => "✅ 成功",
                "Failed" => "❌ 失败",
                "Error" => "⚠️ 错误",
                _ => "❓ 未知"
            };
        }
    }

    public string StatusColor
    {
        get
        {
            return Status switch
            {
                "Success" => "#4CAF50",
                "Failed" => "#F44336",
                "Error" => "#FF9800",
                _ => "#9E9E9E"
            };
        }
    }

    public string ScanTimeDisplay => ScanTime.ToString("yyyy-MM-dd HH:mm:ss");
}