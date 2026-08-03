using ReactiveUI;

namespace ExcelToSQLite.Models;

public class ScanResultItem : ReactiveObject
{
    private string _category = string.Empty;
    private string _indicatorId = string.Empty;
    private string _indicatorName = string.Empty;
    private int _rowCount;
    private string _status = string.Empty;
    private string _sqlStatement = string.Empty;//已经将参数替换完成
    private int _recordId;
    private int _batchId;
    
    private int _index;
    public int Index
    {
        get => _index;
        set => this.RaiseAndSetIfChanged(ref _index, value);
    }

    public string Category
    {
        get => _category;
        set => this.RaiseAndSetIfChanged(ref _category, value);
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

    public string SqlStatement
    {
        get => _sqlStatement;
        set => this.RaiseAndSetIfChanged(ref _sqlStatement, value);
    }

    public int RecordId
    {
        get => _recordId;
        set => this.RaiseAndSetIfChanged(ref _recordId, value);
    }

    public int BatchId
    {
        get => _batchId;
        set => this.RaiseAndSetIfChanged(ref _batchId, value);
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
}