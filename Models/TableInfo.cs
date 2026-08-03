using System;
using Avalonia.Media;

namespace ExcelToSQLite.Models;

public class TableInfo
{
    private int _index;
    private string _tableName = string.Empty;
    private int _recordCount;
    private DateTime _createdAt;
    private bool _isRenaming;

    public int Index
    {
        get => _index;
        set => _index = value;
    }

    public string TableName
    {
        get => _tableName;
        set => _tableName = value;
    }

    public int RecordCount
    {
        get => _recordCount;
        set => _recordCount = value;
    }

    public DateTime CreatedAt
    {
        get => _createdAt;
        set => _createdAt = value;
    }

    public bool IsRenaming
    {
        get => _isRenaming;
        set => _isRenaming = value;
    }

    public IBrush RecordCountColor
    {
        get
        {
            if (RecordCount == 0)
                return new SolidColorBrush(Color.Parse("#78909C"));
            else if (RecordCount < 100)
                return new SolidColorBrush(Color.Parse("#4CAF50"));
            else if (RecordCount < 1000)
                return new SolidColorBrush(Color.Parse("#FF9800"));
            else
                return new SolidColorBrush(Color.Parse("#D32F2F"));
        }
    }
}