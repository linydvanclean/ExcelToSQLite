using ReactiveUI;
using System;

namespace ExcelToSQLite.Models;

public class DataDictionary : ReactiveObject
{
    private string _id = string.Empty;
    private string _name = string.Empty;
    private string _tableName = string.Empty;
    private string _description = string.Empty;
    private string _createdBy = "admin";
    private DateTime _createdAt;
    private DateTime _updatedAt;
    private bool _isActive = true;
    private int _index;
    private bool _isSelected;

    public string Id
    {
        get => _id;
        set => this.RaiseAndSetIfChanged(ref _id, value);
    }

    public string Name
    {
        get => _name;
        set => this.RaiseAndSetIfChanged(ref _name, value);
    }

    public string TableName
    {
        get => _tableName;
        set => this.RaiseAndSetIfChanged(ref _tableName, value);
    }

    public string Description
    {
        get => _description;
        set => this.RaiseAndSetIfChanged(ref _description, value);
    }
    public string CreatedBy
    {
        get => _createdBy;
        set => this.RaiseAndSetIfChanged(ref _createdBy, value);
    }

    public DateTime CreatedAt
    {
        get => _createdAt;
        set => this.RaiseAndSetIfChanged(ref _createdAt, value);
    }

    public DateTime UpdatedAt
    {
        get => _updatedAt;
        set => this.RaiseAndSetIfChanged(ref _updatedAt, value);
    }

    public bool IsActive
    {
        get => _isActive;
        set => this.RaiseAndSetIfChanged(ref _isActive, value);
    }

    public int Index
    {
        get => _index;
        set => this.RaiseAndSetIfChanged(ref _index, value);
    }

    public bool IsSelected
    {
        get => _isSelected;
        set => this.RaiseAndSetIfChanged(ref _isSelected, value);
    }

    public string CreatedAtDisplay => CreatedAt.ToString("yyyy-MM-dd HH:mm");
}
