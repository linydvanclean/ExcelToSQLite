using ReactiveUI;
using System;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.Models;

public class Indicator : ReactiveObject
{
    private string _id = string.Empty;
    private string _name = string.Empty;
    private string _sqlStatement = string.Empty;
    private string _sqlDetailData = string.Empty;
    private string _description = string.Empty;
    private string _category = string.Empty;
    private DateTime _createdAt;
    private DateTime _updatedAt;
    private string _createdBy = "admin";
    private bool _isActive = true;
    private int _index;
    private bool _isSelected;
    
    public static event Action<Indicator, bool>? SelectionChanged;

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

    public string SqlStatement
    {
        get => _sqlStatement;
        set => this.RaiseAndSetIfChanged(ref _sqlStatement, value);
    }

    public string SqlDetailData
    {
        get => _sqlDetailData;
        set => this.RaiseAndSetIfChanged(ref _sqlDetailData, value);
    }

    public string Description
    {
        get => _description;
        set => this.RaiseAndSetIfChanged(ref _description, value);
    }

    public string Category
    {
        get => _category;
        set => this.RaiseAndSetIfChanged(ref _category, value);
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

    public string CreatedBy
    {
        get => _createdBy;
        set => this.RaiseAndSetIfChanged(ref _createdBy, value);
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
        set
        {
            this.RaiseAndSetIfChanged(ref _isSelected, value);
            if (SelectionChanged != null)
            {
                var handler = SelectionChanged;
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    handler?.Invoke(this, value);
                }).ConfigureAwait(false);
            }
        }
    }

    public string CreatedAtDisplay => CreatedAt.ToString("yyyy-MM-dd HH:mm");
}