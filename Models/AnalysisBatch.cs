using ReactiveUI;
using System;

namespace ExcelToSQLite.Models;

public class AnalysisBatch : ReactiveObject
{
    private int _id;
    private string _name = string.Empty;
    private DateTime _periodStart;
    private DateTime _periodEnd;
    private DateTime _createdAt;
    private string _tablePrefix = string.Empty;
    private int _index;
    private bool _isScanning;
    private string _scanStatus = string.Empty;
    private int _scanProgress;

    public int Id
    {
        get => _id;
        set => this.RaiseAndSetIfChanged(ref _id, value);
    }

    public string Name
    {
        get => _name;
        set => this.RaiseAndSetIfChanged(ref _name, value);
    }

    public DateTime PeriodStart
    {
        get => _periodStart;
        set => this.RaiseAndSetIfChanged(ref _periodStart, value);
    }

    public DateTime PeriodEnd
    {
        get => _periodEnd;
        set => this.RaiseAndSetIfChanged(ref _periodEnd, value);
    }

    public DateTime CreatedAt
    {
        get => _createdAt;
        set => this.RaiseAndSetIfChanged(ref _createdAt, value);
    }

    public string TablePrefix
    {
        get => _tablePrefix;
        set => this.RaiseAndSetIfChanged(ref _tablePrefix, value);
    }

    public int Index
    {
        get => _index;
        set => this.RaiseAndSetIfChanged(ref _index, value);
    }

    public bool IsScanning
    {
        get => _isScanning;
        set => this.RaiseAndSetIfChanged(ref _isScanning, value);
    }

    public string ScanStatus
    {
        get => _scanStatus;
        set => this.RaiseAndSetIfChanged(ref _scanStatus, value);
    }

    public int ScanProgress
    {
        get => _scanProgress;
        set => this.RaiseAndSetIfChanged(ref _scanProgress, value);
    }
}