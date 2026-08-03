using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media;
using ExcelToSQLite.Services;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Threading.Tasks;
using ExcelToSQLite.Helpers;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels;

public class DetailDataViewModel : ReactiveObject, IDisposable
{
    private readonly DatabaseService _databaseService;
    private readonly ExcelExportService _exportService;
    private Window? _parentWindow;
    private string? _batchId;
    private string? _tablePrefix;
    private string? _periodQx;
    private string? _periodStart;
    private string? _periodEnd;
    private bool _isDisposed;

    private string _sqlStatement = string.Empty;
    private string _indicatorName = string.Empty;
    private string _category = string.Empty;
    private string _statusMessage = string.Empty;
    private IBrush _statusColorBrush = new SolidColorBrush(Colors.Green);
    private bool _isLoading;
    private int _totalCount;
    private string _queryTime = string.Empty;
    private ObservableCollection<DataGridRow> _tableData = new();
    private ObservableCollection<string> _columnHeaders = new();
    private bool _isExporting;
    private int _displayCount;
    private bool _showStatus;

    public event EventHandler? ColumnsUpdated;

    // ✅ 移除 Interaction，改用回调
    // public Interaction<Unit, Unit> CloseInteraction { get; } = new();

    public DetailDataViewModel(string sqlStatement, 
        string indicatorName,
        string category)
    {
        _databaseService = DatabaseService.Instance;
        _exportService = new ExcelExportService();
        _sqlStatement = sqlStatement;
        _indicatorName = indicatorName;
        _category = category;

        CloseCommand = ReactiveCommand.Create(Close);
        ExportCommand = ReactiveCommand.CreateFromTask(ExportDataAsync);

        _ = LoadDataSafeAsync();
    }

    #region 属性

    public string? BatchId
    {
        get => _batchId;
        set => this.RaiseAndSetIfChanged(ref _batchId, value);
    }

    public string? TablePrefix
    {
        get => _tablePrefix;
        set => this.RaiseAndSetIfChanged(ref _tablePrefix, value);
    }

    public string? PeriodStart
    {
        get => _periodStart;
        set => this.RaiseAndSetIfChanged(ref _periodStart, value);
    }

    public string? PeriodEnd
    {
        get => _periodEnd;
        set => this.RaiseAndSetIfChanged(ref _periodEnd, value);
    }

    public string? PeriodQx
    {
        get => _periodQx;
        set => this.RaiseAndSetIfChanged(ref _periodQx, value);
    }

    public bool IsExporting
    {
        get => _isExporting;
        set => this.RaiseAndSetIfChanged(ref _isExporting, value);
    }

    public bool IsBusy => IsLoading || IsExporting;

    public string IndicatorName => _indicatorName;
    public string Category => _category;
    public string QueryTime => _queryTime;

    public ObservableCollection<DataGridRow> TableData
    {
        get => _tableData;
        set
        {
            this.RaiseAndSetIfChanged(ref _tableData, value);
            OnColumnsUpdated();
            DisplayCount = _tableData?.Count ?? 0;
        }
    }

    public ObservableCollection<string> ColumnHeaders
    {
        get => _columnHeaders;
        set
        {
            this.RaiseAndSetIfChanged(ref _columnHeaders, value);
            OnColumnsUpdated();
        }
    }

    public int TotalCount
    {
        get => _totalCount;
        set => this.RaiseAndSetIfChanged(ref _totalCount, value);
    }

    public int DisplayCount
    {
        get => _displayCount;
        set => this.RaiseAndSetIfChanged(ref _displayCount, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
    }

    public string SqlStatement
    {
        get => _sqlStatement;
        set => this.RaiseAndSetIfChanged(ref _sqlStatement, value);
    }

    public IBrush StatusColorBrush
    {
        get => _statusColorBrush;
        set => this.RaiseAndSetIfChanged(ref _statusColorBrush, value);
    }

    public bool ShowStatus
    {
        get => _showStatus;
        set => this.RaiseAndSetIfChanged(ref _showStatus, value);
    }

    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> CloseCommand { get; }
    public ReactiveCommand<Unit, Unit> ExportCommand { get; }

    #endregion

    #region 公共方法

    public Action? OnClose { get; set; }

    public void SetBatchInfo(string batchId,
        string periodStart, string periodEnd, string periodQx,
        string tablePrefix)
    {
        BatchId = batchId;
        PeriodStart = periodStart;
        PeriodEnd = periodEnd;
        PeriodQx = periodQx;
        TablePrefix = tablePrefix;
    }

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    #endregion

    #region 事件

    protected virtual void OnColumnsUpdated()
    {
        ColumnsUpdated?.Invoke(this, EventArgs.Empty);
    }

    #endregion

    #region 数据加载

    private async Task LoadDataSafeAsync()
    {
        try
        {
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"加载数据失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private async Task LoadDataAsync()
    {
        try
        {
            await SetStatusAsync("正在加载详细数据...", new SolidColorBrush(Colors.Blue));
            await SetLoadingAsync(true);

            _queryTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            this.RaisePropertyChanged(nameof(QueryTime));

            if (string.IsNullOrWhiteSpace(_sqlStatement))
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    SetStatus("SQL语句为空，无法加载数据", new SolidColorBrush(Colors.Orange));
                    TableData = new ObservableCollection<DataGridRow>();
                    ColumnHeaders = new ObservableCollection<string>();
                    TotalCount = 0;
                    DisplayCount = 0;
                    IsLoading = false;
                });
                return;
            }

            var data = await _databaseService.ExecuteQueryAsync(_sqlStatement, new List<object>());

            if (data == null || data.Count <= 1)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    SetStatus("没有数据", new SolidColorBrush(Colors.Orange));
                    TableData = new ObservableCollection<DataGridRow>();
                    ColumnHeaders = new ObservableCollection<string>();
                    TotalCount = 0;
                    DisplayCount = 0;
                    IsLoading = false;
                });
                return;
            }

            // 提取列名和数据
            var (columnNames, dataRows) = await Task.Run(() => ExtractData(data));

            TotalCount = dataRows.Count;

            // 构建动态行数据
            var displayItems = await Task.Run(() => BuildDataRows(columnNames, dataRows, 10000));

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                ColumnHeaders = new ObservableCollection<string>(columnNames);
                TableData = new ObservableCollection<DataGridRow>(displayItems);
                DisplayCount = displayItems.Count;
                
                var statusMsg = $"加载完成，共 {TotalCount} 条记录" +
                               (TotalCount > 10000 ? $" (显示前10000条)" : "");
                SetStatus(statusMsg, new SolidColorBrush(Colors.Green));
                IsLoading = false;
            });
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"加载数据失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await SetLoadingAsync(false);
        }
        finally
        {
            await SetLoadingAsync(false);
        }
    }

    private (List<string> columnNames, List<List<object>> dataRows) ExtractData(List<List<object>> data)
    {
        // 提取列名
        var columnNames = new List<string>();
        if (data.Count > 0 && data[0] != null)
        {
            for (int i = 0; i < data[0].Count; i++)
            {
                var colName = data[0][i]?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(colName))
                {
                    colName = $"列{i + 1}";
                }
                columnNames.Add(colName);
            }
        }

        // 提取数据行
        var dataRows = new List<List<object>>();
        for (int i = 1; i < data.Count; i++)
        {
            if (data[i] != null)
            {
                dataRows.Add(data[i]);
            }
        }

        return (columnNames, dataRows);
    }

    private List<DataGridRow> BuildDataRows(List<string> columnNames, List<List<object>> dataRows, int maxDisplay)
    {
        var displayItems = new List<DataGridRow>();
        int count = Math.Min(maxDisplay, dataRows.Count);

        for (int i = 0; i < count; i++)
        {
            var row = dataRows[i];
            var item = new DataGridRow
            {
                Index = i + 1
            };

            for (int j = 0; j < row.Count && j < columnNames.Count; j++)
            {
                var columnName = columnNames[j];
                var value = row[j]?.ToString() ?? string.Empty;
                item.SetValue(columnName, value);
            }
            displayItems.Add(item);
        }

        return displayItems;
    }

    #endregion

    #region 导出

    private async Task ExportDataAsync()
    {
        if (IsExporting) return;

        try
        {
            await SetStatusAsync("正在生成导出数据...", new SolidColorBrush(Colors.Blue));
            await SetExportingAsync(true);

            if (string.IsNullOrWhiteSpace(_sqlStatement))
            {
                await ShowMessageAsync("没有可用的查询语句");
                return;
            }

            // 重新查询数据
            var data = await _databaseService.ExecuteQueryAsync(_sqlStatement, new List<object>());

            if (data == null || data.Count <= 1)
            {
                await ShowMessageAsync("没有数据可导出");
                await SetStatusAsync("没有数据", new SolidColorBrush(Colors.Orange));
                return;
            }

            // 提取列名和数据
            var (columnNames, dataRows) = await Task.Run(() => ExtractData(data));

            // 构建导出数据
            var items = await Task.Run(() => BuildDataRows(columnNames, dataRows, dataRows.Count));

            // 导出到 Excel
            var fileName = $"{IndicatorName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            var tempFile = await _exportService.ExportDataGridRowToExcelAsync(items, columnNames, fileName);

            // 弹出保存对话框
            var savePath = await SaveFileDialogAsync($"导出数据 - {IndicatorName}", fileName);
            
            if (!string.IsNullOrEmpty(savePath))
            {
                System.IO.File.Copy(tempFile, savePath, true);
                System.IO.File.Delete(tempFile);
                await SetStatusAsync($"数据已导出到: {System.IO.Path.GetFileName(savePath)}", new SolidColorBrush(Colors.Green));
                await ShowMessageAsync($"导出成功！\n共 {items.Count} 条记录\n文件: {System.IO.Path.GetFileName(savePath)}");
            }
            else
            {
                System.IO.File.Delete(tempFile);
                await SetStatusAsync("导出已取消", new SolidColorBrush(Colors.Orange));
            }
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"导出失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"导出失败: {ex.Message}");
        }
        finally
        {
            await SetExportingAsync(false);
        }
    }

    #endregion

    #region UI辅助方法

    private void SetStatus(string message, IBrush color)
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetStatus(message, color));
            return;
        }

        StatusMessage = message;
        StatusColorBrush = color;
        ShowStatus = !string.IsNullOrEmpty(message);
    }

    private async Task SetStatusAsync(string message, IBrush color)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            SetStatus(message, color);
        });
    }

    private async Task SetLoadingAsync(bool isLoading)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsLoading = isLoading;
        });
    }

    private async Task SetExportingAsync(bool isExporting)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsExporting = isExporting;
        });
    }
    
    private void Close()
    {
        try
        {
            // ✅ 只关闭当前预览窗口，不关闭父窗口
            CloseCurrentWindow();
        }
        catch
        {
        }
    }

    /// <summary>
    /// 关闭当前窗口（只关闭预览窗口）
    /// </summary>
    private void CloseCurrentWindow()
    {
        try
        {
            // ✅ 获取所有窗口
            if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var mainWindow = desktop.MainWindow;
            
                // ✅ 查找是预览窗口且不是主窗口的窗口
                foreach (var window in desktop.Windows)
                {
                    if (window.IsVisible && window != mainWindow)
                    {
                        // ✅ 检查窗口标题是否包含"预览"或"数据预览"
                        if (window.Title?.Contains("数据预览") == true || 
                            window.Title?.Contains("预览") == true)
                        {
                            Dispatcher.UIThread.Post(() => window.Close());
                            return;
                        }
                    }
                }
            
            }
        }
        catch
        {
        }
    }
    
    /// <summary>
    /// 尝试查找并关闭当前窗口
    /// </summary>
    private void TryFindAndCloseCurrentWindow()
    {
        try
        {
            // 获取所有窗口
            if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var mainWindow = desktop.MainWindow;
                
                // 查找不是主窗口的活动窗口
                foreach (var window in desktop.Windows)
                {
                    if (window.IsVisible && window != mainWindow)
                    {
                        Dispatcher.UIThread.Post(() => window.Close());
                        return;
                    }
                }
                
            }
        }
        catch
        {
        }
    }
    

    #endregion

    #region 对话框

    private async Task ShowMessageAsync(string msg, string title = "提示", MessageBoxIcon icon = MessageBoxIcon.Information)
    {
        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            try
            {
                if (_parentWindow != null && _parentWindow.IsVisible)
                {
                    await MessageBox.ShowAsync(_parentWindow, msg, title, MessageBoxButtons.OK, icon);
                }
                else
                {
                    var mainWindow = GetMainWindow();
                    if (mainWindow != null && mainWindow.IsVisible)
                    {
                        await MessageBox.ShowAsync(mainWindow, msg, title, MessageBoxButtons.OK, icon);
                    }
                    else
                    {
                        await MessageBox.ShowAsync(msg, title, MessageBoxButtons.OK, icon);
                    }
                }
            }
            catch
            {
            }
        });
    }

    private async Task<string?> SaveFileDialogAsync(string title, string defaultFileName)
    {        
        var window = _parentWindow ?? GetMainWindow();        
        if (window == null) 
        {
            return null;
        }

        try
        {
            return await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var result = await FileDialogHelper.GetSaveFilePathAsync(
                    window, 
                    title, 
                    defaultFileName, 
                    "xlsx", 
                    "Excel 文件"
                );
                return result;
            });
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"SaveFileDialogAsync 异常: {ex.Message}", new SolidColorBrush(Colors.Red));
            return null;
        }
    }

    #endregion

    #region 窗口获取

    private Window? GetMainWindow()
    {
        try
        {
            if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                return desktop.MainWindow;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                CloseCommand?.Dispose();
                ExportCommand?.Dispose();
            }
            catch
            {
            }
            _isDisposed = true;
        }
        GC.SuppressFinalize(this);
    }

    #endregion
}