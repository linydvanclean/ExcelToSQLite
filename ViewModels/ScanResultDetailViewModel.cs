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
using ExcelToSQLite.Views;
using System.IO;
using ExcelToSQLite.Helpers;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels;

public class ScanResultDetailViewModel : ReactiveObject, IDisposable
{
    private readonly DatabaseService _databaseService;
    private readonly ExcelExportService _exportService;
    private readonly IndicatorService _indicatorService;
    private Window? _parentWindow;
    private bool _isDisposed;

    private string? _batchId;
    private string? _tablePrefix;
    private string? _periodQx;
    private string? _periodStart;
    private string? _periodEnd;
    private string? _indicatorId;
    private string? _sqlDetailData;

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

    #region 属性

    public bool IsBusy => IsLoading || IsExporting;

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

    public string? IndicatorId
    {
        get => _indicatorId;
        set => this.RaiseAndSetIfChanged(ref _indicatorId, value);
    }

    public string? SqlDetailData
    {
        get => _sqlDetailData;
        set
        {
            this.RaiseAndSetIfChanged(ref _sqlDetailData, value);
            this.RaisePropertyChanged(nameof(CanOpenDetailData));
        }
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
        set 
        {
            this.RaiseAndSetIfChanged(ref _isExporting, value);
            this.RaisePropertyChanged(nameof(CanExport));
        }
    }
    
    public bool CanOpenDetailData => !string.IsNullOrEmpty(SqlDetailData);
    public bool CanExport => !IsExporting && TableData != null && TableData.Count > 0;

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
            this.RaisePropertyChanged(nameof(CanExport));
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
    public ReactiveCommand<Unit, Unit> OpenDetailDataCommand { get; }
    public Action? OnClose { get; set; }

    #endregion

    public ScanResultDetailViewModel(string indicatorId, 
        string sqlStatement, 
        string indicatorName, 
        string category)
    {
        _databaseService = DatabaseService.Instance;
        _exportService = new ExcelExportService();
        _indicatorService = new IndicatorService();
        
        _indicatorId = indicatorId;
        _sqlStatement = sqlStatement;
        _indicatorName = indicatorName;
        _category = category;

        CloseCommand = ReactiveCommand.Create(Close);
        ExportCommand = ReactiveCommand.CreateFromTask(ExportDataAsync);
        OpenDetailDataCommand = ReactiveCommand.CreateFromTask(OpenDetailDataAsync,
            this.WhenAnyValue(x => x.CanOpenDetailData));

        _ = InitializeAsync();
    }

    #region 初始化

    private async Task InitializeAsync()
    {
        try
        {
            await LoadDataAsync();
            await LoadIndicatorAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private async Task LoadIndicatorAsync()
    {
        try
        {
            var indicator = await _indicatorService.GetByIdAsync(_indicatorId ?? string.Empty);

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (indicator != null)
                {
                    SqlDetailData = FormatSqlDetailData(indicator.SqlDetailData);
                    _category = indicator.Category;
                    this.RaisePropertyChanged(nameof(Category));
                }
                else
                {
                    _category = "未分类";
                    this.RaisePropertyChanged(nameof(Category));
                }
            });
        }
        catch (Exception ex)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                SetStatus($"指标明细加载失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            });
        }
    }

    #endregion

    #region 公共方法

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

    public void SetIndicatorInfo(string indicatorId, string sqlDetailData)
    {
        IndicatorId = indicatorId;
        SqlDetailData = sqlDetailData;
        this.RaisePropertyChanged(nameof(CanOpenDetailData));
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

    private async Task LoadDataAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
                SetStatus("正在加载数据...", new SolidColorBrush(Colors.Blue));
                ShowStatus = true;
                _queryTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                this.RaisePropertyChanged(nameof(QueryTime));
            });

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
                this.RaisePropertyChanged(nameof(ColumnHeaders));
                this.RaisePropertyChanged(nameof(TableData));
                this.RaisePropertyChanged(nameof(DisplayCount));

                var statusMsg = $"加载完成，共 {TotalCount} 条记录" +
                               (TotalCount > 10000 ? $" (显示前10000条)" : "");
                SetStatus(statusMsg, new SolidColorBrush(Colors.Green));
                IsLoading = false;
            });
        }
        catch (Exception ex)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                SetStatus($"加载数据失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                IsLoading = false;
            });
        }
    }

    private (List<string> columnNames, List<List<object>> dataRows) ExtractData(List<List<object>> data)
    {
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

    #region SQL格式化

    private string FormatSqlDetailData(string sqlDetailData)
    {
        if (string.IsNullOrWhiteSpace(sqlDetailData))
            return string.Empty;

        var result = sqlDetailData;
        result = result.Replace("@FXPC", TablePrefix ?? string.Empty);
        result = result.Replace("@FXQQ", $"'{PeriodStart}'");

        if (DateTime.TryParse(PeriodEnd, out var endDate))
        {
            var nextDay = endDate.AddDays(1);
            result = result.Replace("@FXQZ", $"'{nextDay:yyyy-MM-dd}'");
        }
        else
        {
            var defaultEnd = DateTime.Now;
            var nextDay = defaultEnd.AddDays(1);
            result = result.Replace("@FXQZ", $"'{nextDay:yyyy-MM-dd}'");
        }

        return result;
    }

    #endregion

    #region 导出

    private async Task ExportDataAsync()
    {
        if (TableData == null || TableData.Count == 0)
        {
            await ShowMessageAsync("没有数据可导出");
            return;
        }

        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = true;
                SetStatus("正在导出数据...", new SolidColorBrush(Colors.Blue));
                ShowStatus = true;
            });

            var fileName = $"{IndicatorName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            var tempFile = await _exportService.ExportDataGridRowToExcelAsync(
                TableData.ToList(),
                ColumnHeaders.ToList(),
                fileName
            );

            var savePath = await SaveFileDialogAsync($"导出数据 - {IndicatorName}", fileName);
            if (!string.IsNullOrEmpty(savePath))
            {
                File.Copy(tempFile, savePath, true);
                File.Delete(tempFile);
                await SetStatusAsync($"数据已导出到: {Path.GetFileName(savePath)}", new SolidColorBrush(Colors.Green));
                await ShowMessageAsync($"导出成功！\n文件: {Path.GetFileName(savePath)}");
            }
            else
            {
                File.Delete(tempFile);
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
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = false;
            });
        }
    }

    #endregion

    #region 打开详细数据

    public async Task OpenDetailDataAsync()
{
    if (string.IsNullOrEmpty(SqlDetailData))
    {
        await ShowMessageAsync("没有可用的明细查询语句");
        return;
    }

    try
    {
        await SetStatusAsync("正在准备详细数据...", new SolidColorBrush(Colors.Blue));

        // 处理 SQL 参数替换
        var sql = FormatSqlDetailData(SqlDetailData);


        var detailViewModel = new DetailDataViewModel(sql, $"{IndicatorName}_明细", Category);
        
        if (!string.IsNullOrEmpty(BatchId))
        {
            detailViewModel.SetBatchInfo(BatchId, PeriodStart ?? "", PeriodEnd ?? "", PeriodQx ?? "", TablePrefix ?? "");
        }

        // ✅ 保存窗口引用，以便在 OnClose 中关闭
        Window? detailWindow = null;

        // ✅ 设置 OnClose 回调 - 关闭当前窗口（DetailDataView 所在的窗口）
        detailViewModel.OnClose = () =>
        {
            if (detailWindow != null && detailWindow.IsVisible)
            {
                Dispatcher.UIThread.Post(() => detailWindow.Close());
            }
        };

        var detailView = new DetailDataView
        {
            DataContext = detailViewModel
        };

        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            var window = new Window
            {
                Title = $"详细数据预览 - {IndicatorName}",
                Width = 1200,
                Height = 700,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Content = detailView,
                CanResize = true,
                MinWidth = 800,
                MinHeight = 500,
                Icon = IconHelper.GetAppIcon()
            };

            // ✅ 保存窗口引用
            detailWindow = window;

            var parentWindow = _parentWindow ?? GetMainWindow();
            if (parentWindow != null)
            {
                detailView.SetParentWindow(parentWindow);
                detailViewModel.SetParentWindow(parentWindow);
            }
            else
            {
            }

            // ✅ 窗口关闭时清理资源
            window.Closed += (s, e) =>
            {
                detailViewModel?.Dispose();
                detailWindow = null;
            };

            if (parentWindow != null)
            {
                await window.ShowDialog(parentWindow);
            }
            else
            {
                window.Show();
            }
        });

        await SetStatusAsync("详细数据预览已关闭", new SolidColorBrush(Colors.Green));
    }
    catch (Exception ex)
    {
        await SetStatusAsync($"打开详细数据失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        await ShowMessageAsync($"打开详细数据失败: {ex.Message}");
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
    private void Close()
    {
        try
        {
            // 直接调用 OnClose 回调
            if (OnClose != null)
            {
                Dispatcher.UIThread.Post(() =>
                {
                    try
                    {
                        OnClose?.Invoke();
                    }
                    catch
                    {
                        // 如果 OnClose 失败，尝试直接关闭父窗口
                        TryCloseWindowDirectly();
                    }
                });
            }
            else
            {
                // 如果没有 OnClose，尝试直接关闭父窗口
                TryCloseWindowDirectly();
            }
        }
        catch
        {
        }
    }

    private void TryCloseWindowDirectly()
    {
        try
        {
            if (_parentWindow != null && _parentWindow.IsVisible)
            {
                Dispatcher.UIThread.Post(() => _parentWindow.Close());
            }
            else
            {
                // 尝试通过 VisualRoot 查找窗口
                var window = GetMainWindow();
                if (window != null && window.IsVisible)
                {
                    Dispatcher.UIThread.Post(() => window.Close());
                }
            }
        }
        catch
        {
        }
    }

    #endregion

    #region 对话框
    
    /// <summary>
    /// 显示提示消息（使用优化后的 MessageBox）
    /// </summary>
    private async Task ShowMessageAsync(string msg, string title = "提示", MessageBoxIcon icon = MessageBoxIcon.Information)
    {
        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            try
            {
                // ✅ 这里使用 await MessageBox.ShowAsync()
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
        if (window == null) return null;

        try
        {
            return await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var extension = Path.GetExtension(defaultFileName)?.TrimStart('.') ?? "xlsx";
                var fileNameWithoutExt = Path.GetFileNameWithoutExtension(defaultFileName);

                if (string.IsNullOrEmpty(fileNameWithoutExt))
                {
                    fileNameWithoutExt = $"导出数据_{DateTime.Now:yyyyMMdd_HHmmss}";
                }

                var filePath = await FileDialogHelper.GetSaveFilePathAsync(
                    window,
                    title,
                    fileNameWithoutExt,
                    extension,
                    extension.ToUpper() + " 文件"
                );

                return filePath;
            });
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"打开保存对话框失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            return null;
        }
    }

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
                OpenDetailDataCommand?.Dispose();
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