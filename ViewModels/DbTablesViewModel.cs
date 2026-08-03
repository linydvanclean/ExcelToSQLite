// ViewModels/DbTablesViewModel.cs
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ExcelToSQLite.Helpers;
using ExcelToSQLite.Views;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.Controls;

namespace ExcelToSQLite.ViewModels
{
    public class DbTablesViewModel : ReactiveObject, ICleanupPage, IDisposable
    {
        private readonly DatabaseService _databaseService;
        private readonly ExcelExportService _exportService;
        private readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
        private readonly CompositeDisposable _subscriptions = new CompositeDisposable();
        private CancellationTokenSource? _cts;
        private bool _disposed;
        private bool _isCleaned;
        private bool _isLoadingTables;
        private bool _isExporting;
        private bool _isPreviewing;
        private bool _showStatus;

        // ✅ 添加父窗口字段
        private Window? _parentWindow;

        private ObservableCollection<string> _tableNames = new();
        private string? _selectedTable;
        private string _searchText = string.Empty;
        private string _statusMessage = string.Empty;
        private IBrush _statusColor = new SolidColorBrush(Colors.Green);
        private ObservableCollection<string> _filteredTableNames = new();
        private int _totalCount;
        private string _tableInfo = string.Empty;

        #region 属性

        public ObservableCollection<string> TableNames
        {
            get => _tableNames;
            set => this.RaiseAndSetIfChanged(ref _tableNames, value);
        }

        public ObservableCollection<string> FilteredTableNames
        {
            get => _filteredTableNames;
            set => this.RaiseAndSetIfChanged(ref _filteredTableNames, value);
        }

        public string? SelectedTable
        {
            get => _selectedTable;
            set
            {
                this.RaiseAndSetIfChanged(ref _selectedTable, value);
                if (!string.IsNullOrEmpty(value))
                {
                    _ = LoadTableInfoAsync(value);
                }
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                this.RaiseAndSetIfChanged(ref _searchText, value);
                FilterTables();
            }
        }

        public bool IsLoadingTables
        {
            get => _isLoadingTables;
            set => this.RaiseAndSetIfChanged(ref _isLoadingTables, value);
        }

        public bool IsExporting
        {
            get => _isExporting;
            set => this.RaiseAndSetIfChanged(ref _isExporting, value);
        }

        public bool IsPreviewing
        {
            get => _isPreviewing;
            set => this.RaiseAndSetIfChanged(ref _isPreviewing, value);
        }

        public bool IsBusy => IsLoadingTables || IsExporting || IsPreviewing;

        public string StatusMessage
        {
            get => _statusMessage;
            set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
        }

        public IBrush StatusColor
        {
            get => _statusColor;
            set => this.RaiseAndSetIfChanged(ref _statusColor, value);
        }

        public bool ShowStatus
        {
            get => _showStatus;
            set => this.RaiseAndSetIfChanged(ref _showStatus, value);
        }

        public int TotalCount
        {
            get => _totalCount;
            set => this.RaiseAndSetIfChanged(ref _totalCount, value);
        }

        public string TableInfo
        {
            get => _tableInfo;
            set => this.RaiseAndSetIfChanged(ref _tableInfo, value);
        }

        #endregion

        public ReactiveCommand<Unit, Unit> RefreshCommand { get; }
        public ReactiveCommand<Unit, Unit> PreviewCommand { get; }
        public ReactiveCommand<Unit, Unit> ExportCommand { get; }

        public DbTablesViewModel() : this(DatabaseService.Instance, null) { }

        // ✅ 添加带父窗口的构造函数
        public DbTablesViewModel(Window? parentWindow) : this(DatabaseService.Instance, parentWindow) { }

        public DbTablesViewModel(DatabaseService databaseService, Window? parentWindow = null)
        {
            _databaseService = databaseService ?? throw new ArgumentNullException(nameof(databaseService));
            _exportService = new ExcelExportService();
            _parentWindow = parentWindow;  // ✅ 保存父窗口

            RefreshCommand = ReactiveCommand.CreateFromTask(RefreshTablesAsync);
            RefreshCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex => SetStatus($"刷新失败: {ex.Message}", new SolidColorBrush(Colors.Red)))
                .DisposeWith(_subscriptions);

            PreviewCommand = ReactiveCommand.CreateFromTask(PreviewDataAsync);
            PreviewCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex => SetStatus($"预览失败: {ex.Message}", new SolidColorBrush(Colors.Red)))
                .DisposeWith(_subscriptions);

            ExportCommand = ReactiveCommand.CreateFromTask(ExportDataAsync);
            ExportCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex => SetStatus($"导出失败: {ex.Message}", new SolidColorBrush(Colors.Red)))
                .DisposeWith(_subscriptions);

            Task.Run(async () => await InitializeAsync());
            
        }

        // ✅ 添加设置父窗口的方法
        public void SetParentWindow(Window? parentWindow)
        {
            _parentWindow = parentWindow;
        }

        private async Task InitializeAsync()
        {
            if (_disposed || _isCleaned) return;
            await _lock.WaitAsync();
            try
            {
                if (_disposed || _isCleaned) return;
                await RefreshTablesAsync();
            }
            catch (Exception ex)
            {
                await SetStatusAsync($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            }
            finally
            {
                _lock.Release();
            }
        }

        private async Task RefreshTablesAsync()
        {
            if (_disposed || _isCleaned) return;
            CancelCurrentOperation();

            try
            {
                await SetLoadingTablesAsync(true, "正在加载表列表...");
                _cts = new CancellationTokenSource();
                var token = _cts.Token;

                var tables = await _databaseService.GetAllTableNamesAsync();
                if (token.IsCancellationRequested || _disposed || _isCleaned) return;

                var userTables = tables
                    .Where(t => !t.StartsWith("sqlite_", StringComparison.OrdinalIgnoreCase))
                    .Where(t => !Models.TableNames.AllowedSet.Contains(t))
                    .OrderBy(t => t)
                    .ToList();

                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    TableNames = new ObservableCollection<string>(userTables);
                    FilteredTableNames = new ObservableCollection<string>(userTables);
                    SelectedTable = null;
                    TotalCount = 0;
                    TableInfo = string.Empty;
                    SearchText = string.Empty;
                    IsLoadingTables = false;
                    SetStatus($"加载完成，共 {userTables.Count} 个表", new SolidColorBrush(Colors.Green));
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (!_disposed && !_isCleaned)
                    {
                        SetStatus($"加载表失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                        IsLoadingTables = false;
                    }
                });
            }
        }

        private void FilterTables()
        {
            if (_disposed || _isCleaned) return;
            _ = ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                FilteredTableNames = string.IsNullOrWhiteSpace(SearchText)
                    ? new ObservableCollection<string>(TableNames)
                    : new ObservableCollection<string>(TableNames.Where(t => 
                        t.Contains(SearchText, StringComparison.OrdinalIgnoreCase)));
            });
        }

        private async Task LoadTableInfoAsync(string tableName)
        {
            if (_disposed || _isCleaned || string.IsNullOrEmpty(tableName)) return;
            
            try
            {
                var countSql = $"SELECT COUNT(*) FROM \"{tableName}\"";
                var countResult = await _databaseService.ExecuteQueryAsync(countSql, new List<object>());

                int totalCount = 0;
                if (countResult != null && countResult.Count > 1 && countResult[1] != null && countResult[1].Count > 0)
                {
                    int.TryParse(countResult[1][0]?.ToString(), out totalCount);
                }

                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    TotalCount = totalCount;
                    TableInfo = totalCount > 0 ? $"共 {totalCount:N0} 条记录" : "空表";
                });
            }
            catch
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (!_disposed && !_isCleaned)
                    {
                        TableInfo = "获取记录数失败";
                    }
                });
            }
        }

        // ✅ 按照你的工作方法重写的 PreviewDataAsync
        private async Task PreviewDataAsync()
        {
            try
            {
                if (_disposed || _isCleaned) return;
                
                if (string.IsNullOrEmpty(SelectedTable))
                {
                    await SetStatusAsync("请先选择要预览的表", new SolidColorBrush(Colors.Orange));
                    return;
                }

                var tableName = SelectedTable;
                var sql = $"SELECT * FROM \"{tableName}\" LIMIT 10000";

                // ✅ 优先使用传入的父窗口，否则使用 GetMainWindow()
                var window = _parentWindow ?? GetMainWindow();
                if (window == null) 
                {
                    await SetStatusAsync("无法获取父窗口", new SolidColorBrush(Colors.Red));
                    return;
                }

                // ✅ 如果窗口不可见或最小化，尝试恢复
                if (!window.IsVisible || window.WindowState == WindowState.Minimized) 
                {
                    // ✅ 在 UI 线程上恢复窗口
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        if (window.WindowState == WindowState.Minimized)
                        {
                            window.WindowState = WindowState.Normal;
                        }
                        window.Show();
                        window.Activate();
                    });
                    
                    // 等待窗口恢复
                    await Task.Delay(200);
                    
                    // 再次检查
                    if (!window.IsVisible || window.WindowState == WindowState.Minimized)
                    {
                        await SetStatusAsync("无法恢复父窗口，请手动恢复后重试", new SolidColorBrush(Colors.Orange));
                        return;
                    }
                }

                await SetPreviewingAsync(true, $"正在准备预览表 '{tableName}' 数据...");

                await ThreadingHelper.RunOnUIThreadAsync(async () =>
                {
                    try
                    {
                        if (_disposed || _isCleaned) return;
                        
                        // ✅ 再次检查窗口状态
                        if (!window.IsVisible || window.WindowState == WindowState.Minimized) 
                        {
                            await SetStatusAsync("父窗口不可见，请恢复窗口后重试", new SolidColorBrush(Colors.Orange));
                            return;
                        }

                        // ✅ 与你的工作方法完全一致
                        var previewViewModel = new DetailDataViewModel(
                            sql, 
                            tableName, 
                            "原始数据"
                        );
                        
                        var detailView = new DetailDataView 
                        { 
                            DataContext = previewViewModel 
                        };

                        previewViewModel.SetParentWindow(window);
                        detailView.SetParentWindow(window);

                        var previewWindow = new Window
                        {
                            Title = $"数据预览 - {tableName} (前10000条)",
                            Width = 1200,
                            Height = 700,
                            WindowStartupLocation = WindowStartupLocation.CenterOwner,
                            Content = detailView,
                            CanResize = true,
                            MinWidth = 800,
                            MinHeight = 500,
                            Icon = IconHelper.GetAppIcon()
                        };

                        // ✅ 使用 ShowDialog 显示模态窗口
                        await previewWindow.ShowDialog(window);
                        
                        await SetStatusAsync($"预览窗口已关闭", new SolidColorBrush(Colors.Green));
                    }
                    catch (Exception ex)
                    {
                        await SetStatusAsync($"预览失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                    }
                    finally
                    {
                        await SetPreviewingAsync(false, string.Empty);
                    }
                });
            }
            catch (Exception ex)
            {
                await SetStatusAsync($"预览失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                await SetPreviewingAsync(false, string.Empty);
            }
        }

        private async Task ExportDataAsync()
        {
            if (_disposed || _isCleaned || string.IsNullOrEmpty(SelectedTable) || IsExporting) return;

            try
            {
                await SetExportingAsync(true, $"正在导出表 '{SelectedTable}' 全部数据...");
                CancelCurrentOperation();

                var tableName = SelectedTable;

                var sql = $"SELECT * FROM \"{tableName}\"";
                var data = await _databaseService.ExecuteQueryAsync(sql, new List<object>());

                if (data == null || data.Count <= 1)
                {
                    await SetStatusAsync("没有数据可导出", new SolidColorBrush(Colors.Orange));
                    await ShowMessageAsync($"表 '{tableName}' 没有数据可导出");
                    return;
                }

                var columnNames = new List<string>();
                if (data[0] != null)
                {
                    for (int i = 0; i < data[0].Count; i++)
                    {
                        var colName = data[0][i]?.ToString() ?? $"列{i + 1}";
                        columnNames.Add(colName);
                    }
                }

                var items = new List<DataGridRow>();
                for (int i = 1; i < data.Count; i++)
                {
                    var row = data[i];
                    var item = new DataGridRow { Index = i };
                    for (int j = 0; j < row.Count && j < columnNames.Count; j++)
                    {
                        var value = row[j]?.ToString() ?? string.Empty;
                        item.SetValue(columnNames[j], value);
                    }
                    items.Add(item);
                }

                var fileName = $"{tableName}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                var tempFile = await _exportService.ExportDataGridRowToExcelAsync(items, columnNames, fileName);

                // ✅ 优先使用传入的父窗口
                var mainWindow = _parentWindow ?? GetMainWindow();
                if (mainWindow != null)
                {
                    var savePath = await FileDialogHelper.GetSaveFilePathAsync(
                        mainWindow,
                        $"导出表数据 - {tableName}",
                        fileName,
                        "xlsx",
                        "Excel 文件");

                    if (!string.IsNullOrEmpty(savePath))
                    {
                        System.IO.File.Copy(tempFile, savePath, true);
                        System.IO.File.Delete(tempFile);
                        await SetStatusAsync($"数据已导出到: {System.IO.Path.GetFileName(savePath)}", new SolidColorBrush(Colors.Green));
                        await ShowMessageAsync($"导出成功！\n共 {items.Count:N0} 条记录\n文件: {System.IO.Path.GetFileName(savePath)}");
                    }
                    else
                    {
                        System.IO.File.Delete(tempFile);
                        await SetStatusAsync("导出已取消", new SolidColorBrush(Colors.Orange));
                    }
                }
                else
                {
                    System.IO.File.Delete(tempFile);
                    await SetStatusAsync("无法获取主窗口，导出失败", new SolidColorBrush(Colors.Red));
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                await SetStatusAsync($"导出失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                await ShowMessageAsync($"导出失败: {ex.Message}");
            }
            finally
            {
                await SetExportingAsync(false, string.Empty);
            }
        }

        private void CancelCurrentOperation()
        {
            try
            {
                _cts?.Cancel();
                _cts?.Dispose();
                _cts = null;
            }
            catch { }
        }

        private void SetStatus(string message, IBrush color)
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.Post(() => SetStatus(message, color));
                return;
            }
            StatusMessage = message;
            StatusColor = color;
            ShowStatus = !string.IsNullOrEmpty(message);
        }

        private async Task SetStatusAsync(string message, IBrush color)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() => SetStatus(message, color));
        }

        private async Task SetLoadingTablesAsync(bool isLoading, string? message = null)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoadingTables = isLoading;
                if (message != null) SetStatus(message, new SolidColorBrush(Colors.Orange));
            });
        }

        private async Task SetPreviewingAsync(bool isPreviewing, string? message = null)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsPreviewing = isPreviewing;
                if (message != null) SetStatus(message, new SolidColorBrush(Colors.Orange));
            });
        }

        private async Task SetExportingAsync(bool isExporting, string? message = null)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = isExporting;
                if (message != null) SetStatus(message, new SolidColorBrush(Colors.Orange));
            });
        }

        private async Task ShowMessageAsync(string msg, string title = "提示")
        {
            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                try
                {
                    var mainWindow = _parentWindow ?? GetMainWindow();
                    if (mainWindow != null && mainWindow.IsVisible)
                    {
                        await MessageBox.ShowAsync(mainWindow, msg, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        await MessageBox.ShowAsync(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch
                {
                }
            });
        }

        private Window? GetMainWindow()
        {
            try
            {
                if (global::Avalonia.Application.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
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

        public void Cleanup()
        {
            if (_isCleaned) return;
            try
            {
                CancelCurrentOperation();
                _lock.Wait(TimeSpan.FromSeconds(2));
                try
                {
                    _subscriptions?.Dispose();
                    _ = ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        try
                        {
                            TableNames?.Clear();
                            FilteredTableNames?.Clear();
                            SelectedTable = null;
                            TableInfo = string.Empty;
                            TotalCount = 0;
                        }
                        catch { }
                    });
                    _isCleaned = true;
                }
                finally
                {
                    try { _lock.Release(); } catch { }
                }
            }
            catch
            {
                _isCleaned = true;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            Cleanup();
            _lock?.Dispose();
            _cts?.Dispose();
            RefreshCommand?.Dispose();
            PreviewCommand?.Dispose();
            ExportCommand?.Dispose();
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}