using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
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
using ExcelToSQLite.Views;
using ExcelToSQLite.Helpers;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 扫描结果 ViewModel - 实现 ICleanupPage 和 IDisposable 以确保资源正确释放
/// </summary>
public class ScanResultViewModel : ReactiveObject, ICleanupPage, IDisposable
{
    private readonly AnalysisBatchService? _batchService;
    private readonly IndicatorService? _indicatorService;
    private readonly ScanResultService? _scanResultService;
    private readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
    private readonly CompositeDisposable _subscriptions = new CompositeDisposable();
    private CancellationTokenSource? _cts;
    private Window? _parentWindow;
    private bool _disposed;
    private bool _isCleaned;
    private bool _isLoading;
    private bool _showStatus;

    private ObservableCollection<AnalysisBatch> _batches = new();
    private AnalysisBatch? _selectedBatch;
    private ObservableCollection<ScanResultItem> _scanResults = new();
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Colors.Green);

    public ScanResultViewModel()
    {
        try
        {

            _batchService = new AnalysisBatchService();
            _indicatorService = new IndicatorService();
            _scanResultService = new ScanResultService();

            // 初始化命令
            LoadBatchesCommand = ReactiveCommand.CreateFromTask(LoadBatchesAsync);
            LoadResultsCommand = ReactiveCommand.CreateFromTask<AnalysisBatch>(LoadResultsAsync);
            ViewDetailCommand = ReactiveCommand.CreateFromTask<ScanResultItem>(ViewDetailAsync);
            RefreshCommand = ReactiveCommand.CreateFromTask(RefreshAsync);

            // ✅ 订阅命令异常
            LoadBatchesCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex =>
                {
                    SetStatus($"加载批次失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                })
                .DisposeWith(_subscriptions);

            RefreshCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex =>
                {
                    SetStatus($"刷新失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                })
                .DisposeWith(_subscriptions);

            // ✅ 启动异步初始化
            _ = InitializeAsync();

        }
        catch (Exception ex)
        {
            SetStatusSafe($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #region 属性

    public bool IsBusy => IsLoading;

    public ObservableCollection<AnalysisBatch> Batches
    {
        get => _batches;
        set => this.RaiseAndSetIfChanged(ref _batches, value);
    }

    public AnalysisBatch? SelectedBatch
    {
        get => _selectedBatch;
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedBatch, value);
            if (value != null && !_disposed && !_isCleaned)
            {
                _ = LoadResultsSafeAsync(value);
            }
        }
    }

    public ObservableCollection<ScanResultItem> ScanResults
    {
        get => _scanResults;
        set => this.RaiseAndSetIfChanged(ref _scanResults, value);
    }

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

    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> LoadBatchesCommand { get; } = null!;
    public ReactiveCommand<AnalysisBatch, Unit> LoadResultsCommand { get; } = null!;
    public ReactiveCommand<ScanResultItem, Unit> ViewDetailCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> RefreshCommand { get; } = null!;

    #endregion

    #region 公共方法

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    #endregion

    #region 初始化

    private async Task InitializeAsync()
    {
        if (_disposed || _isCleaned) return;

        await _lock.WaitAsync();
        try
        {
            if (_disposed || _isCleaned) return;

            await LoadBatchesAsync();
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

    #endregion

    #region 数据加载

    private async Task LoadBatchesAsync()
    {
        if (_disposed || _isCleaned) return;
        if (_batchService == null) return;

        // ✅ 取消之前的操作
        CancelCurrentOperation();

        try
        {
            await SetLoadingAsync(true, "正在加载批次...");

            // ✅ 使用新的取消令牌
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            var list = await _batchService.GetAllAsync(1000);

            if (token.IsCancellationRequested || _disposed || _isCleaned) return;

            var sortedList = list.OrderByDescending(b => b.CreatedAt).ToList();

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;

                Batches = new ObservableCollection<AnalysisBatch>(sortedList);
                if (Batches.Count > 0)
                {
                    SelectedBatch = Batches.First();
                }
                SetStatus($"加载完成，共 {Batches.Count} 个批次", new SolidColorBrush(Colors.Green));
                IsLoading = false;
            });
        }
        catch (OperationCanceledException)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed && !_isCleaned)
                {
                    IsLoading = false;
                    SetStatus("操作已取消", new SolidColorBrush(Colors.Orange));
                }
            });
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"加载批次失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await SetLoadingAsync(false);
        }
    }

    private async Task LoadResultsAsync(AnalysisBatch batch)
    {
        if (_disposed || _isCleaned) return;
        if (batch == null) return;
        if (_scanResultService == null || _indicatorService == null) return;

        // ✅ 取消之前的操作
        CancelCurrentOperation();

        try
        {
            await SetLoadingAsync(true, $"正在加载批次 '{batch.Name}' 的扫描结果...");

            // ✅ 使用新的取消令牌
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            var records = await _scanResultService.GetResultsByBatchIdAsync(batch.Id);

            if (token.IsCancellationRequested || _disposed || _isCleaned) return;

            if (records.Count == 0)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    ScanResults = new ObservableCollection<ScanResultItem>();
                    SetStatus($"批次 '{batch.Name}' 暂无扫描结果，请先执行扫描", new SolidColorBrush(Colors.Orange));
                    IsLoading = false;
                });
                return;
            }

            var indicators = await _indicatorService.GetAllAsync();

            if (token.IsCancellationRequested || _disposed || _isCleaned) return;

            var indicatorDict = indicators.ToDictionary(i => i.Id, i => i);

            var items = new List<ScanResultItem>();
            int index = 1;  // ✅ 序号从1开始
            foreach (var record in records)
            {
                if (token.IsCancellationRequested) break;

                var indicator = indicatorDict.GetValueOrDefault(record.IndicatorId);
                var item = new ScanResultItem
                {
                    RecordId = record.Id,
                    BatchId = record.BatchId,
                    IndicatorId = record.IndicatorId,
                    IndicatorName = record.IndicatorName,
                    Category = indicator?.Category ?? "未分类",
                    RowCount = record.RowCount,
                    Status = record.Status,
                    SqlStatement = record.SqlStatement,
                    Index = index++  // ✅ 设置序号
                };
                items.Add(item);
            }

            if (token.IsCancellationRequested || _disposed || _isCleaned) return;

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                ScanResults = new ObservableCollection<ScanResultItem>(items);
                SetStatus($"加载完成，共 {ScanResults.Count} 条扫描结果", new SolidColorBrush(Colors.Green));
                IsLoading = false;
            });
        }
        catch (OperationCanceledException)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed && !_isCleaned)
                {
                    IsLoading = false;
                    SetStatus("操作已取消", new SolidColorBrush(Colors.Orange));
                }
            });
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"加载扫描结果失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await SetLoadingAsync(false);
        }
    }

    private async Task LoadResultsSafeAsync(AnalysisBatch batch)
    {
        if (_disposed || _isCleaned) return;

        try
        {
            await LoadResultsAsync(batch);
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"加载结果失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #endregion

    #region 查看详情

    private async Task ViewDetailAsync(ScanResultItem item)
    {
        if (_disposed || _isCleaned) return;
        if (item == null)
        {
            await ShowMessageAsync("请选择要查看的扫描结果");
            return;
        }

        if (item.Status != "Success")
        {
            await ShowMessageAsync($"该指标扫描失败，无法查看详情\n状态: {item.Status}");
            return;
        }

        if (string.IsNullOrWhiteSpace(item.SqlStatement))
        {
            await ShowMessageAsync("SQL语句为空，无法查看详情");
            return;
        }

        try
        {
            await SetStatusAsync("正在加载详情...", new SolidColorBrush(Colors.Blue));

            var detailViewModel = new ScanResultDetailViewModel(
                item.IndicatorId,
                item.SqlStatement,
                item.IndicatorName,
                item.Category
            );

            if (_parentWindow != null)
            {
                detailViewModel.SetParentWindow(_parentWindow);
            }

            if (SelectedBatch != null)
            {
                var periodQx = $"{SelectedBatch.PeriodStart:yyyyMMdd}-{SelectedBatch.PeriodEnd:yyyyMMdd}";

                detailViewModel.SetBatchInfo(
                    SelectedBatch.Id.ToString(),
                    SelectedBatch.PeriodStart.ToString("yyyy-MM-dd"),
                    SelectedBatch.PeriodEnd.ToString("yyyy-MM-dd"),
                    periodQx,
                    SelectedBatch.TablePrefix
                );

            }

            if (_indicatorService != null)
            {
                try
                {
                    var indicator = await _indicatorService.GetByIdAsync(item.IndicatorId);
                    if (indicator != null && !string.IsNullOrEmpty(indicator.SqlDetailData))
                    {
                        detailViewModel.SetIndicatorInfo(item.IndicatorId, indicator.SqlDetailData);
                    }
                }
                catch
                {
                }
            }

            var view = new ScanResultDetailView();
            view.DataContext = detailViewModel;

            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                if (_disposed || _isCleaned) return;

                var dialog = new Window
                {
                    Title = $"扫描结果详情 - {item.IndicatorName}",
                    Width = 900,
                    Height = 650,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = true,
                    MinWidth = 700,
                    MinHeight = 500,
                    Content = view,
                    Icon = IconHelper.GetAppIcon()
                };

                var owner = _parentWindow ?? GetMainWindow();
                if (owner == null)
                {
                    await ShowMessageAsync("无法获取父窗口，请重试");
                    return;
                }

                await dialog.ShowDialog(owner);
            });

            await SetStatusAsync("详情已关闭", new SolidColorBrush(Colors.Green));
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"加载详情失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"加载详情失败: {ex.Message}");
        }
    }

    #endregion

    #region 刷新

    private async Task RefreshAsync()
    {
        if (_disposed || _isCleaned) return;

        try
        {
            await SetStatusAsync("正在刷新...", new SolidColorBrush(Colors.Blue));
            await LoadBatchesAsync();

            if (!_disposed && !_isCleaned && SelectedBatch != null)
            {
                await LoadResultsAsync(SelectedBatch);
            }

            if (!_disposed && !_isCleaned)
            {
                await SetStatusAsync("刷新完成", new SolidColorBrush(Colors.Green));
            }
        }
        catch (Exception ex)
        {
            if (!_disposed && !_isCleaned)
            {
                await SetStatusAsync($"刷新失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            }
        }
    }

    #endregion

    #region 取消操作

    private void CancelCurrentOperation()
    {
        try
        {
            // ✅ 安全取消
            if (_cts != null)
            {
                if (!_cts.IsCancellationRequested)
                {
                    _cts.Cancel();
                }
                _cts.Dispose();
                _cts = null;
            }
        }
        catch
        {
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
        StatusColor = color;
        ShowStatus = !string.IsNullOrEmpty(message);
    }

    private void SetStatusSafe(string message, IBrush color)
    {
        _ = ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            SetStatus(message, color);
        });
    }

    private async Task SetStatusAsync(string message, IBrush color)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            SetStatus(message, color);
        });
    }

    private async Task SetLoadingAsync(bool isLoading, string? message = null)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsLoading = isLoading;
            if (message != null)
            {
                SetStatus(message, new SolidColorBrush(Colors.Orange));
            }
        });
    }

    #endregion

    #region 对话框

    private Window? GetMainWindow()
    {
        try
        {
            if (Avalonia.Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
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

    private async Task ShowMessageAsync(string msg)
    {
        if (_disposed || _isCleaned) return;

        var window = _parentWindow ?? GetMainWindow();
        if (window != null)
        {
            await MessageBox.ShowAsync(window, msg, "提示", MessageBoxButtons.OK);
        }
        else
        {
            await MessageBox.ShowAsync(msg, "提示", MessageBoxButtons.OK);
        }
    }

    #endregion

    #region ICleanupPage 实现

    /// <summary>
    /// 清理资源 - 在 View 卸载时调用
    /// </summary>
    public void Cleanup()
    {
        if (_isCleaned) return;

        try
        {
            // ✅ 1. 取消所有正在运行的操作
            CancelCurrentOperation();

            // ✅ 2. 取消所有订阅（使用 safe dispose）
            try
            {
                _subscriptions?.Dispose();
            }
            catch
            {
            }

            // ✅ 3. 清空集合（在 UI 线程）
            Dispatcher.UIThread.Post(() =>
            {
                try
                {
                    Batches?.Clear();
                    ScanResults?.Clear();
                    SelectedBatch = null;
                    StatusMessage = string.Empty;
                    ShowStatus = false;
                    IsLoading = false;
                }
                catch
                {
                }
            }, DispatcherPriority.Background);

            _isCleaned = true;
        }
        catch
        {
            _isCleaned = true;
        }
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            // ✅ 先清理
            Cleanup();

            // ✅ 释放锁（带超时）
            try
            {
                if (_lock != null)
                {
                    // 尝试获取锁，如果获取不到，说明锁已被占用，直接释放
                    if (_lock.Wait(TimeSpan.FromMilliseconds(500)))
                    {
                        try
                        {
                            // 锁内不需要做任何事情
                        }
                        finally
                        {
                            _lock.Release();
                        }
                    }
                    // 释放 SemaphoreSlim
                    _lock.Dispose();
                }
            }
            catch
            {
            }

            // ✅ 释放取消令牌（已在 Cleanup 中处理）
            _cts = null;

            // ✅ 释放命令
            try
            {
                LoadBatchesCommand?.Dispose();
                LoadResultsCommand?.Dispose();
                ViewDetailCommand?.Dispose();
                RefreshCommand?.Dispose();
            }
            catch
            {
            }

            _disposed = true;
        }
        catch
        {
            _disposed = true;
        }

        GC.SuppressFinalize(this);
    }

    #endregion
}