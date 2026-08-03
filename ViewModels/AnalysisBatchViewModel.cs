using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ReactiveUI;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Threading.Tasks;
using System.Collections.Generic;
using ExcelToSQLite.Views;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.ViewModels;

public class AnalysisBatchViewModel : ReactiveObject, IRefreshablePage
{
    private readonly AnalysisBatchService _batchService = null!;
    private readonly IndicatorService _indicatorService = null!;
    private Window? _parentWindow;

    private ObservableCollection<AnalysisBatch> _batches = new();
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = Brushes.Green;
    private Window? _editDialogWindow;
    private Window? _scanDialogWindow;
    private bool _isLoading = false;

    public AnalysisBatchViewModel()
    {
        try
        {
            _batchService = new AnalysisBatchService();
            _indicatorService = new IndicatorService();

            CreateBatchCommand = ReactiveCommand.CreateFromTask(CreateBatchAsync);
            EditCommand = ReactiveCommand.CreateFromTask<AnalysisBatch>(EditBatchAsync);
            DeleteCommand = ReactiveCommand.CreateFromTask<AnalysisBatch>(DeleteBatchAsync);
            RefreshCommand = ReactiveCommand.CreateFromTask(LoadDataAsync);
            ScanCommand = ReactiveCommand.CreateFromTask<AnalysisBatch>(ScanBatchAsync);

            // ✅ 使用 Task.Run 避免阻塞构造函数
            _ = Task.Run(async () =>
            {
                try
                {
                    await LoadDataAsync();
                }
                catch (Exception ex)
                {
                    await SetStatusSafeAsync($"加载数据失败: {ex.Message}", Brushes.Red);
                }
            });
        }
        catch (Exception ex)
        {
            _ = SetStatusSafeAsync($"初始化失败: {ex.Message}", Brushes.Red);
        }
    }

    #region 线程安全的UI更新方法

    /// <summary>
    /// 线程安全的状态设置
    /// </summary>
    private async Task SetStatusSafeAsync(string message, IBrush color)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = color;
        });
    }

    /// <summary>
    /// 线程安全的同步状态设置（用于紧急情况）
    /// </summary>
    private void SetStatusSafe(string message, IBrush color)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = color;
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 线程安全的消息显示
    /// </summary>
    private async Task ShowMessageSafeAsync(string msg)
    {
        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            try
            {
                // 优先使用 _parentWindow
                if (_parentWindow != null && _parentWindow.IsVisible)
                {
                    await MessageBox.ShowAsync(_parentWindow, msg, "提示", MessageBoxButtons.OK);
                    return;
                }

                // 尝试获取主窗口
                var mainWindow = (App.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
                if (mainWindow != null && mainWindow.IsVisible)
                {
                    await MessageBox.ShowAsync(mainWindow, msg, "提示", MessageBoxButtons.OK);
                    return;
                }

                // 如果所有窗口都不可用，使用无所有者的 MessageBox
                await MessageBox.ShowAsync(msg, "提示", MessageBoxButtons.OK);
            }
            catch
            {
            }
        });
    }

    /// <summary>
    /// 线程安全的确认对话框
    /// </summary>
    private async Task<bool> ShowConfirmDialogSafeAsync(string message)
    {
        return await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            try
            {
                // 优先使用 _parentWindow
                if (_parentWindow != null && _parentWindow.IsVisible)
                {
                    var result = await MessageBox.ShowAsync(_parentWindow, message, "确认删除", MessageBoxButtons.YesNo);
                    return result == MessageBoxResult.Yes;
                }

                // 尝试获取主窗口
                var mainWindow = (App.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime)?.MainWindow;
                if (mainWindow != null && mainWindow.IsVisible)
                {
                    var result = await MessageBox.ShowAsync(mainWindow, message, "确认删除", MessageBoxButtons.YesNo);
                    return result == MessageBoxResult.Yes;
                }

                // 如果所有窗口都不可用，使用无所有者的 MessageBox
                var result2 = await MessageBox.ShowAsync(message, "确认删除", MessageBoxButtons.YesNo);
                return result2 == MessageBoxResult.Yes;
            }
            catch
            {
                return false;
            }
        });
    }

    #endregion

    #region 属性

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    public ObservableCollection<AnalysisBatch> Batches
    {
        get => _batches;
        set => this.RaiseAndSetIfChanged(ref _batches, value);
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

    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> CreateBatchCommand { get; } = null!;
    public ReactiveCommand<AnalysisBatch, Unit> EditCommand { get; } = null!;
    public ReactiveCommand<AnalysisBatch, Unit> DeleteCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> RefreshCommand { get; } = null!;
    public ReactiveCommand<AnalysisBatch, Unit> ScanCommand { get; } = null!;

    #endregion

    #region 批次操作

    private async Task CreateBatchAsync()
    {
        await OpenEditDialogAsync(null);
    }

    private async Task EditBatchAsync(AnalysisBatch batch)
    {
        if (batch == null) return;
        await OpenEditDialogAsync(batch);
    }

    private async Task OpenEditDialogAsync(AnalysisBatch? batch)
    {
        if (_parentWindow == null)
        {
            if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                _parentWindow = desktop.MainWindow;
            }

            if (_parentWindow == null)
            {
                await ShowMessageSafeAsync("无法获取父窗口");
                return;
            }
        }

        var editViewModel = new AnalysisBatchEditViewModel();
        if (batch != null)
        {
            editViewModel.LoadBatch(batch);
        }

        editViewModel.OnSaveBatch = async () =>
        {
            try
            {
                // 验证
                if (string.IsNullOrWhiteSpace(editViewModel.Name))
                {
                    await ShowMessageSafeAsync("请输入批次名称");
                    return false;
                }

                if (editViewModel.PeriodStart > editViewModel.PeriodEnd)
                {
                    await ShowMessageSafeAsync("期间开始日期不能晚于结束日期");
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(editViewModel.TablePrefix))
                {
                    if (!AnalysisBatchService.IsValidTablePrefix(editViewModel.TablePrefix))
                    {
                        await ShowMessageSafeAsync(
                            "表名前缀格式不正确！\n\n" +
                            "要求：\n" +
                            "• 必须以英文字母开头\n" +
                            "• 只能包含英文字母、数字和下划线\n" +
                            "• 不能包含空格或特殊字符\n\n" +
                            $"当前输入: '{editViewModel.TablePrefix}'"
                        );
                        return false;
                    }
                }

                var newBatch = new AnalysisBatch
                {
                    Id = editViewModel.Id,
                    Name = editViewModel.Name.Trim(),
                    PeriodStart = editViewModel.PeriodStart.DateTime,
                    PeriodEnd = editViewModel.PeriodEnd.DateTime,
                    CreatedAt = editViewModel.IsEditing
                        ? (batch?.CreatedAt ?? DateTime.Now)
                        : DateTime.Now,
                    TablePrefix = editViewModel.TablePrefix.Trim()
                };

                try
                {
                    if (editViewModel.IsEditing)
                    {
                        await _batchService.UpdateAsync(newBatch);
                        await SetStatusSafeAsync($"批次 '{newBatch.Name}' 更新成功！", new SolidColorBrush(Colors.Green));
                    }
                    else
                    {
                        await _batchService.AddAsync(newBatch);
                        await SetStatusSafeAsync($"批次 '{newBatch.Name}' 创建成功！", new SolidColorBrush(Colors.Green));
                    }

                    await LoadDataAsync();

                    // ✅ 保存成功后关闭窗口（在 UI 线程）
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        _editDialogWindow?.Close();
                    });

                    return true;
                }
                catch (Exception ex)
                {
                    await SetStatusSafeAsync($"保存失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                    await ShowMessageSafeAsync($"保存失败: {ex.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                await SetStatusSafeAsync($"操作失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                return false;
            }
        };

        // ✅ 修复：OnCancel 使用 ThreadingHelper 确保在 UI 线程关闭窗口
        editViewModel.OnCancel = () =>
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                _editDialogWindow?.Close();
            }).ConfigureAwait(false);
        };

        var dialog = new Window
        {
            Title = editViewModel.DialogTitle,
            Width = 580,
            Height = 480,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false,
            Icon = IconHelper.GetAppIcon()
        };

        var view = new Views.AnalysisBatchEditDialog
        {
            DataContext = editViewModel
        };

        dialog.Content = view;
        _editDialogWindow = dialog;

        await dialog.ShowDialog(_parentWindow);
        _editDialogWindow = null;
    }

    private async Task DeleteBatchAsync(AnalysisBatch batch)
    {
        if (batch == null)
        {
            await SetStatusSafeAsync("未选择要删除的批次", Brushes.Red);
            return;
        }

        if (AnalysisBatchService.IsDefaultBatch(batch.Id))
        {
            await ShowMessageSafeAsync("不能删除默认分析批次，它是系统必需的。");
            await SetStatusSafeAsync("不能删除默认分析批次", Brushes.Orange);
            return;
        }

        var confirm = await ShowConfirmDialogSafeAsync(
            $"确认删除批次\n\n" +
            $"批次名称: {batch.Name}\n" +
            $"期间: {batch.PeriodStart:yyyy-MM-dd} 至 {batch.PeriodEnd:yyyy-MM-dd}\n\n" +
            $"确认删除？"
        );

        if (!confirm) return;

        try
        {
            await _batchService.DeleteAsync(batch.Id);
            await SetStatusSafeAsync($"批次 '{batch.Name}' 删除成功！", Brushes.Green);
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"删除失败: {ex.Message}", Brushes.Red);
            await ShowMessageSafeAsync($"删除失败: {ex.Message}");
        }
    }

    #endregion

    #region 数据加载

    public async Task RefreshAsync()
    {
        await LoadDataAsync();
    }

    private async Task LoadDataAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
            });

            var list = await _batchService.GetAllAsync(1000);
            var sortedList = list.OrderByDescending(b => b.CreatedAt).ToList();

            for (int i = 0; i < sortedList.Count; i++)
            {
                sortedList[i].Index = i + 1;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Batches = new ObservableCollection<AnalysisBatch>(sortedList);
                IsLoading = false;
            });

            await SetStatusSafeAsync($"加载完成，共 {Batches.Count} 条记录", Brushes.Green);
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"加载失败: {ex.Message}", Brushes.Red);
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = false;
            });
        }
    }

    #endregion

    #region 扫描操作

    private async Task ScanBatchAsync(AnalysisBatch batch)
    {
        if (batch == null)
        {
            await SetStatusSafeAsync("未选择要扫描的批次", Brushes.Red);
            return;
        }

        if (batch.IsScanning)
        {
            await ShowMessageSafeAsync("该批次正在扫描中，请稍候...");
            return;
        }

        try
        {
            // 获取所有指标
            var allIndicators = await _indicatorService.GetAllAsync();

            if (allIndicators.Count == 0)
            {
                await ShowMessageSafeAsync("没有可用的指标，请先在指标管理中创建指标。");
                return;
            }

            // 打开扫描对话框，让用户选择指标
            await OpenScanDialogAsync(batch, allIndicators);
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"打开扫描对话框失败: {ex.Message}", Brushes.Red);
            await ShowMessageSafeAsync($"打开扫描对话框失败: {ex.Message}");
        }
    }

    private async Task OpenScanDialogAsync(AnalysisBatch batch, List<Indicator> indicators)
    {
        if (_parentWindow == null)
        {
            if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                _parentWindow = desktop.MainWindow;
            }

            if (_parentWindow == null)
            {
                await ShowMessageSafeAsync("无法获取父窗口");
                return;
            }
        }

        var scanViewModel = new AnalysisScanViewModel(batch, indicators);
        scanViewModel.OnStartScan = async (selectedIndicators) =>
        {
            if (selectedIndicators == null || selectedIndicators.Count == 0)
            {
                await ShowMessageSafeAsync("请至少选择一个指标");
                return;
            }

            batch.IsScanning = true;
            _scanDialogWindow?.Close();

            await PerformScanAsync(batch, selectedIndicators);
        };

        // ✅ 修复：OnCancel 使用 ThreadingHelper
        scanViewModel.OnCancel = () =>
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                _scanDialogWindow?.Close();
            }).ConfigureAwait(false);
        };

        var dialog = new Window
        {
            Title = $"扫描批次: {batch.Name}",
            Width = 680,
            Height = 600,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = true,
            MinWidth = 600,
            MinHeight = 500,
            Icon = IconHelper.GetAppIcon()
        };

        var view = new Views.AnalysisScanDialog
        {
            DataContext = scanViewModel
        };

        dialog.Content = view;
        _scanDialogWindow = dialog;

        await dialog.ShowDialog(_parentWindow);
        _scanDialogWindow = null;
    }

    /// <summary>
    /// 扫描数据
    /// </summary>
    private async Task PerformScanAsync(AnalysisBatch batch, List<Indicator> indicators)
    {
        Window? progressWindow = null;

        try
        {
            // 确保有父窗口
            if (_parentWindow == null)
            {
                if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                {
                    _parentWindow = desktop.MainWindow;
                }

                if (_parentWindow == null)
                {
                    await ShowMessageSafeAsync("无法获取父窗口，请重试");
                    return;
                }
            }

            // 创建进度窗口
            progressWindow = await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                var window = new Window
                {
                    Title = "扫描进度",
                    Width = 500,
                    Height = 300,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = false
                };

                var progressPanel = new StackPanel
                {
                    Margin = new Avalonia.Thickness(20),
                    Spacing = 16
                };

                var titleText = new TextBlock
                {
                    Text = $"正在扫描批次: {batch.Name}",
                    FontSize = 16,
                    FontWeight = Avalonia.Media.FontWeight.SemiBold
                };

                var progressBar = new ProgressBar
                {
                    Minimum = 0,
                    Maximum = 100,
                    Value = 0,
                    Height = 24
                };

                var statusText = new TextBlock
                {
                    Text = "准备开始扫描...",
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    FontSize = 13
                };

                var resultText = new TextBlock
                {
                    Text = "",
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    Foreground = Brushes.Green,
                    FontSize = 13
                };

                progressPanel.Children.Add(titleText);
                progressPanel.Children.Add(progressBar);
                progressPanel.Children.Add(statusText);
                progressPanel.Children.Add(resultText);
                window.Content = progressPanel;

                return window;
            });

            // 显示进度窗口
            _ = progressWindow.ShowDialog(_parentWindow);

            // ✅ 使用线程安全的进度更新
            var progress = new Progress<(int percent, string status)>(update =>
            {
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (progressWindow?.Content is StackPanel panel)
                    {
                        if (panel.Children[1] is ProgressBar bar)
                            bar.Value = update.percent;
                        if (panel.Children[2] is TextBlock text)
                            text.Text = update.status;
                    }
                }).ConfigureAwait(false);
            });

            // 执行扫描
            var result = await _batchService.ScanAsync(batch, indicators, (percent, status) =>
            {
                ((IProgress<(int, string)>)progress).Report((percent, status));
            });

            // 更新结果信息
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (progressWindow?.Content is StackPanel panel)
                {
                    if (panel.Children[3] is TextBlock resultText)
                    {
                        var successMsg = result.IsSuccess ? "✅ 扫描完成！" : "⚠️ 扫描完成，但有失败项";
                        resultText.Text = $"{successMsg}\n成功: {result.SuccessCount} 项，失败: {result.FailCount} 项，耗时: {result.Duration.TotalSeconds:F1}秒";
                        resultText.Foreground = result.IsSuccess ? Brushes.Green : Brushes.Orange;
                    }
                }
            });

            // 延迟关闭进度窗口，让用户看到完成信息
            await Task.Delay(1500);

            // 关闭进度窗口
            if (progressWindow != null)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    progressWindow.Close();
                });
                progressWindow = null;
            }

            // 显示扫描结果详情
            await ShowScanResultAsync(result);
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                batch.IsScanning = false;
            });
        }
        catch (Exception ex)
        {
            // 关闭进度窗口
            if (progressWindow != null)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    progressWindow.Close();
                });
                progressWindow = null;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                batch.IsScanning = false;
            });

            await SetStatusSafeAsync($"扫描失败: {ex.Message}", Brushes.Red);
            await ShowMessageSafeAsync($"扫描失败: {ex.Message}");
        }
    }

    private async Task ShowScanResultAsync(ScanResult result)
    {
        var message = $"扫描完成!\n\n" +
                      $"批次: {result.BatchName}\n" +
                      $"总指标数: {result.TotalCount}\n" +
                      $"成功: {result.SuccessCount}\n" +
                      $"失败: {result.FailCount}\n" +
                      $"耗时: {result.Duration.TotalSeconds:F2}秒\n";

        if (result.SuccessCount > 0)
        {
            if (result.SaveSuccess)
            {
                message += $"\n✅ 已保存 {result.SavedCount} 条扫描结果到数据库（已覆盖旧记录）";
            }
            else
            {
                message += $"\n⚠️ 扫描结果保存失败: {result.ErrorMessage ?? "未知错误"}";
            }
        }

        if (!string.IsNullOrEmpty(result.ErrorMessage))
        {
            message += $"\n错误: {result.ErrorMessage}";
        }

        if (result.FailCount > 0)
        {
            message += "\n\n失败详情:\n";
            foreach (var item in result.Results.Where(r => !r.IsSuccess))
            {
                message += $"• {item.IndicatorName}: {item.ErrorMessage}\n";
            }
        }

        await SetStatusSafeAsync(
            $"扫描完成: 成功 {result.SuccessCount} 项，失败 {result.FailCount} 项" +
            (result.SaveSuccess ? $" (已保存 {result.SavedCount} 条)" : " (保存失败)"),
            result.IsSuccess ? Brushes.Green : Brushes.Orange);

        await ShowMessageSafeAsync(message);
    }

    #endregion
}