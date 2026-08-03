using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reactive;
using System.Threading.Tasks;
using System.Linq;
using Avalonia.Controls;
using ExcelToSQLite.Models;
using ExcelToSQLite.Helpers;
using ExcelToSQLite.Services;
using ReactiveUI;
using Avalonia.Media;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels;

public class IndicatorManagementViewModel : ReactiveObject, IRefreshablePage, IDisposable
{
    private readonly IndicatorService _indicatorService;
    private readonly IndicatorImportExportService _importExportService;
    private Window? _parentWindow;
    private bool _isDisposed;
    private bool _isLoading;

    // 属性
    private ObservableCollection<Indicator> _indicators = new();
    private ObservableCollection<Indicator> _selectedIndicators = new();
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Color.Parse("#78909C"));
    private bool _showStatus = false;
    private bool _isImporting;
    private bool _isExporting;
    private bool _hasSelectedIndicators;
    private bool _isAllSelected;

    // 命令
    public ReactiveCommand<Unit, Unit> CreateIndicatorCommand { get; }
    public ReactiveCommand<Unit, Unit> RefreshCommand { get; }
    public ReactiveCommand<Indicator, Unit> EditCommand { get; }
    public ReactiveCommand<Indicator, Unit> DeleteCommand { get; }
    public ReactiveCommand<Unit, Unit> DeleteAllCommand { get; }
    // 导入导出命令
    public ReactiveCommand<Unit, Unit> ExportSelectedCommand { get; }
    public ReactiveCommand<Unit, Unit> ExportAllCommand { get; }
    public ReactiveCommand<Unit, Unit> ImportCommand { get; }
    public ReactiveCommand<Unit, Unit> SelectAllCommand { get; }
    public ReactiveCommand<Unit, Unit> ClearSelectionCommand { get; }

    #region 属性

    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    public ObservableCollection<Indicator> Indicators
    {
        get => _indicators;
        set => this.RaiseAndSetIfChanged(ref _indicators, value);
    }

    public ObservableCollection<Indicator> SelectedIndicators
    {
        get => _selectedIndicators;
        set => this.RaiseAndSetIfChanged(ref _selectedIndicators, value);
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

    public bool IsImporting
    {
        get => _isImporting;
        set => this.RaiseAndSetIfChanged(ref _isImporting, value);
    }

    public bool IsExporting
    {
        get => _isExporting;
        set => this.RaiseAndSetIfChanged(ref _isExporting, value);
    }

    public bool IsBusy => IsLoading || IsImporting || IsExporting;

    public bool HasSelectedIndicators
    {
        get => _hasSelectedIndicators;
        set => this.RaiseAndSetIfChanged(ref _hasSelectedIndicators, value);
    }

    public bool IsAllSelected
    {
        get => _isAllSelected;
        set
        {
            this.RaiseAndSetIfChanged(ref _isAllSelected, value);
            if (value)
            {
                SelectAll();
            }
            else
            {
                ClearSelection();
            }
        }
    }

    #endregion

    public IndicatorManagementViewModel()
    {
        _indicatorService = new IndicatorService();
        _importExportService = new IndicatorImportExportService(_indicatorService);
        
        // 初始化命令
        CreateIndicatorCommand = ReactiveCommand.Create(CreateIndicator);
        RefreshCommand = ReactiveCommand.CreateFromTask(RefreshIndicatorsAsync);
        DeleteCommand = ReactiveCommand.CreateFromTask<Indicator>(DeleteIndicatorAsync);
        EditCommand = ReactiveCommand.Create<Indicator>(EditIndicator);
        DeleteAllCommand = ReactiveCommand.CreateFromTask(DeleteAllIndicatorsAsync);

        // 导入导出命令
        ExportSelectedCommand = ReactiveCommand.CreateFromTask(ExportSelectedAsync);
        ExportAllCommand = ReactiveCommand.CreateFromTask(ExportAllAsync);
        ImportCommand = ReactiveCommand.CreateFromTask(ImportAsync);
        SelectAllCommand = ReactiveCommand.Create(SelectAll);
        ClearSelectionCommand = ReactiveCommand.Create(ClearSelection);
        
        // 初始化集合
        Indicators = new ObservableCollection<Indicator>();
        SelectedIndicators = new ObservableCollection<Indicator>();
        
        // 订阅 Indicator 的选择变化事件
        Indicator.SelectionChanged += OnIndicatorSelectionChanged;
        
        // 监听选中项变化
        this.WhenAnyValue(x => x.SelectedIndicators.Count)
            .Subscribe(count => 
            {
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    HasSelectedIndicators = count > 0;
                    UpdateAllSelectedState();
                }).ConfigureAwait(false);
            });
        
        // 监听 Indicators 变化
        this.WhenAnyValue(x => x.Indicators)
            .Subscribe(_ => 
            {
                ThreadingHelper.RunOnUIThreadAsync(UpdateAllSelectedState)
                    .ConfigureAwait(false);
            });
        
        _ = InitializeAndLoadSafe();
    }

    #region 公共方法

    public void SetParentWindow(Window? window)
    {
        _parentWindow = window;
    }

    public async Task RefreshAsync()
    {
        await RefreshIndicatorsAsync();
    }

    #endregion

    #region 初始化

    private async Task InitializeAndLoadSafe()
    {
        try
        {
            await RefreshIndicatorsAsync();
        }
        catch (Exception ex)
        {
            if (ex.InnerException != null)
            {
            }
            await SetStatusSafeAsync($"❌ 初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #endregion

    #region 状态管理

    private async Task SetStatusSafeAsync(string message, IBrush color)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = color;
            ShowStatus = !string.IsNullOrEmpty(message);
        });
    }

    private void SetStatusSafe(string message, IBrush color)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = color;
            ShowStatus = !string.IsNullOrEmpty(message);
        }).ConfigureAwait(false);
    }

    #endregion

    #region 选择方法

    private void OnIndicatorSelectionChanged(Indicator indicator, bool isSelected)
    {
        if (indicator == null) return;
    
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (isSelected)
            {
                if (!SelectedIndicators.Contains(indicator))
                {
                    SelectedIndicators.Add(indicator);
                }
            }
            else
            {
                SelectedIndicators.Remove(indicator);
            }
            UpdateAllSelectedState();
        }).ConfigureAwait(false);
    }

    private void UpdateAllSelectedState()
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(UpdateAllSelectedState);
            return;
        }

        if (Indicators.Count == 0)
        {
            _isAllSelected = false;
            this.RaisePropertyChanged(nameof(IsAllSelected));
            return;
        }
        
        var allSelected = Indicators.All(i => i.IsSelected);
        if (_isAllSelected != allSelected)
        {
            _isAllSelected = allSelected;
            this.RaisePropertyChanged(nameof(IsAllSelected));
        }
        
        var hasSelected = SelectedIndicators.Count > 0;
        if (_hasSelectedIndicators != hasSelected)
        {
            _hasSelectedIndicators = hasSelected;
            this.RaisePropertyChanged(nameof(HasSelectedIndicators));
        }
    }

    private void SelectAll()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            foreach (var indicator in Indicators)
            {
                indicator.IsSelected = true;
            }
            UpdateAllSelectedState();
        }).ConfigureAwait(false);
    }

    private void ClearSelection()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            foreach (var indicator in Indicators)
            {
                indicator.IsSelected = false;
            }
            UpdateAllSelectedState();
        }).ConfigureAwait(false);
    }

    #endregion

    #region 基础CRUD操作

    private void CreateIndicator()
    {
        var editViewModel = new IndicatorEditViewModel();
        editViewModel.OnSaveIndicator = async () =>
        {
            try
            {
                var indicator = new Indicator
                {
                    Id = Guid.NewGuid().ToString(),
                    Name = editViewModel.IndicatorName,
                    SqlStatement = editViewModel.IndicatorSqlStatement ?? string.Empty,
                    SqlDetailData = editViewModel.IndicatorSqlDetailData ?? string.Empty,
                    Description = editViewModel.IndicatorDescription ?? string.Empty,
                    Category = editViewModel.IndicatorCategory,
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now,
                    CreatedBy = "admin",
                    IsActive = true
                };

                await _indicatorService.AddAsync(indicator);
                await RefreshIndicatorsAsync();
            
                await SetStatusSafeAsync($"✅ 指标 '{indicator.Name}' 创建成功", new SolidColorBrush(Colors.Green));
                return true;
            }
            catch (Exception ex)
            {
                await SetStatusSafeAsync($"❌ 创建指标失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                return false;
            }
        };

        ShowEditDialog(editViewModel);
    }

    private async Task RefreshIndicatorsAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
                SetStatusSafe("正在加载指标...", new SolidColorBrush(Colors.Orange));
            });

            var indicators = await _indicatorService.GetAllAsync();
            
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Indicators.Clear();
                SelectedIndicators.Clear();

                int index = 1;
                foreach (var indicator in indicators)
                {
                    indicator.Index = index++;
                    indicator.IsSelected = false;
                    Indicators.Add(indicator);
                }

                IsLoading = false;
                SetStatusSafe($"✅ 已加载 {Indicators.Count} 个指标", new SolidColorBrush(Colors.Green));
                UpdateAllSelectedState();
            });
        }
        catch (Exception ex)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = false;
            });
            await SetStatusSafeAsync($"❌ 刷新列表失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private void EditIndicator(Indicator indicator)
    {
        if (indicator == null) return;

        var editViewModel = new IndicatorEditViewModel();
        editViewModel.LoadIndicator(indicator);
        editViewModel.OnSaveIndicator = async () =>
        {
            try
            {
                indicator.Name = editViewModel.IndicatorName;
                indicator.SqlStatement = editViewModel.IndicatorSqlStatement ?? string.Empty;
                indicator.SqlDetailData = editViewModel.IndicatorSqlDetailData ?? string.Empty;
                indicator.Description = editViewModel.IndicatorDescription ?? string.Empty;
                indicator.Category = editViewModel.IndicatorCategory;
                indicator.UpdatedAt = DateTime.Now;

                await _indicatorService.UpdateAsync(indicator);
                await RefreshIndicatorsAsync();
            
                await SetStatusSafeAsync($"✅ 指标 '{indicator.Name}' 更新成功", new SolidColorBrush(Colors.Green));
                return true;
            }
            catch (Exception ex)
            {
                await SetStatusSafeAsync($"❌ 更新指标失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                return false;
            }
        };

        ShowEditDialog(editViewModel);
    }

    private async void ShowEditDialog(IndicatorEditViewModel viewModel)
    {
        if (_parentWindow == null)
        {
            await SetStatusSafeAsync("无法显示对话框：未设置父窗口", new SolidColorBrush(Colors.Red));
            return;
        }

        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var dialog = new Window
                {
                    Title = viewModel.DialogTitle,
                    Width = 740,
                    Height = 660,
                    MinWidth = 650,
                    MinHeight = 580,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = true,
                    Icon = IconHelper.GetAppIcon()
                };

                var editControl = new Views.IndicatorEditDialog
                {
                    DataContext = viewModel
                };
                dialog.Content = editControl;

                viewModel.SetDialogWindow(dialog);

                viewModel.OnClose = (success) =>
                {
                    dialog.Close();
                    if (success)
                    {
                        SetStatusSafe("操作成功", new SolidColorBrush(Colors.Green));
                    }
                    else
                    {
                        SetStatusSafe("已取消操作", new SolidColorBrush(Color.Parse("#78909C")));
                    }
                };

                await dialog.ShowDialog(_parentWindow);
            });
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"显示对话框失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private async Task DeleteIndicatorAsync(Indicator indicator)
    {
        if (indicator == null) return;

        if (_parentWindow != null)
        {
            var confirm = await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var dialog = new MessageBox
                {
                    Title = "确认删除",
                    MessageBoxContent = $"确定要删除指标 '{indicator.Name}' 吗？",
                    Buttons = MessageBoxButtons.YesNo
                };

                var result = await dialog.ShowDialogAsync(_parentWindow);
                return result == MessageBoxResult.Yes;
            });

            if (!confirm) return;
        }

        try
        {
            await _indicatorService.DeleteAsync(indicator.Id);
            await RefreshIndicatorsAsync();

            await SetStatusSafeAsync($"✅ 指标 '{indicator.Name}' 删除成功", new SolidColorBrush(Colors.Green));
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 删除指标失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    /// <summary>
    /// 删除所有指标
    /// </summary>
    private async Task DeleteAllIndicatorsAsync()
    {
        if (Indicators.Count == 0)
        {
            await SetStatusSafeAsync("没有指标可以删除", new SolidColorBrush(Colors.Red));
            return;
        }

        if (_parentWindow != null)
        {
            var confirm = await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var dialog = new MessageBox
                {
                    Title = "确认全部删除",
                    MessageBoxContent = $"⚠️⚠️ 警告：即将删除全部 {Indicators.Count} 个指标！\n\n此操作不可恢复！\n\n确认删除全部指标？",
                    Buttons = MessageBoxButtons.YesNo
                };

                var result = await dialog.ShowDialogAsync(_parentWindow);
                return result == MessageBoxResult.Yes;
            });

            if (!confirm) return;
        }

        try
        {
            var deletedCount = 0;
            var errors = new List<string>();

            foreach (var indicator in Indicators.ToList())
            {
                try
                {
                    await _indicatorService.DeleteAsync(indicator.Id);
                    deletedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"{indicator.Name}: {ex.Message}");
                }
            }

            await RefreshIndicatorsAsync();

            var message = $"✅ 已删除 {deletedCount} 个指标";
            if (errors.Any())
            {
                message += $"\n⚠️ 以下指标删除失败:\n{string.Join("\n", errors.Take(5))}";
            }

            await SetStatusSafeAsync(message, errors.Any() ? new SolidColorBrush(Colors.Orange) : new SolidColorBrush(Colors.Green));
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 批量删除失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #endregion

    #region 导入导出方法

    private async Task ExportSelectedAsync()
    {
        if (SelectedIndicators.Count == 0)
        {
            await SetStatusSafeAsync("请至少选择一个指标进行导出", new SolidColorBrush(Colors.Red));
            return;
        }

        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = true;
                SetStatusSafe("正在生成导出文件...", new SolidColorBrush(Colors.Orange));
            });

            var indicators = SelectedIndicators.ToList();
            var json = _importExportService.ExportIndicators(indicators, "admin");

            var filePath = await FileDialogHelper.SaveFileAsync(
                _parentWindow, 
                json, 
                "保存指标文件", 
                "van", 
                "指标文件");
                        
            if (!string.IsNullOrEmpty(filePath))
            {
                await SetStatusSafeAsync($"✅ 成功导出 {indicators.Count} 个指标到: {System.IO.Path.GetFileName(filePath)}", new SolidColorBrush(Colors.Green));
            }
            else
            {
                await SetStatusSafeAsync("导出已取消", new SolidColorBrush(Color.Parse("#78909C")));
            }
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 导出失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = false;
            });
        }
    }

    private async Task ExportAllAsync()
    {
        if (Indicators.Count == 0)
        {
            await SetStatusSafeAsync("没有可导出的指标", new SolidColorBrush(Colors.Red));
            return;
        }

        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = true;
                SetStatusSafe("正在生成导出文件...", new SolidColorBrush(Colors.Orange));
            });

            var json = _importExportService.ExportIndicators(Indicators.ToList(), "admin");

            var filePath = await FileDialogHelper.SaveFileAsync(
                _parentWindow, 
                json, 
                "保存指标文件", 
                "van", 
                "指标文件");
                        
            if (!string.IsNullOrEmpty(filePath))
            {
                await SetStatusSafeAsync($"✅ 成功导出全部 {Indicators.Count} 个指标", new SolidColorBrush(Colors.Green));
            }
            else
            {
                await SetStatusSafeAsync("导出已取消", new SolidColorBrush(Color.Parse("#78909C")));
            }
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 导出失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = false;
            });
        }
    }

    private async Task ImportAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsImporting = true;
                SetStatusSafe("请选择要导入的 .van 文件...", new SolidColorBrush(Colors.Orange));
            });

            var fileContent = await FileDialogHelper.OpenVanFileAndReadContentAsync(
                _parentWindow,
                "选择指标文件");
                
            if (string.IsNullOrEmpty(fileContent))
            {
                await SetStatusSafeAsync("导入已取消", new SolidColorBrush(Color.Parse("#78909C")));
                return;
            }

            if (!_importExportService.ValidateVanFile(fileContent))
            {
                await SetStatusSafeAsync("❌ 文件格式无效，请选择有效的 .van 文件", new SolidColorBrush(Colors.Red));
                return;
            }

            var summary = _importExportService.GetExportSummary(fileContent);
            if (summary != null)
            {
                await SetStatusSafeAsync($"📄 检测到 {summary.TotalCount} 个指标，准备导入...", new SolidColorBrush(Color.Parse("#1976D2")));
            }

            if (_parentWindow != null)
            {
                var confirm = await ThreadingHelper.RunOnUIThreadAsync(async () =>
                {
                    var confirmDialog = new MessageBox
                    {
                        Title = "确认导入",
                        MessageBoxContent = $"即将导入 {summary?.TotalCount ?? 0} 个指标（追加模式），\n" +
                                           $"统计SQL 为空: {summary?.EmptySqlStatement ?? 0} 个\n" +
                                           $"详细SQL 为空: {summary?.EmptySqlDetailData ?? 0} 个\n\n" +
                                           $"确定要继续吗？",
                        Buttons = MessageBoxButtons.YesNo
                    };
                    
                    var result = await confirmDialog.ShowDialogAsync(_parentWindow);
                    return result == MessageBoxResult.Yes;
                });
                
                if (!confirm)
                {
                    await SetStatusSafeAsync("导入已取消", new SolidColorBrush(Color.Parse("#78909C")));
                    return;
                }
            }
            
            await SetStatusSafeAsync("正在导入指标...", new SolidColorBrush(Colors.Orange));
            
            var importResult = await _importExportService.ImportIndicatorsAsync(fileContent);
            await RefreshIndicatorsAsync();

            var summaryMessage = importResult.GetSummary();
            await SetStatusSafeAsync($"✅ 导入完成: {summaryMessage}", new SolidColorBrush(Colors.Green));
            
            if (importResult.Errors.Any())
            {
                await ShowImportResultAsync(importResult);
            }
            else
            {
                await ShowImportResultAsync(importResult);
            }
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 导入失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsImporting = false;
            });
        }
    }

    private async Task ShowImportResultAsync(ImportResult result)
    {
        if (_parentWindow == null) return;

        var details = new System.Text.StringBuilder();
        details.AppendLine("📊 导入结果");
        details.AppendLine("━━━━━━━━━━━━━━━━━━━━━━━━━");
        details.AppendLine($"📝 总计: {result.TotalCount} 个指标");
        details.AppendLine($"✅ 新增: {result.ImportedCount} 个");
        
        if (result.EmptySqlStatement > 0)
            details.AppendLine($"⚠️ 统计SQL 为空: {result.EmptySqlStatement} 个");
        if (result.EmptySqlDetailData > 0)
            details.AppendLine($"⚠️ 详细SQL 为空: {result.EmptySqlDetailData} 个");
        
        if (result.Errors.Any())
        {
            details.AppendLine($"\n❌ 错误: {result.Errors.Count} 个");
            foreach (var error in result.Errors.Take(5))
            {
                details.AppendLine($"  • {error}");
            }
            if (result.Errors.Count > 5)
                details.AppendLine($"  ... 还有 {result.Errors.Count - 5} 个错误");
        }

        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            var dialog = new MessageBox
            {
                Title = "导入完成",
                MessageBoxContent = details.ToString(),
                Buttons = MessageBoxButtons.OK
            };
            
            await dialog.ShowDialogAsync(_parentWindow);
        });
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                // 取消事件订阅
                Indicator.SelectionChanged -= OnIndicatorSelectionChanged;
                
                // 释放命令
                CreateIndicatorCommand?.Dispose();
                RefreshCommand?.Dispose();
                EditCommand?.Dispose();
                DeleteCommand?.Dispose();
                DeleteAllCommand?.Dispose();
                ExportSelectedCommand?.Dispose();
                ExportAllCommand?.Dispose();
                ImportCommand?.Dispose();
                SelectAllCommand?.Dispose();
                ClearSelectionCommand?.Dispose();
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