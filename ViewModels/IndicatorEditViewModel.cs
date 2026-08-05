using ReactiveUI;
using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Reactive;
using System.Threading.Tasks;
using System.Linq;
using Avalonia.Controls;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ExcelToSQLite.Views;
using ExcelToSQLite.Helpers;
using Avalonia.Media;

namespace ExcelToSQLite.ViewModels;

public class IndicatorEditViewModel : ReactiveObject, IDisposable
{
    private readonly AppConfigService _configService;
    private readonly IndicatorService _indicatorService;
    private bool _isDisposed;
    private bool _isLoading;

    private string _indicatorId = string.Empty;
    private string _indicatorName = string.Empty;
    private string _indicatorSqlStatement = string.Empty;
    private string _indicatorSqlDetailData = string.Empty;
    private string _indicatorDescription = string.Empty;
    private string _indicatorCategory = string.Empty;
    private int _selectedTabIndex = 0;
    private string _dialogTitle = "创建新指标";
    private string _dialogSubtitle = "填写指标信息，创建新的分析指标";
    private bool _isEditing = false;
    private string _defaultCategory = string.Empty;
    private bool _isCategoriesLoaded = false;
    private string _selectedCategoryOnLoad = string.Empty;

    private Window? _dialogWindow;

    public ObservableCollection<string> CategoryOptions { get; } = new();

    #region 属性

    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    public string IndicatorId
    {
        get => _indicatorId;
        set => this.RaiseAndSetIfChanged(ref _indicatorId, value);
    }

    public string IndicatorName
    {
        get => _indicatorName;
        set => this.RaiseAndSetIfChanged(ref _indicatorName, value);
    }

    public string IndicatorSqlStatement
    {
        get => _indicatorSqlStatement;
        set => this.RaiseAndSetIfChanged(ref _indicatorSqlStatement, value);
    }

    public string IndicatorSqlDetailData
    {
        get => _indicatorSqlDetailData;
        set => this.RaiseAndSetIfChanged(ref _indicatorSqlDetailData, value);
    }

    public string IndicatorDescription
    {
        get => _indicatorDescription;
        set => this.RaiseAndSetIfChanged(ref _indicatorDescription, value);
    }

    public string IndicatorCategory
    {
        get => _indicatorCategory;
        set => this.RaiseAndSetIfChanged(ref _indicatorCategory, value);
    }
    public bool CanPreview => SelectedTabIndex != 0; // 0 = 指标描述标签页
    public int SelectedTabIndex
    {
        get => _selectedTabIndex;
        set 
        {
            this.RaiseAndSetIfChanged(ref _selectedTabIndex, value);
            this.RaisePropertyChanged(nameof(CanPreview));
        }
    }

    public string DialogTitle
    {
        get => _dialogTitle;
        set => this.RaiseAndSetIfChanged(ref _dialogTitle, value);
    }

    public string DialogSubtitle
    {
        get => _dialogSubtitle;
        set => this.RaiseAndSetIfChanged(ref _dialogSubtitle, value);
    }

    public bool IsEditing
    {
        get => _isEditing;
        set
        {
            this.RaiseAndSetIfChanged(ref _isEditing, value);
            UpdateDialogTitle();
        }
    }   

    #endregion

    #region 公共属性

    public Indicator? EditingIndicator { get; private set; }
    public Func<Task<bool>>? OnSaveIndicator { get; set; }
    public Action<bool>? OnClose { get; set; }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> SaveCommand { get; }
    public ReactiveCommand<Unit, Unit> CancelCommand { get; }
    public ReactiveCommand<Unit, Unit> RefreshCategoriesCommand { get; }
    public ReactiveCommand<Unit, Unit> PreviewCommand { get; }

    #endregion

    public IndicatorEditViewModel()
    {
        _configService = new AppConfigService();
        _indicatorService = new IndicatorService();

        LoadDefaultCategories();
        _defaultCategory = CategoryOptions.FirstOrDefault() ?? "其他";
        IndicatorCategory = _defaultCategory;

        SaveCommand = ReactiveCommand.CreateFromTask(SaveAsync);
        CancelCommand = ReactiveCommand.Create(Cancel);
        RefreshCategoriesCommand = ReactiveCommand.CreateFromTask(RefreshCategoriesAsync);
        PreviewCommand = ReactiveCommand.CreateFromTask(PreviewDataAsync);

        _ = LoadCategoriesFromConfigAsync();
    }

    #region 公共方法

    public void SetDialogWindow(Window window)
    {
        _dialogWindow = window;
    }

    public void LoadIndicator(Indicator indicator)
    {
        if (indicator == null)
            return;

        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            EditingIndicator = indicator;
            IndicatorId = indicator.Id;
            IndicatorName = indicator.Name;
            IndicatorSqlStatement = indicator.SqlStatement;
            IndicatorSqlDetailData = indicator.SqlDetailData;
            IndicatorDescription = indicator.Description;
            IsEditing = true;

            string targetCategory = string.IsNullOrEmpty(indicator.Category) ? _defaultCategory : indicator.Category;

            if (_isCategoriesLoaded)
            {
                if (CategoryOptions.Contains(targetCategory))
                {
                    IndicatorCategory = targetCategory;
                }
                else
                {
                    IndicatorCategory = _defaultCategory;
                }
            }
            else
            {
                _selectedCategoryOnLoad = targetCategory;
                IndicatorCategory = _defaultCategory;
            }
        }).ConfigureAwait(false);
    }

    public void Reset()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IndicatorId = string.Empty;
            IndicatorName = string.Empty;
            IndicatorSqlStatement = string.Empty;
            IndicatorSqlDetailData = string.Empty;
            IndicatorDescription = string.Empty;
            IndicatorCategory = _defaultCategory;
            SelectedTabIndex = 0;
            IsEditing = false;
            EditingIndicator = null;
            _selectedCategoryOnLoad = string.Empty;
            IsLoading = false;
        }).ConfigureAwait(false);
    }

    #endregion

    #region 分类加载

    private void LoadDefaultCategories()
    {
        CategoryOptions.Clear();
        var defaultCategories = GetDefaultCategories();
        foreach (var category in defaultCategories)
        {
            CategoryOptions.Add(category);
        }
    }

    private List<string> GetDefaultCategories()
    {
        return _configService.GetDefaultCategories();
    }

    private async Task LoadCategoriesFromConfigAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
            });

            var categories = await _configService.GetCategoriesAsync();

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (categories != null && categories.Count > 0)
                {
                    CategoryOptions.Clear();
                    foreach (var category in categories)
                    {
                        CategoryOptions.Add(category);
                    }
                    _defaultCategory = CategoryOptions.FirstOrDefault() ?? "其他";
                }

                _isCategoriesLoaded = true;

                if (!string.IsNullOrEmpty(_selectedCategoryOnLoad))
                {
                    if (CategoryOptions.Contains(_selectedCategoryOnLoad))
                    {
                        IndicatorCategory = _selectedCategoryOnLoad;
                    }
                    else
                    {
                        IndicatorCategory = _defaultCategory;
                    }
                    _selectedCategoryOnLoad = string.Empty;
                }
                else
                {
                    if (!CategoryOptions.Contains(IndicatorCategory))
                    {
                        IndicatorCategory = _defaultCategory;
                    }
                }

                IsLoading = false;
            });
        }
        catch
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = false;
            });
        }
    }

    private async Task RefreshCategoriesAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
            });

            var categories = await _configService.GetCategoriesAsync();

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                CategoryOptions.Clear();

                if (categories == null || categories.Count == 0)
                {
                    var defaultCategories = GetDefaultCategories();
                    foreach (var category in defaultCategories)
                    {
                        CategoryOptions.Add(category);
                    }
                }
                else
                {
                    foreach (var category in categories)
                    {
                        CategoryOptions.Add(category);
                    }
                }

                _defaultCategory = CategoryOptions.FirstOrDefault() ?? "其他";

                if (!CategoryOptions.Contains(IndicatorCategory))
                {
                    IndicatorCategory = _defaultCategory;
                }

                _isCategoriesLoaded = true;
                IsLoading = false;

            });
        }
        catch
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = false;
            });
        }
    }

    #endregion

    #region 对话框标题

    private void UpdateDialogTitle()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (_isEditing)
            {
                DialogTitle = "编辑指标";
                DialogSubtitle = "修改指标信息并保存";
            }
            else
            {
                DialogTitle = "创建新指标";
                DialogSubtitle = "填写指标信息，创建新的分析指标";
            }
        }).ConfigureAwait(false);
    }

    #endregion

    #region 预览

    private async Task PreviewDataAsync()
    {
        try
        {
            var sql = SelectedTabIndex == 0 ? IndicatorSqlStatement : IndicatorSqlDetailData;
            var sqlLabel = SelectedTabIndex == 0 ? "统计SQL" : "详细SQL";

            if (string.IsNullOrWhiteSpace(sql))
            {
                return;
            }


            var previewViewModel = new DetailDataViewModel(sql, $"{IndicatorName}_{sqlLabel}", IndicatorCategory);
            var detailView = new DetailDataView
            {
                DataContext = previewViewModel
            };

            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var previewWindow = new Window
                {
                    Title = $"数据预览 - {IndicatorName} ({sqlLabel})",
                    Width = 1200,
                    Height = 700,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Content = detailView,
                    CanResize = true,
                    MinWidth = 800,
                    MinHeight = 500,
                    Icon = IconHelper.GetAppIcon()
                };

                if (_dialogWindow != null)
                {
                    detailView.SetParentWindow(_dialogWindow);
                    previewViewModel.SetParentWindow(_dialogWindow);
                    await previewWindow.ShowDialog(_dialogWindow);
                }
                else
                {
                    previewWindow.Show();
                }
            });
        }
        catch
        {
        }
    }

    #endregion

    #region 保存和取消
    private async Task SaveAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
            });

            // ✅ 只验证名称，不验证 SQL
            if (string.IsNullOrWhiteSpace(IndicatorName))
            {
                await MessageBox.ShowAsync("指标名称不能为空！",icon:MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(IndicatorCategory))
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    IndicatorCategory = _defaultCategory;
                });
            }
            // ✅ 只验证名称，不验证 SQL
            if (string.IsNullOrWhiteSpace(IndicatorSqlStatement))
            {
                await MessageBox.ShowAsync("统计SQL 内容不能为空...",icon:MessageBoxIcon.Warning);
                return;
            }
            
            bool success = false;
            if (OnSaveIndicator != null)
            {
                success = await OnSaveIndicator();
            }
        
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                OnClose?.Invoke(success);
                IsLoading = false;
            });
        }
        catch
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                OnClose?.Invoke(false);
                IsLoading = false;
            });
        }
    }

    private void Cancel()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            OnClose?.Invoke(false);
        }).ConfigureAwait(false);
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                SaveCommand?.Dispose();
                CancelCommand?.Dispose();
                RefreshCategoriesCommand?.Dispose();
                PreviewCommand?.Dispose();
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