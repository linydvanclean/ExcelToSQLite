using ExcelToSQLite.Models;
using ExcelToSQLite.Helpers;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Linq;
using System.Reactive;
using System.Reactive.Linq;
using System.Threading.Tasks;

namespace ExcelToSQLite.ViewModels;

public class AnalysisScanViewModel : ReactiveObject, IDisposable
{
    private readonly AnalysisBatch _batch;
    private ObservableCollection<Indicator> _indicators = new();
    private string _statusMessage = string.Empty;
    private bool _canStartScan = true;
    private int _selectedCount;
    private int _totalCount;
    private IDisposable? _subscription;

    // ✅ 追踪所有订阅以便清理
    private readonly List<IDisposable> _indicatorSubscriptions = new();
    private NotifyCollectionChangedEventHandler? _collectionChangedHandler;
    private bool _disposed;

    public AnalysisScanViewModel(AnalysisBatch batch, List<Indicator> indicators)
    {
        _batch = batch ?? throw new ArgumentNullException(nameof(batch));
        
        if (indicators == null)
            throw new ArgumentNullException(nameof(indicators));

        // ✅ 在UI线程初始化
        InitializeIndicatorsSafe(indicators);

        // ✅ 使用 WhenAnyValue 监听 Indicators 集合的变化
        _subscription = this.WhenAnyValue(x => x.Indicators)
            .Select(indicatorsList => indicatorsList != null ? indicatorsList.ToList() : new List<Indicator>())
            .Select(list => list.Count(i => i.IsSelected))
            .Subscribe(count =>
            {
                // ✅ 订阅回调可能在后台线程，使用 ThreadingHelper
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    SelectedCount = count;
                    CanStartScan = count > 0;
                    StatusMessage = $"已选择 {count} / {TotalCount} 个指标";
                }).ConfigureAwait(false);
            });

        SetupIndicatorMonitoring();

        // ✅ 全选命令 - 修复
        SelectAllCommand = ReactiveCommand.Create(() =>
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                foreach (var item in Indicators)
                {
                    item.IsSelected = true;
                }
            }).ConfigureAwait(false);
            return Unit.Default; // ✅ 添加返回
        });

        // ✅ 取消全选命令 - 修复
        DeselectAllCommand = ReactiveCommand.Create(() =>
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                foreach (var item in Indicators)
                {
                    item.IsSelected = false;
                }
            }).ConfigureAwait(false);
            return Unit.Default; // ✅ 添加返回
        });

        // ✅ 开始扫描命令 - 修复
        StartScanCommand = ReactiveCommand.Create(() =>
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                var selected = Indicators.Where(i => i.IsSelected).ToList();
                if (selected.Count == 0)
                {
                    StatusMessage = "请至少选择一个指标";
                    return;
                }

                CanStartScan = false;
                OnStartScan?.Invoke(selected);
            }).ConfigureAwait(false);
            return Unit.Default; // ✅ 添加返回
        });

        // ✅ 取消命令 - 修复
        CancelCommand = ReactiveCommand.Create(() =>
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                CleanupSubscriptions();
                OnCancel?.Invoke();
            }).ConfigureAwait(false);
            return Unit.Default; // ✅ 添加返回
        });

        StatusMessage = $"共 {Indicators.Count} 个指标可供选择";
    }

    #region 属性

    public ObservableCollection<Indicator> Indicators
    {
        get => _indicators;
        set
        {
            // ✅ 确保在UI线程更新
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                // 清理所有旧订阅
                CleanupSubscriptions();
                _subscription?.Dispose();

                this.RaiseAndSetIfChanged(ref _indicators, value);

                // 重新设置监控
                SetupIndicatorMonitoring();

                TotalCount = Indicators.Count;
                UpdateSelectedCount();
            }).ConfigureAwait(false);
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
    }

    public bool CanStartScan
    {
        get => _canStartScan;
        set => this.RaiseAndSetIfChanged(ref _canStartScan, value);
    }

    public int SelectedCount
    {
        get => _selectedCount;
        set => this.RaiseAndSetIfChanged(ref _selectedCount, value);
    }

    public int TotalCount
    {
        get => _totalCount;
        set => this.RaiseAndSetIfChanged(ref _totalCount, value);
    }

    /// <summary>
    /// 筛选后的指标列表（根据分类筛选）
    /// </summary>
    private ObservableCollection<Indicator> _filteredIndicators = new();
    public ObservableCollection<Indicator> FilteredIndicators
    {
        get => _filteredIndicators;
        set => this.RaiseAndSetIfChanged(ref _filteredIndicators, value);
    }

    /// <summary>
    /// 分类筛选选项
    /// </summary>
    private ObservableCollection<string> _categoryFilterOptions = new() { "全部" };
    public ObservableCollection<string> CategoryFilterOptions
    {
        get => _categoryFilterOptions;
        set => this.RaiseAndSetIfChanged(ref _categoryFilterOptions, value);
    }

    /// <summary>
    /// 当前选中的分类筛选
    /// </summary>
    private string _selectedCategoryFilter = "全部";
    public string SelectedCategoryFilter
    {
        get => _selectedCategoryFilter;
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedCategoryFilter, value);
            ApplyCategoryFilter();
        }
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> SelectAllCommand { get; }
    public ReactiveCommand<Unit, Unit> DeselectAllCommand { get; }
    public ReactiveCommand<Unit, Unit> StartScanCommand { get; }
    public ReactiveCommand<Unit, Unit> CancelCommand { get; }

    #endregion

    #region 事件

    public Action<List<Indicator>>? OnStartScan { get; set; }
    public Action? OnCancel { get; set; }

    #endregion

    #region 公共方法

    public AnalysisBatch Batch => _batch;

    /// <summary>
    /// 安全初始化指标
    /// </summary>
    private void InitializeIndicatorsSafe(List<Indicator> indicators)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Indicators.Clear();
            foreach (var indicator in indicators)
            {
                Indicators.Add(indicator);
            }
            TotalCount = Indicators.Count;

            // 构建分类筛选选项
            var categories = indicators
                .Select(i => i.Category)
                .Where(c => !string.IsNullOrEmpty(c))
                .Distinct()
                .OrderBy(c => c)
                .ToList();
            CategoryFilterOptions = new ObservableCollection<string>(new[] { "全部" }.Concat(categories));
            SelectedCategoryFilter = "全部";

            // 初始显示全部
            ApplyCategoryFilter();
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 根据选中的分类筛选指标
    /// </summary>
    private void ApplyCategoryFilter()
    {
        if (string.IsNullOrEmpty(SelectedCategoryFilter) || SelectedCategoryFilter == "全部")
        {
            FilteredIndicators = new ObservableCollection<Indicator>(Indicators);
        }
        else
        {
            var filtered = Indicators
                .Where(i => i.Category == SelectedCategoryFilter)
                .ToList();
            FilteredIndicators = new ObservableCollection<Indicator>(filtered);
        }
    }

    /// <summary>
    /// 重置选择状态
    /// </summary>
    public void ResetSelection()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            foreach (var item in Indicators)
            {
                item.IsSelected = false;
            }
            UpdateSelectedCount();
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 选择所有指标
    /// </summary>
    public void SelectAll()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            foreach (var item in Indicators)
            {
                item.IsSelected = true;
            }
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 取消选择所有指标
    /// </summary>
    public void DeselectAll()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            foreach (var item in Indicators)
            {
                item.IsSelected = false;
            }
        }).ConfigureAwait(false);
    }

    #endregion

    #region 私有方法

    private void SetupIndicatorMonitoring()
    {
        // ✅ 确保在UI线程执行
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            // 先清理旧订阅
            CleanupSubscriptions();

            // 当 Indicators 集合变化时更新计数
            _collectionChangedHandler = (s, e) =>
            {
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    UpdateSelectedCount();
                    TotalCount = Indicators.Count;
                }).ConfigureAwait(false);
            };
            Indicators.CollectionChanged += _collectionChangedHandler;

            // 监听每个指标的 IsSelected 变化
            foreach (var indicator in Indicators)
            {
                var sub = indicator.WhenAnyValue(x => x.IsSelected)
                    .Subscribe(_ =>
                    {
                        ThreadingHelper.RunOnUIThreadAsync(() =>
                        {
                            UpdateSelectedCount();
                        }).ConfigureAwait(false);
                    });
                _indicatorSubscriptions.Add(sub);
            }
        }).ConfigureAwait(false);
    }

    private void CleanupSubscriptions()
    {
        // ✅ 确保在UI线程执行
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            // 移除 CollectionChanged 处理器
            if (_collectionChangedHandler != null)
            {
                Indicators.CollectionChanged -= _collectionChangedHandler;
                _collectionChangedHandler = null;
            }

            // 释放所有独立指标订阅
            foreach (var sub in _indicatorSubscriptions)
            {
                try
                {
                    sub.Dispose();
                }
                catch
                {
                }
            }
            _indicatorSubscriptions.Clear();
        }).ConfigureAwait(false);
    }

    private void UpdateSelectedCount()
    {
        // ✅ 确保在UI线程执行
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            try
            {
                SelectedCount = Indicators.Count(i => i.IsSelected);
                CanStartScan = SelectedCount > 0;
                StatusMessage = $"已选择 {SelectedCount} / {TotalCount} 个指标";
            }
            catch
            {
            }
        }).ConfigureAwait(false);
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            try
            {
                // 清理所有订阅
                CleanupSubscriptions();
                _subscription?.Dispose();
                
                // 清理命令
                SelectAllCommand?.Dispose();
                DeselectAllCommand?.Dispose();
                StartScanCommand?.Dispose();
                CancelCommand?.Dispose();
            }
            catch
            {
            }
        }

        _disposed = true;
    }

    #endregion
}