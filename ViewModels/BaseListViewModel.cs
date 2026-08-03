using ReactiveUI;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System;
using System.Collections.Generic; 
using System.Linq;
using System.Threading.Tasks;
using ExcelToSQLite.Helpers;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 带列表的 ViewModel 基类
/// </summary>
public abstract class BaseListViewModel<T> : BaseViewModel where T : class
{
    private ObservableCollection<T> _items = new();
    private int _itemsCount = 0;
    private bool _hasItems = false;
    private T? _selectedItem;
    private bool _isLoadingItems = false;

    protected BaseListViewModel()
    {
        // 初始化时订阅集合变化
        _items.CollectionChanged += OnItemsCollectionChanged;
        UpdateItemsState();
    }

    #region 属性

    /// <summary>
    /// 数据集合
    /// </summary>
    public ObservableCollection<T> Items
    {
        get => _items;
        set
        {
            // ✅ 确保在UI线程更新
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                var oldItems = _items;
                if (oldItems != null)
                {
                    oldItems.CollectionChanged -= OnItemsCollectionChanged;
                }

                this.RaiseAndSetIfChanged(ref _items, value);

                var newItems = _items;
                if (newItems != null)
                {
                    newItems.CollectionChanged += OnItemsCollectionChanged;
                    UpdateItemsState();
                }
            }).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// 项目数量
    /// </summary>
    public int ItemsCount
    {
        get => _itemsCount;
        private set => this.RaiseAndSetIfChanged(ref _itemsCount, value);
    }

    /// <summary>
    /// 是否有项目
    /// </summary>
    public bool HasItems
    {
        get => _hasItems;
        private set => this.RaiseAndSetIfChanged(ref _hasItems, value);
    }

    /// <summary>
    /// 是否为空
    /// </summary>
    public bool IsEmpty => !HasItems;

    /// <summary>
    /// 选中的项目
    /// </summary>
    public T? SelectedItem
    {
        get => _selectedItem;
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedItem, value);
            OnSelectedItemChanged(value);
        }
    }

    /// <summary>
    /// 是否正在加载项目
    /// </summary>
    public bool IsLoadingItems
    {
        get => _isLoadingItems;
        set => this.RaiseAndSetIfChanged(ref _isLoadingItems, value);
    }

    #endregion

    #region 集合操作

    /// <summary>
    /// 添加项目（线程安全）
    /// </summary>
    protected async Task AddItemAsync(T item)
    {
        if (item == null) return;

        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (!Items.Contains(item))
            {
                Items.Add(item);
            }
        });
    }

    /// <summary>
    /// 添加项目（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected void AddItem(T item)
    {
        if (item == null) return;
        
        if (!Dispatcher.UIThread.CheckAccess())
        {
            throw new InvalidOperationException("AddItem 必须在UI线程调用，请使用 AddItemAsync");
        }

        if (!Items.Contains(item))
        {
            Items.Add(item);
        }
    }

    /// <summary>
    /// 移除项目（线程安全）
    /// </summary>
    protected async Task RemoveItemAsync(T item)
    {
        if (item == null) return;

        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (Items.Contains(item))
            {
                Items.Remove(item);
            }
        });
    }

    /// <summary>
    /// 移除项目（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected void RemoveItem(T item)
    {
        if (item == null) return;

        if (!Dispatcher.UIThread.CheckAccess())
        {
            throw new InvalidOperationException("RemoveItem 必须在UI线程调用，请使用 RemoveItemAsync");
        }

        if (Items.Contains(item))
        {
            Items.Remove(item);
        }
    }

    /// <summary>
    /// 清空集合（线程安全）
    /// </summary>
    protected async Task ClearItemsAsync()
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Items.Clear();
        });
    }

    /// <summary>
    /// 清空集合（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected void ClearItems()
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            throw new InvalidOperationException("ClearItems 必须在UI线程调用，请使用 ClearItemsAsync");
        }

        Items.Clear();
    }

    /// <summary>
    /// 添加范围（线程安全）
    /// </summary>
    protected async Task AddRangeAsync(IEnumerable<T> items)
    {
        if (items == null) return;

        var itemList = items.ToList();
        if (itemList.Count == 0) return;

        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            foreach (var item in itemList)
            {
                if (!Items.Contains(item))
                {
                    Items.Add(item);
                }
            }
        });
    }

    /// <summary>
    /// 添加范围（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected void AddRange(IEnumerable<T> items)
    {
        if (items == null) return;

        if (!Dispatcher.UIThread.CheckAccess())
        {
            throw new InvalidOperationException("AddRange 必须在UI线程调用，请使用 AddRangeAsync");
        }

        var itemList = items.ToList();
        foreach (var item in itemList)
        {
            if (!Items.Contains(item))
            {
                Items.Add(item);
            }
        }
    }

    /// <summary>
    /// 替换所有项目（线程安全）
    /// </summary>
    protected async Task ReplaceItemsAsync(IEnumerable<T> items)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Items.Clear();
            if (items != null)
            {
                foreach (var item in items)
                {
                    Items.Add(item);
                }
            }
        });
    }

    /// <summary>
    /// 替换所有项目（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected void ReplaceItems(IEnumerable<T> items)
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            throw new InvalidOperationException("ReplaceItems 必须在UI线程调用，请使用 ReplaceItemsAsync");
        }

        Items.Clear();
        if (items != null)
        {
            foreach (var item in items)
            {
                Items.Add(item);
            }
        }
    }

    /// <summary>
    /// 获取项目数量（线程安全）
    /// </summary>
    protected async Task<int> GetItemsCountAsync()
    {
        return await ThreadingHelper.RunOnUIThreadAsync(() => Items.Count);
    }

    #endregion

    #region 事件处理

    /// <summary>
    /// 集合变化事件处理
    /// </summary>
    protected virtual void OnItemsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        // ✅ 确保在UI线程更新状态
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            UpdateItemsState();
            OnItemsChanged();
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 项目变化时调用（子类可重写）
    /// </summary>
    protected virtual void OnItemsChanged()
    {
        // 子类可以重写
    }

    /// <summary>
    /// 选中项目变化时调用（子类可重写）
    /// </summary>
    protected virtual void OnSelectedItemChanged(T? item)
    {
        // 子类可以重写
    }

    #endregion

    #region 状态更新

    /// <summary>
    /// 更新项目状态
    /// </summary>
    private void UpdateItemsState()
    {
        // 确保在UI线程
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(UpdateItemsState);
            return;
        }

        ItemsCount = Items.Count;
        HasItems = ItemsCount > 0;
    }

    /// <summary>
    /// 设置加载状态（线程安全）
    /// </summary>
    protected async Task SetItemsLoadingStateAsync(bool isLoading, string? message = null)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsLoadingItems = isLoading;
            if (message != null)
            {
                StatusMessage = message;
            }
        });
    }

    #endregion

    #region 清理

    /// <summary>
    /// 清空集合（重写基类方法）
    /// </summary>
    protected override void ClearCollections()
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(ClearCollections);
            return;
        }

        try
        {
            if (_items != null)
            {
                _items.CollectionChanged -= OnItemsCollectionChanged;
                _items.Clear();
                _items.CollectionChanged += OnItemsCollectionChanged;
                UpdateItemsState();
            }
        }
        catch
        {
        }
    }

    #endregion

    #region 资源释放

    /// <summary>
    /// 释放资源
    /// </summary>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            try
            {
                if (_items != null)
                {
                    _items.CollectionChanged -= OnItemsCollectionChanged;
                }
                
                SelectedItem = null;
            }
            catch
            {
            }
        }
        base.Dispose(disposing);
    }

    #endregion
}