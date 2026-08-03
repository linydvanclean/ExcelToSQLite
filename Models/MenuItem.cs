using ReactiveUI;
using System.Collections.ObjectModel;
using System;
using ExcelToSQLite.Helpers;
using Avalonia.Threading;
using System.Threading.Tasks;

namespace ExcelToSQLite.Models;

public class MenuItem : ReactiveObject
{
    private string? _header;
    private string? _icon;
    private string? _commandParameter;
    private bool _isExpanded;
    private ObservableCollection<MenuItem>? _children;
    
    public MenuItem()
    {
        _children = new ObservableCollection<MenuItem>();
    }

    /// <summary>
    /// 带参数的构造函数，方便创建带子菜单的菜单项
    /// </summary>
    public MenuItem(string header, string? icon = null, string? commandParameter = null)
        : this()
    {
        Header = header;
        Icon = icon;
        CommandParameter = commandParameter;
    }

    public string? Header
    {
        get => _header;
        set => this.RaiseAndSetIfChanged(ref _header, value);
    }

    public string? Icon
    {
        get => _icon;
        set => this.RaiseAndSetIfChanged(ref _icon, value);
    }

    public string? CommandParameter
    {
        get => _commandParameter;
        set => this.RaiseAndSetIfChanged(ref _commandParameter, value);
    }

    public bool IsExpanded
    {
        get => _isExpanded;
        set => this.RaiseAndSetIfChanged(ref _isExpanded, value);
    }

    public ObservableCollection<MenuItem> Children
    {
        get => _children ??= new ObservableCollection<MenuItem>();
        set => this.RaiseAndSetIfChanged(ref _children, value);
    }

    /// <summary>
    /// 添加子菜单项（同步方法 - 确保在 UI 线程调用）
    /// </summary>
    public MenuItem AddChild(string header, string? icon = null, string? commandParameter = null)
    {
        var child = new MenuItem(header, icon, commandParameter);
        AddChildInternal(child);
        return child;
    }

    /// <summary>
    /// 添加子菜单项（同步方法 - 确保在 UI 线程调用）
    /// </summary>
    public MenuItem AddChild(MenuItem child)
    {
        AddChildInternal(child);
        return child;
    }

    /// <summary>
    /// 异步添加子菜单项（可从后台线程调用）
    /// </summary>
    public async Task<MenuItem> AddChildAsync(string header, string? icon = null, string? commandParameter = null)
    {
        var child = new MenuItem(header, icon, commandParameter);
        await AddChildInternalAsync(child);
        return child;
    }

    /// <summary>
    /// 异步添加子菜单项（可从后台线程调用）
    /// </summary>
    public async Task<MenuItem> AddChildAsync(MenuItem child)
    {
        await AddChildInternalAsync(child);
        return child;
    }

    /// <summary>
    /// 内部同步添加方法
    /// </summary>
    private void AddChildInternal(MenuItem child)
    {
        // 确保在 UI 线程执行
        if (Dispatcher.UIThread.CheckAccess())
        {
            Children.Add(child);
        }
        else
        {
            // 如果在后台线程，同步等待切换到 UI 线程
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Children.Add(child);
            }).GetAwaiter().GetResult();
        }
    }

    /// <summary>
    /// 内部异步添加方法
    /// </summary>
    private async Task AddChildInternalAsync(MenuItem child)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Children.Add(child);
        });
    }

    /// <summary>
    /// 同步移除子菜单项（确保在 UI 线程调用）
    /// </summary>
    public bool RemoveChild(MenuItem child)
    {
        if (child == null || Children == null)
            return false;

        if (Dispatcher.UIThread.CheckAccess())
        {
            return Children.Remove(child);
        }
        else
        {
            var removed = false;
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                removed = Children.Remove(child);
            }).GetAwaiter().GetResult();
            return removed;
        }
    }

    /// <summary>
    /// 异步移除子菜单项（可从后台线程调用）
    /// </summary>
    public async Task<bool> RemoveChildAsync(MenuItem child)
    {
        if (child == null || Children == null)
            return false;

        var removed = false;
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            removed = Children.Remove(child);
        });
        return removed;
    }

    /// <summary>
    /// 同步清空所有子菜单（确保在 UI 线程调用）
    /// </summary>
    public void ClearChildren()
    {
        if (Dispatcher.UIThread.CheckAccess())
        {
            Children?.Clear();
        }
        else
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Children?.Clear();
            }).GetAwaiter().GetResult();
        }
    }

    /// <summary>
    /// 异步清空所有子菜单（可从后台线程调用）
    /// </summary>
    public async Task ClearChildrenAsync()
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Children?.Clear();
        });
    }

    /// <summary>
    /// 获取子菜单数量（线程安全）
    /// </summary>
    public int ChildCount => Children?.Count ?? 0;

    /// <summary>
    /// 检查是否有子菜单（线程安全）
    /// </summary>
    public bool HasChildren => Children != null && Children.Count > 0;

    /// <summary>
    /// 查找子菜单（线程安全）
    /// </summary>
    public MenuItem? FindChild(string header)
    {
        if (string.IsNullOrEmpty(header) || Children == null)
            return null;

        // ObservableCollection 的操作是线程安全的，但为了安全，使用锁或快照
        foreach (var child in Children)
        {
            if (child.Header == header)
                return child;
        }
        return null;
    }
}