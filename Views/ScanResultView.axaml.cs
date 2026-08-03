using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using System;
using System.Reactive.Disposables;

namespace ExcelToSQLite.Views;

public partial class ScanResultView : UserControl
{
    private ScanResultViewModel? _viewModel;
    private ListBox? _listBox;
    private CompositeDisposable? _subscriptions;

    public ScanResultView()
    {
        InitializeComponent();
        InitializeViewModel(null);
        SetupDoubleClickEvent();
    }

    // 重载构造函数，传入父窗口
    public ScanResultView(Window parentWindow)
    {
        InitializeComponent();
        InitializeViewModel(parentWindow);
        SetupDoubleClickEvent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    /// <summary>
    /// 初始化 ViewModel
    /// </summary>
    private void InitializeViewModel(Window? parentWindow)
    {
        // 清理旧的 ViewModel
        CleanupViewModel();

        _viewModel = new ScanResultViewModel();
        if (parentWindow != null)
        {
            _viewModel.SetParentWindow(parentWindow);
        }
        DataContext = _viewModel;
        
        // ✅ 订阅 Unloaded 事件
        this.Unloaded += OnUnloaded;
    }

    /// <summary>
    /// 设置父窗口（外部调用）
    /// </summary>
    public void SetParentWindow(Window parentWindow)
    {
        if (parentWindow == null)
        {
            return;
        }

        if (_viewModel != null)
        {
            _viewModel.SetParentWindow(parentWindow);
        }
    }

    /// <summary>
    /// 获取当前 ViewModel
    /// </summary>
    public ScanResultViewModel? ViewModel => _viewModel;

    private void SetupDoubleClickEvent()
    {
        _listBox = this.FindControl<ListBox>("ResultListBox");
        if (_listBox != null)
        {
            _listBox.DoubleTapped += OnListBoxDoubleTapped;
        }
    }

    private void OnListBoxDoubleTapped(object? sender, TappedEventArgs e)
    {
        // ✅ 使用 ThreadingHelper 确保在 UI 线程执行
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (_viewModel == null)
                return;

            var listBox = sender as ListBox;
            if (listBox?.SelectedItem is Models.ScanResultItem selectedItem)
            {
                // ✅ 使用 DisposeWith 管理订阅，防止内存泄漏
                _viewModel.ViewDetailCommand.Execute(selectedItem)
                    .Subscribe(
                        onNext: _ => { },
                        onError: ex =>
                        {
                        }
                    );
            }
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 清理 ViewModel 资源
    /// </summary>
    private void CleanupViewModel()
    {
        if (_viewModel == null)
            return;

        try
        {
            // 调用 Cleanup（如果存在）
            var cleanupMethod = _viewModel.GetType().GetMethod("Cleanup");
            if (cleanupMethod != null)
            {
                cleanupMethod.Invoke(_viewModel, null);
            }
            else if (_viewModel is IDisposable disposable)
            {
                disposable.Dispose();
            }
        }
        catch
        {
        }

        _viewModel = null;
        DataContext = null;
    }

    /// <summary>
    /// 清理事件订阅
    /// </summary>
    private void CleanupEventSubscriptions()
    {
        if (_listBox != null)
        {
            _listBox.DoubleTapped -= OnListBoxDoubleTapped;
            _listBox = null;
        }
        
        // 清理 Reactive 订阅
        _subscriptions?.Dispose();
        _subscriptions = null;
    }

    /// <summary>
    /// 当控件卸载时触发
    /// </summary>
    private void OnUnloaded(object? sender, EventArgs e)
    {
        CleanupEventSubscriptions();
        CleanupViewModel();
        this.Unloaded -= OnUnloaded;
    }

    /// <summary>
    /// 显式清理方法（可由父窗口调用）
    /// </summary>
    public void Cleanup()
    {
        CleanupEventSubscriptions();
        CleanupViewModel();
        this.Unloaded -= OnUnloaded;
    }
}