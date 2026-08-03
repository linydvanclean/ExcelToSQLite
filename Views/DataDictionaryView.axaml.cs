using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using System;

namespace ExcelToSQLite.Views;

public partial class DataDictionaryView : UserControl
{
    private DataDictionaryViewModel? _viewModel;

    public DataDictionaryView()
    {
        try
        {
            InitializeComponent();

            // ✅ 订阅 Unloaded 事件
            this.Unloaded += OnUnloaded;
        }
        catch
        {
            throw;
        }
    }

    public DataDictionaryView(Window parentWindow) : this()
    {
        try
        {
            _viewModel = new DataDictionaryViewModel();

            _viewModel.SetParentWindow(parentWindow);

            DataContext = _viewModel;
        }
        catch
        {
            throw;
        }
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    /// <summary>
    /// 当控件从视觉树卸载时触发
    /// </summary>
    private void OnUnloaded(object? sender, EventArgs e)
    {
        CleanupViewModel();
    }

    /// <summary>
    /// 获取当前 ViewModel
    /// </summary>
    public DataDictionaryViewModel? ViewModel => _viewModel;

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
        else
        {
            // 如果 ViewModel 还未创建，先创建
            _viewModel = new DataDictionaryViewModel();
            _viewModel.SetParentWindow(parentWindow);
            DataContext = _viewModel;
        }
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
            _viewModel.Cleanup();
        }
        catch
        {
        }

        _viewModel = null;
        DataContext = null;
    }

    /// <summary>
    /// 显式清理方法（可由父窗口调用）
    /// </summary>
    public void Cleanup()
    {
        CleanupViewModel();
        this.Unloaded -= OnUnloaded; // 取消订阅
    }
}