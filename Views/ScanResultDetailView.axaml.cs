using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using System;
using System.Threading.Tasks;

namespace ExcelToSQLite.Views;

public partial class ScanResultDetailView : UserControl
{
    private ScanResultDetailViewModel? _viewModel;
    private DataGrid? _dataGrid;

    public ScanResultDetailView()
    {
        InitializeComponent();
        this.DataContextChanged += OnDataContextChanged;
        
        // ✅ 订阅 Unloaded 事件
        this.Unloaded += OnUnloaded;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    /// <summary>
    /// 获取父窗口
    /// </summary>
    private Window? GetParentWindow()
    {
        // 方法1：使用 VisualRoot
        if (this.VisualRoot is Window window)
            return window;
        
        // 方法2：遍历父级
        var parent = this.Parent;
        while (parent != null)
        {
            if (parent is Window win)
                return win;
            parent = parent.Parent;
        }
        
        return null;
    }

    /// <summary>
    /// 获取当前 ViewModel
    /// </summary>
    public ScanResultDetailViewModel? ViewModel => _viewModel;

    private void OnDataContextChanged(object? sender, EventArgs e)
    {
        if (DataContext is ScanResultDetailViewModel viewModel)
        {
            // ✅ 清理旧的 ViewModel
            if (_viewModel != null && _viewModel != viewModel)
            {
                CleanupOldViewModel();
            }
            
            _viewModel = viewModel;
            
            // 设置父窗口
            var parentWindow = GetParentWindow();
            if (parentWindow != null)
            {
                viewModel.SetParentWindow(parentWindow);
            }
            else
            {
            }
        
            // 设置关闭回调
            viewModel.OnClose = () =>
            {
                var window = GetParentWindow();
                window?.Close();
            };
            
            // 获取 DataGrid 引用
            _dataGrid = this.FindControl<DataGrid>("DataGrid");
            
            // 订阅列头变化事件
            viewModel.ColumnsUpdated += OnColumnsUpdated;
            
            // ✅ 初始化生成列（确保在 UI 线程）
            _ = GenerateColumnsAsync(viewModel);
        }
    }

    private void OnColumnsUpdated(object? sender, EventArgs e)
    {
        if (sender is ScanResultDetailViewModel viewModel)
        {
            // ✅ 确保在 UI 线程执行
            _ = GenerateColumnsAsync(viewModel);
        }
    }
    
    /// <summary>
    /// 异步生成列（确保在 UI 线程执行）
    /// </summary>
    private async Task GenerateColumnsAsync(ScanResultDetailViewModel viewModel)
    {
        // ✅ 确保在 UI 线程执行
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            GenerateColumnsInternal(viewModel);
        });
    }

    /// <summary>
    /// 内部生成列方法（必须在 UI 线程调用）
    /// </summary>
    private void GenerateColumnsInternal(ScanResultDetailViewModel viewModel)
    {
        if (_dataGrid == null)
        {
            return;
        }

        try
        {
            // 清除自动生成的列
            _dataGrid.Columns.Clear();
            _dataGrid.AutoGenerateColumns = false;

            var columnHeaders = viewModel.ColumnHeaders;
            if (columnHeaders == null || columnHeaders.Count == 0)
            {
                return;
            }

            // 添加序号列 - 固定宽度
            var indexColumn = new DataGridTextColumn
            {
                Header = "#",
                Binding = new Avalonia.Data.Binding("Index"),
                Width = new DataGridLength(60, DataGridLengthUnitType.Pixel),
                IsReadOnly = true,
                CanUserResize = true,
                CanUserSort = true,
                MinWidth = 50,
                MaxWidth = 80
            };
            _dataGrid.Columns.Add(indexColumn);

            // 添加数据列 - 使用更合理的宽度
            int totalColumns = columnHeaders.Count;
            foreach (var columnName in columnHeaders)
            {
                // 根据列名长度和内容估算列宽
                int estimatedWidth = Math.Max(120, Math.Min(300, columnName.Length * 12 + 40));
                
                var binding = new Avalonia.Data.Binding($"[{columnName}]");
                
                var column = new DataGridTextColumn
                {
                    Header = columnName,
                    Binding = binding,
                    Width = new DataGridLength(estimatedWidth, DataGridLengthUnitType.Pixel),
                    IsReadOnly = true,
                    CanUserResize = true,
                    CanUserSort = true,
                    MinWidth = 80,
                    MaxWidth = 500
                };
                
                _dataGrid.Columns.Add(column);
            }

            // 绑定数据
            _dataGrid.ItemsSource = viewModel.TableData;
        }
        catch
        {
        }
    }

    /// <summary>
    /// 清理旧的 ViewModel
    /// </summary>
    private void CleanupOldViewModel()
    {
        if (_viewModel == null)
            return;

        try
        {
            // 取消事件订阅
            _viewModel.ColumnsUpdated -= OnColumnsUpdated;
            
            // 尝试调用 Cleanup（如果存在）
            CleanupViewModelInternal(_viewModel);
        }
        catch
        {
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
            // 取消事件订阅
            _viewModel.ColumnsUpdated -= OnColumnsUpdated;

            // 尝试调用 Cleanup（如果存在）
            CleanupViewModelInternal(_viewModel);
        }
        catch
        {
        }

        _viewModel = null;
        _dataGrid = null;
    }

    /// <summary>
    /// 内部清理 ViewModel 方法
    /// </summary>
    private void CleanupViewModelInternal(ScanResultDetailViewModel viewModel)
    {
        if (viewModel == null)
            return;

        // 方式1：尝试调用 Cleanup 方法（通过反射）
        var cleanupMethod = viewModel.GetType().GetMethod("Cleanup");
        if (cleanupMethod != null)
        {
            cleanupMethod.Invoke(viewModel, null);
        }
        // 方式2：如果实现了 IDisposable
        else if (viewModel is IDisposable disposable)
        {
            disposable.Dispose();
        }
        else
        {
        }
    }

    /// <summary>
    /// 当控件卸载时触发
    /// </summary>
    private void OnUnloaded(object? sender, EventArgs e)
    {
        CleanupViewModel();
        this.Unloaded -= OnUnloaded;
    }

    /// <summary>
    /// 清理资源
    /// </summary>
    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);
        CleanupViewModel();
    }

    /// <summary>
    /// 显式清理方法（可由父窗口调用）
    /// </summary>
    public void Cleanup()
    {
        CleanupViewModel();
        this.Unloaded -= OnUnloaded;
    }
}