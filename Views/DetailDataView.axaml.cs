using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using ReactiveUI;
using System;
using System.Linq;
using System.Reactive;
using System.Reactive.Disposables;
using System.Threading.Tasks;

namespace ExcelToSQLite.Views;

public partial class DetailDataView : UserControl
{
    private DetailDataViewModel? _viewModel;
    private DataGrid? _dataGrid;
    private Window? _parentWindow;
    private CompositeDisposable? _disposables;
    private bool _isCleanedUp;

    public DetailDataView()
    {
        InitializeComponent();
        this.DataContextChanged += OnDataContextChanged;
        // 如果需要保留 Unloaded 事件，使用正确的签名
        // this.Unloaded += OnUnloaded;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    private Window? GetParentWindow()
    {
        if (this.VisualRoot is Window window)
            return window;
        
        var parent = this.Parent;
        while (parent != null)
        {
            if (parent is Window win)
                return win;
            parent = parent.Parent;
        }
        
        return null;
    }

    public void SetParentWindow(Window? parentWindow)
    {
        if (parentWindow == null)
        {
            return;
        }

        _parentWindow = parentWindow;
    
        if (_viewModel != null)
        {
            _viewModel.SetParentWindow(parentWindow);
        }
    }

    public DetailDataViewModel? ViewModel => _viewModel;

    private void OnDataContextChanged(object? sender, EventArgs e)
    {
        if (DataContext is DetailDataViewModel viewModel)
        {
            if (_viewModel != null && _viewModel != viewModel)
            {
                CleanupOldViewModel();
            }
            
            _viewModel = viewModel;
            
            SetupParentWindow(viewModel);
            
            _dataGrid = this.FindControl<DataGrid>("DataGrid");
            
            viewModel.ColumnsUpdated += OnColumnsUpdated;
            
            _ = GenerateColumnsAsync(viewModel);
        }
    }

    private void SetupParentWindow(DetailDataViewModel viewModel)
    {
        if (_parentWindow != null)
        {
            viewModel.SetParentWindow(_parentWindow);
        }
        else
        {
            var parentWindow = GetParentWindow();
            if (parentWindow != null)
            {
                viewModel.SetParentWindow(parentWindow);
            }
            else
            {
                try
                {
                    if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                    {
                        var mainWindow = desktop.MainWindow;
                        if (mainWindow != null)
                        {
                            viewModel.SetParentWindow(mainWindow);
                        }
                    }
                }
                catch
                {
                }
            }
        }
    }

    private void OnColumnsUpdated(object? sender, EventArgs e)
    {
        if (sender is DetailDataViewModel viewModel)
        {
            _ = GenerateColumnsAsync(viewModel);
        }
    }

    private async Task GenerateColumnsAsync(DetailDataViewModel viewModel)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            GenerateColumnsInternal(viewModel);
        });
    }

    private void GenerateColumnsInternal(DetailDataViewModel viewModel)
    {
        if (_dataGrid == null)
        {
            return;
        }

        try
        {
            
            _dataGrid.Columns.Clear();
            _dataGrid.AutoGenerateColumns = false;

            var columnHeaders = viewModel.ColumnHeaders;
            if (columnHeaders == null || columnHeaders.Count == 0)
            {
                return;
            }

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

            foreach (var columnName in columnHeaders)
            {
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

            _dataGrid.ItemsSource = viewModel.TableData;
        }
        catch
        {
        }
    }

    private void CleanupOldViewModel()
    {
        if (_viewModel == null)
            return;

        try
        {

            _viewModel.ColumnsUpdated -= OnColumnsUpdated;

            _disposables?.Dispose();
            _disposables = null;
        }
        catch
        {
        }
    }

    private void CleanupViewModel()
    {
        if (_viewModel == null)
            return;

        try
        {

            _viewModel.ColumnsUpdated -= OnColumnsUpdated;

            _disposables?.Dispose();
            _disposables = null;
        }
        catch
        {
        }

        _viewModel = null;
        _dataGrid = null;
    }

    // ✅ 如果使用 Unloaded 事件，使用正确的签名
    private void OnUnloaded(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
    {
        CleanupViewModel();
        _parentWindow = null;
    }

    public void Cleanup()
    {
        if (_isCleanedUp) return;
        _isCleanedUp = true;
        
        CleanupViewModel();
        _parentWindow = null;
        this.Unloaded -= OnUnloaded;
    }
}