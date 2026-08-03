using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using System;
using System.Threading.Tasks;

namespace ExcelToSQLite.Views;

public partial class TableFieldManagementView : UserControl
{
    private TableFieldManagementViewModel? _viewModel;
    private bool _isCleaning;

    public TableFieldManagementView()
    {
        InitializeComponent();
        this.Unloaded += OnUnloaded;
    }

    public TableFieldManagementView(Window parentWindow) : this()
    {
        _viewModel = new TableFieldManagementViewModel();
        DataContext = _viewModel;
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private void OnUnloaded(object? sender, EventArgs e)
    {
        if (_isCleaning) return;
        _isCleaning = true;
        try
        {
            if (_viewModel is IDisposable d) d.Dispose();
            if (_viewModel is ICleanupPage cp) cp.Cleanup();
        }
        catch { }
        finally { _viewModel = null; DataContext = null; _isCleaning = false; }
        this.Unloaded -= OnUnloaded;
    }

    public void Cleanup()
    {
        if (_isCleaning) return;
        _isCleaning = true;
        try
        {
            if (_viewModel is IDisposable d) d.Dispose();
            if (_viewModel is ICleanupPage cp) cp.Cleanup();
        }
        catch { }
        finally { _viewModel = null; DataContext = null; }
        this.Unloaded -= OnUnloaded;
    }
}