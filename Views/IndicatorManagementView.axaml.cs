using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using System;
using ExcelToSQLite.Services;

namespace ExcelToSQLite.Views;

public partial class IndicatorManagementView : UserControl
{
    public IndicatorManagementView()
    {
        InitializeComponent();
    }

    public IndicatorManagementView(Window parentWindow) : this()
    {
        try
        {
            var viewModel = new IndicatorManagementViewModel();
            viewModel.SetParentWindow(parentWindow);
            DataContext = viewModel;
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
}