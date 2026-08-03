using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class TaxExcelDataView : UserControl
{
    public TaxExcelDataView()
    {
        InitializeComponent();
        var viewModel = new TaxExcelDataViewModel();
        DataContext = viewModel;
    }

    public TaxExcelDataView(Window parentWindow)
    {
        InitializeComponent();
        var viewModel = new TaxExcelDataViewModel();
        viewModel.SetParentWindow(parentWindow);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}