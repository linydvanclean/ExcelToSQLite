using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class ExcelDataView : UserControl
{
    public ExcelDataView()
    {
        InitializeComponent();
        var viewModel = new ExcelDataViewModel();
        DataContext = viewModel;
    }

    // 重载构造函数，传入父窗口
    public ExcelDataView(Window parentWindow)
    {
        InitializeComponent();
        var viewModel = new ExcelDataViewModel();
        viewModel.SetParentWindow(parentWindow);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}