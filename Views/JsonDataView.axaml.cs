using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class JsonDataView : UserControl
{
    public JsonDataView()
    {
        InitializeComponent();
        var viewModel = new JsonDataViewModel();
        DataContext = viewModel;
    }

    // 重载构造函数，传入父窗口
    public JsonDataView(Window parentWindow)
    {
        InitializeComponent();
        var viewModel = new JsonDataViewModel();
        viewModel.SetParentWindow(parentWindow);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}