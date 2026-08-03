using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class FuelCardView : UserControl
{
    public FuelCardView()
    {
        InitializeComponent();
        var viewModel = new FuelCardViewModel();
        DataContext = viewModel;
    }

    // 重载构造函数，传入父窗口
    public FuelCardView(Window parentWindow)
    {
        InitializeComponent();
        var viewModel = new FuelCardViewModel();
        viewModel.SetParentWindow(parentWindow);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}