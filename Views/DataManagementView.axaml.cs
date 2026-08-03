using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class DataManagementView : UserControl
{
    public DataManagementView()
    {
        InitializeComponent();
        DataContext = new DataManagementViewModel();
    }

    public DataManagementView(Window parentWindow)
    {
        InitializeComponent();
        var vm = new DataManagementViewModel();
        vm.SetParentWindow(parentWindow);
        DataContext = vm;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}