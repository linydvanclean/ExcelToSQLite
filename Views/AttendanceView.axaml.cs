using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class AttendanceView : UserControl
{
    public AttendanceView()
    {
        InitializeComponent();
        var viewModel = new AttendanceViewModel();
        DataContext = viewModel;
    }

    public AttendanceView(Window parentWindow)
    {
        InitializeComponent();
        var viewModel = new AttendanceViewModel();
        viewModel.SetParentWindow(parentWindow);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}