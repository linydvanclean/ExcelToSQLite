using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class AnalysisBatchView : UserControl
{
    public AnalysisBatchView()
    {
        InitializeComponent();
        DataContext = new AnalysisBatchViewModel();
    }

    public AnalysisBatchView(Window parentWindow)
    {
        InitializeComponent();
        var vm = new AnalysisBatchViewModel();
        vm.SetParentWindow(parentWindow);
        DataContext = vm;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}