using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class IndicatorEditDialog : UserControl
{
    public IndicatorEditDialog()
    {
        InitializeComponent();
        // 不在构造函数中创建 ViewModel，由外部设置
    }

    public IndicatorEditDialog(Window parentWindow) : this()
    {
        // 不在构造函数中创建 ViewModel，由外部设置
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}