using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ExcelToSQLite.Views;

public partial class AnalysisScanDialog : UserControl
{
    public AnalysisScanDialog()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}