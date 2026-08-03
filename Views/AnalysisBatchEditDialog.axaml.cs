using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ExcelToSQLite.Views;

public partial class AnalysisBatchEditDialog : UserControl
{
    public AnalysisBatchEditDialog()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}