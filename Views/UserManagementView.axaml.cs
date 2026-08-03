using Avalonia.Controls;
using Avalonia.Markup.Xaml;

namespace ExcelToSQLite.Views;

public partial class UserManagementView : UserControl
{
    public UserManagementView()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}