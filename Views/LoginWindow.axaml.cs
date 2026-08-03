using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.Views;

public partial class LoginWindow : Window
{
    public LoginWindow()
    {
        InitializeComponent();
        DataContext = new LoginViewModel(this);
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
        this.SetAppIcon();
    }
}