using Avalonia;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;

namespace ExcelToSQLite.Views;

public partial class ChangePasswordDialog : UserControl
{
    // 添加默认构造函数以消除 AVLN3001 警告
    public ChangePasswordDialog()
    {
        InitializeComponent();
    }

    public ChangePasswordDialog(string username) : this()
    {
        var viewModel = new ChangePasswordViewModel(username);
        DataContext = viewModel;
    }

    public ChangePasswordDialog(Window parentWindow, string username) : this()
    {
        var viewModel = new ChangePasswordViewModel(username);
        viewModel.SetParentWindow(parentWindow);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
}