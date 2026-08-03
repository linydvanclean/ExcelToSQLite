using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using System;

namespace ExcelToSQLite.Views;

public partial class DataDictionaryEditDialog : Window
{
    public DataDictionaryEditDialog()
    {
        InitializeComponent();
    }

    public DataDictionaryEditDialog(DataDictionaryEditViewModel viewModel) : this()
    {
        // 设置对话框窗口引用
        viewModel.SetDialogWindow(this);
        DataContext = viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
        try
        {
            Icon = IconHelper.GetAppIcon();
        }
        catch
        {
        }
    }
}