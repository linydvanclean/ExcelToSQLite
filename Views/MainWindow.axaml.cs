using Avalonia.Controls;
using ExcelToSQLite.ViewModels;
using System;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel(this);
    }
    public MainWindow(string username)
    {
        InitializeComponent();
        DataContext = new MainViewModel(this);
    }
    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
        this.SetAppIcon();
    }
}