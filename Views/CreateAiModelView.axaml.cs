using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using System;
using Avalonia.VisualTree;

namespace ExcelToSQLite.Views;

public partial class CreateAiModelView : UserControl
{
    private CreateAiModelViewModel? _viewModel;
    private Window? _parentWindow;

    public CreateAiModelView()
    {
        InitializeComponent();
    }

    public CreateAiModelView(Window? parentWindow) : this()
    {
        _parentWindow = parentWindow;
        _viewModel = new CreateAiModelViewModel(parentWindow);
        DataContext = _viewModel;
    }

    private void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }
    
    /// <summary>
    /// 清理资源
    /// </summary>
    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);
        CleanupViewModel();
        _parentWindow = null;
    }

    /// <summary>
    /// 清理 ViewModel 资源
    /// </summary>
    private void CleanupViewModel()
    {
        if (_viewModel == null)
            return;

        try
        {
            _viewModel.Cleanup();
        }
        catch
        {
        }

        _viewModel = null;
        DataContext = null;
    }

    /// <summary>
    /// 显式清理方法（可由父窗口调用）
    /// </summary>
    public void Cleanup()
    {
        CleanupViewModel();
        _parentWindow = null;
    }
}