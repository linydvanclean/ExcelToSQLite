// Views/DbTablesView.axaml.cs
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using System;

namespace ExcelToSQLite.Views
{
    public partial class DbTablesView : UserControl
    {
        private DbTablesViewModel? _viewModel;
        private bool _isCleanedUp;
        private Window? _parentWindow;  // ✅ 添加父窗口字段

        public DbTablesView()
        {
            InitializeComponent();
        }

        public DbTablesView(Window parentWindow) : this()
        {
            _parentWindow = parentWindow;  // ✅ 保存父窗口
            var viewModel = new DbTablesViewModel(parentWindow);  // ✅ 传入父窗口
            DataContext = viewModel;
            _viewModel = viewModel;
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public DbTablesViewModel? ViewModel => _viewModel;

        public void SetViewModel(DbTablesViewModel viewModel)
        {
            CleanupViewModel();
            _viewModel = viewModel;
            DataContext = viewModel;
        }

        public void SetParentWindow(Window? parentWindow)
        {
            _parentWindow = parentWindow;
            if (_viewModel != null)
            {
                _viewModel.SetParentWindow(parentWindow);
            }
        }

        private void CleanupViewModel()
        {
            if (_viewModel == null)
                return;

            try
            {

                if (_viewModel is IDisposable disposable)
                {
                    disposable.Dispose();
                }

                if (_viewModel is ICleanupPage cleanupPage)
                {
                    cleanupPage.Cleanup();
                }
            }
            catch
            {
            }

            _viewModel = null;
            DataContext = null;
        }

        private void OnUnloaded(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            CleanupViewModel();
            this.Unloaded -= OnUnloaded;
        }

        public void Cleanup()
        {
            if (_isCleanedUp) return;
            _isCleanedUp = true;
            
            CleanupViewModel();
            this.Unloaded -= OnUnloaded;
        }
    }
}