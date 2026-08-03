using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using System;

namespace ExcelToSQLite.Views
{
    public partial class ConfigEditView : UserControl
    {
        private ConfigEditViewModel? _viewModel;

        public ConfigEditView()
        {
            InitializeComponent();
        }

        public ConfigEditView(ConfigEditViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            DataContext = viewModel;
        }

        public ConfigEditView(Window parentWindow) : this()
        {
            var viewModel = new ConfigEditViewModel();
            DataContext = viewModel;
            _viewModel = viewModel;
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        public void SetViewModel(ConfigEditViewModel viewModel)
        {
            _viewModel = viewModel;
            DataContext = viewModel;
        }
    }
}