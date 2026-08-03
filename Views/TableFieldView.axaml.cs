// Views/TableFieldView.axaml.cs - 只读表字段查看
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Markup.Xaml;
using ExcelToSQLite.ViewModels;
using ExcelToSQLite.Helpers;
using ReactiveUI;
using System;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Threading.Tasks;

namespace ExcelToSQLite.Views
{
    public partial class TableFieldView : UserControl
    {
        private TableFieldViewViewModel? _viewModel;
        private Button? _copyAllButton;
        private CompositeDisposable? _subscriptions;
        private bool _isCleaning;

        public TableFieldView()
        {
            InitializeComponent();
            this.Loaded += OnLoaded;
            this.Unloaded += OnUnloaded;
        }

        public TableFieldView(TableFieldViewViewModel viewModel) : this()
        {
            _viewModel = viewModel;
            DataContext = viewModel;
        }

        public TableFieldView(Window parentWindow) : this()
        {
            var viewModel = new TableFieldViewViewModel();
            DataContext = viewModel;
            _viewModel = viewModel;
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void SetViewModel(TableFieldViewViewModel viewModel)
        {
            CleanupViewModel();
            _viewModel = viewModel;
            DataContext = viewModel;
            SetupSubscriptions();
        }

        public TableFieldViewViewModel? ViewModel => _viewModel;

        private void OnLoaded(object? sender, EventArgs e)
        {
            _copyAllButton = this.FindControl<Button>("CopyAllButton");
            if (_copyAllButton != null)
                _copyAllButton.Click += OnCopyAllClick;
            SetupSubscriptions();
        }

        private void SetupSubscriptions()
        {
            if (_viewModel == null) return;
            _subscriptions?.Dispose();
            _subscriptions = new CompositeDisposable();

            if (_copyAllButton != null)
            {
                _copyAllButton.IsEnabled = !string.IsNullOrEmpty(_viewModel.SelectedTable);
                _viewModel.WhenAnyValue(x => x.SelectedTable)
                    .ObserveOn(RxApp.MainThreadScheduler)
                    .Subscribe(table =>
                    {
                        _ = ThreadingHelper.RunOnUIThreadAsync(() =>
                        {
                            if (_copyAllButton != null && !_isCleaning)
                                _copyAllButton.IsEnabled = !string.IsNullOrEmpty(table);
                        });
                    })
                    .DisposeWith(_subscriptions);
            }
        }

        private async void OnCopyAllClick(object? sender, EventArgs e)
        {
            if (_viewModel == null || string.IsNullOrEmpty(_viewModel.FieldInfo) || _isCleaning) return;

            try
            {
                var topLevel = TopLevel.GetTopLevel(this);
                if (topLevel?.Clipboard != null)
                {
                    await topLevel.Clipboard.SetTextAsync(_viewModel.FieldInfo);
                }
            }
            catch
            {
            }
        }

        private void CleanupViewModel()
        {
            if (_viewModel == null || _isCleaning) return;
            _isCleaning = true;
            try
            {
                if (_viewModel is IDisposable d) d.Dispose();
                if (_viewModel is ICleanupPage cp) cp.Cleanup();
            }
            catch { }
            finally { _viewModel = null; DataContext = null; _isCleaning = false; }
        }

        private void OnUnloaded(object? sender, EventArgs e)
        {
            _subscriptions?.Dispose(); _subscriptions = null;
            if (_copyAllButton != null) { _copyAllButton.Click -= OnCopyAllClick; _copyAllButton = null; }
            CleanupViewModel();
            this.Unloaded -= OnUnloaded;
        }

        public void Cleanup()
        {
            _subscriptions?.Dispose(); _subscriptions = null;
            if (_copyAllButton != null) { _copyAllButton.Click -= OnCopyAllClick; _copyAllButton = null; }
            CleanupViewModel();
            this.Unloaded -= OnUnloaded;
        }
    }
}