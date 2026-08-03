using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ExcelToSQLite.Helpers;
using Avalonia.Media;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels
{
    public class TableFieldViewViewModel : ReactiveObject, ICleanupPage, IDisposable
    {
        private readonly DatabaseService _databaseService;
        private readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
        private readonly CompositeDisposable _subscriptions = new CompositeDisposable();
        private CancellationTokenSource? _cts;
        private bool _disposed;
        private bool _isCleaned;
        private bool _isLoading;
        private bool _showStatus;

        private ObservableCollection<string> _tableNames = new();
        private string? _selectedTable;
        private string _fieldInfo = string.Empty;
        private string _searchText = string.Empty;
        private string _statusMessage = string.Empty;
        private IBrush _statusColor = new SolidColorBrush(Colors.Green);
        private ObservableCollection<string> _filteredTableNames = new();

        #region 属性

        public ObservableCollection<string> TableNames
        {
            get => _tableNames;
            set => this.RaiseAndSetIfChanged(ref _tableNames, value);
        }

        public ObservableCollection<string> FilteredTableNames
        {
            get => _filteredTableNames;
            set => this.RaiseAndSetIfChanged(ref _filteredTableNames, value);
        }

        public string? SelectedTable
        {
            get => _selectedTable;
            set
            {
                this.RaiseAndSetIfChanged(ref _selectedTable, value);
                if (!string.IsNullOrEmpty(value))
                    _ = LoadTableFieldsAsync(value);
            }
        }

        public string FieldInfo
        {
            get => _fieldInfo;
            set => this.RaiseAndSetIfChanged(ref _fieldInfo, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => this.RaiseAndSetIfChanged(ref _isLoading, value);
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                this.RaiseAndSetIfChanged(ref _searchText, value);
                FilterTables();
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
        }

        public IBrush StatusColor
        {
            get => _statusColor;
            set => this.RaiseAndSetIfChanged(ref _statusColor, value);
        }

        public bool ShowStatus
        {
            get => _showStatus;
            set => this.RaiseAndSetIfChanged(ref _showStatus, value);
        }

        #endregion

        public ReactiveCommand<Unit, Unit> RefreshCommand { get; }

        public TableFieldViewViewModel() : this(DatabaseService.Instance) { }

        public TableFieldViewViewModel(DatabaseService databaseService)
        {
            _databaseService = databaseService ?? throw new ArgumentNullException(nameof(databaseService));
            RefreshCommand = ReactiveCommand.CreateFromTask(RefreshTablesAsync);
            RefreshCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex => SetStatus($"刷新失败: {ex.Message}", new SolidColorBrush(Colors.Red)))
                .DisposeWith(_subscriptions);
            _ = InitializeAsync();
        }

        private async Task InitializeAsync()
        {
            if (_disposed || _isCleaned) return;
            await _lock.WaitAsync();
            try
            {
                if (_disposed || _isCleaned) return;
                await RefreshTablesAsync();
            }
            catch (Exception ex) { await SetStatusAsync($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
            finally { _lock.Release(); }
        }

        private async Task RefreshTablesAsync()
        {
            if (_disposed || _isCleaned) return;
            CancelCurrentOperation();
            try
            {
                await SetLoadingAsync(true, "正在加载表列表...");
                _cts = new CancellationTokenSource();
                var token = _cts.Token;
                var tables = await _databaseService.GetAllTableNamesAsync();
                if (token.IsCancellationRequested || _disposed || _isCleaned) return;
                var userTables = tables
                    .Where(t => !t.StartsWith("sqlite_", StringComparison.OrdinalIgnoreCase))
                    .Where(t => !Models.TableNames.AllowedSet.Contains(t))
                    .OrderBy(t => t).ToList();
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    TableNames = new ObservableCollection<string>(userTables);
                    FilteredTableNames = new ObservableCollection<string>(userTables);
                    SelectedTable = null;
                    FieldInfo = string.Empty;
                    SearchText = string.Empty;
                    IsLoading = false;
                    SetStatus($"加载完成，共 {userTables.Count} 个表", new SolidColorBrush(Colors.Green));
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (!_disposed && !_isCleaned) { FieldInfo = $"加载失败: {ex.Message}"; IsLoading = false; }
                });
            }
        }

        private void FilterTables()
        {
            if (_disposed || _isCleaned) return;
            _ = ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                FilteredTableNames = string.IsNullOrWhiteSpace(SearchText)
                    ? new ObservableCollection<string>(TableNames)
                    : new ObservableCollection<string>(TableNames.Where(t => t.Contains(SearchText, StringComparison.OrdinalIgnoreCase)));
            });
        }

        private async Task LoadTableFieldsAsync(string tableName)
        {
            if (_disposed || _isCleaned || string.IsNullOrEmpty(tableName)) return;
            CancelCurrentOperation();
            try
            {
                await SetLoadingAsync(true, $"正在加载表 '{tableName}'...");
                _cts = new CancellationTokenSource();
                var token = _cts.Token;
                var fields = await GetTableFieldsAsync(tableName, token);
                if (token.IsCancellationRequested || _disposed || _isCleaned) return;
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    var result = new System.Text.StringBuilder();
                    result.AppendLine($"📋 表名: {tableName}");
                    result.AppendLine($"📊 字段数: {fields.Count}");
                    result.AppendLine(new string('=', 80));
                    result.AppendLine();
                    result.AppendLine($"{"序号",-6} {"字段名",-25} {"数据类型",-15} {"是否为空",-10} {"默认值",-15} {"主键",-6}");
                    result.AppendLine(new string('-', 80));
                    foreach (var field in fields)
                        result.AppendLine(string.Format("{0,-6} {1,-25} {2,-15} {3,-10} {4,-15} {5,-6}",
                            field.Cid, field.Name, field.Type,
                            field.NotNull ? "NOT NULL" : "NULL", field.DfltValue ?? "-", field.Pk > 0 ? "✓" : ""));
                    result.AppendLine(new string('=', 80));
                    result.AppendLine();
                    result.AppendLine("💡 提示: 可以选择下方内容并复制 (Ctrl+C)");
                    FieldInfo = result.ToString();
                    SetStatus($"表 '{tableName}' 加载完成，共 {fields.Count} 个字段", new SolidColorBrush(Colors.Green));
                    IsLoading = false;
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { await SetStatusAsync($"加载失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
        }

        private async Task<List<TableFieldInfo>> GetTableFieldsAsync(string tableName, CancellationToken token)
        {
            var fields = new List<TableFieldInfo>();
            var sql = $"PRAGMA table_info(\"{tableName}\")";
            var result = await _databaseService.ExecuteQueryAsync(sql, new List<object>());
            token.ThrowIfCancellationRequested();
            if (result == null || result.Count <= 1) return fields;
            for (int i = 1; i < result.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var row = result[i];
                if (row.Count >= 6)
                    fields.Add(new TableFieldInfo { Cid = Convert.ToInt32(row[0]), Name = row[1]?.ToString() ?? "", Type = row[2]?.ToString() ?? "", NotNull = Convert.ToInt32(row[3]) == 1, DfltValue = row[4]?.ToString(), Pk = Convert.ToInt32(row[5]) });
            }
            return fields;
        }

        private void CancelCurrentOperation() { try { _cts?.Cancel(); _cts?.Dispose(); _cts = null; } catch { } }

        private void SetStatus(string message, IBrush color)
        {
            if (!Dispatcher.UIThread.CheckAccess()) { Dispatcher.UIThread.Post(() => SetStatus(message, color)); return; }
            StatusMessage = message; StatusColor = color; ShowStatus = !string.IsNullOrEmpty(message);
        }

        private async Task SetStatusAsync(string message, IBrush color) => await ThreadingHelper.RunOnUIThreadAsync(() => SetStatus(message, color));
        private async Task SetLoadingAsync(bool isLoading, string? message = null) => await ThreadingHelper.RunOnUIThreadAsync(() => { IsLoading = isLoading; if (message != null) SetStatus(message, new SolidColorBrush(Colors.Orange)); });

        public void Cleanup()
        {
            if (_isCleaned) return;
            try
            {
                CancelCurrentOperation();
                _lock.Wait(TimeSpan.FromSeconds(2));
                try { _subscriptions?.Dispose(); _ = ThreadingHelper.RunOnUIThreadAsync(() => { try { TableNames?.Clear(); FilteredTableNames?.Clear(); FieldInfo = string.Empty; SelectedTable = null; } catch { } }); _isCleaned = true; }
                finally { try { _lock.Release(); } catch { } }
            }
            catch { _isCleaned = true; }
        }

        public void Dispose()
        {
            if (_disposed) return;
            Cleanup(); _lock?.Dispose(); _cts?.Dispose(); RefreshCommand?.Dispose();
            _disposed = true; GC.SuppressFinalize(this);
        }

        private class TableFieldInfo
        {
            public int Cid { get; set; }
            public string Name { get; set; } = string.Empty;
            public string Type { get; set; } = string.Empty;
            public bool NotNull { get; set; }
            public string? DfltValue { get; set; }
            public int Pk { get; set; }
        }
    }
}