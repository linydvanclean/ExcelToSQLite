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
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;

namespace ExcelToSQLite.ViewModels
{
    public class TableFieldManagementViewModel : ReactiveObject, ICleanupPage, IDisposable
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
        private ObservableCollection<EditableFieldInfo> _editableFields = new();
        private bool _isEditingFields;
        private bool _isTableSelected;

        #region 属性

        public ObservableCollection<string> TableNames { get => _tableNames; set => this.RaiseAndSetIfChanged(ref _tableNames, value); }
        public ObservableCollection<string> FilteredTableNames { get => _filteredTableNames; set => this.RaiseAndSetIfChanged(ref _filteredTableNames, value); }
        public string? SelectedTable
        {
            get => _selectedTable;
            set
            {
                this.RaiseAndSetIfChanged(ref _selectedTable, value);
                IsTableSelected = !string.IsNullOrEmpty(value);
                if (!string.IsNullOrEmpty(value))
                    _ = LoadTableFieldsAsync(value);
            }
        }
        public string FieldInfo { get => _fieldInfo; set => this.RaiseAndSetIfChanged(ref _fieldInfo, value); }
        public bool IsLoading { get => _isLoading; set => this.RaiseAndSetIfChanged(ref _isLoading, value); }
        public string SearchText { get => _searchText; set { this.RaiseAndSetIfChanged(ref _searchText, value); FilterTables(); } }
        public string StatusMessage { get => _statusMessage; set => this.RaiseAndSetIfChanged(ref _statusMessage, value); }
        public IBrush StatusColor { get => _statusColor; set => this.RaiseAndSetIfChanged(ref _statusColor, value); }
        public bool ShowStatus { get => _showStatus; set => this.RaiseAndSetIfChanged(ref _showStatus, value); }

        public ObservableCollection<EditableFieldInfo> EditableFields { get => _editableFields; set => this.RaiseAndSetIfChanged(ref _editableFields, value); }
        public bool IsEditingFields { get => _isEditingFields; set => this.RaiseAndSetIfChanged(ref _isEditingFields, value); }
        public bool IsTableSelected { get => _isTableSelected; set => this.RaiseAndSetIfChanged(ref _isTableSelected, value); }

        #endregion

        public ReactiveCommand<Unit, Unit> RefreshCommand { get; }
        public ReactiveCommand<Unit, Unit> CopyFieldInfoCommand { get; }
        public ReactiveCommand<Unit, Unit> StartEditFieldsCommand { get; }
        public ReactiveCommand<Unit, Unit> SaveFieldChangesCommand { get; }
        public ReactiveCommand<Unit, Unit> CancelEditFieldsCommand { get; }

        public TableFieldManagementViewModel() : this(DatabaseService.Instance) { }

        public TableFieldManagementViewModel(DatabaseService databaseService)
        {
            _databaseService = databaseService ?? throw new ArgumentNullException(nameof(databaseService));
            RefreshCommand = ReactiveCommand.CreateFromTask(RefreshTablesAsync);
            CopyFieldInfoCommand = ReactiveCommand.CreateFromTask(CopyFieldInfoAsync);
            StartEditFieldsCommand = ReactiveCommand.Create(StartEditFields);
            SaveFieldChangesCommand = ReactiveCommand.CreateFromTask(SaveFieldChangesAsync);
            CancelEditFieldsCommand = ReactiveCommand.Create(CancelEditFields);

            RefreshCommand.ThrownExceptions
                .ObserveOn(RxApp.MainThreadScheduler)
                .Subscribe(ex => SetStatus($"刷新失败: {ex.Message}", new SolidColorBrush(Colors.Red)))
                .DisposeWith(_subscriptions);

            _ = InitializeAsync();
        }

        #region 初始化

        private async Task InitializeAsync()
        {
            if (_disposed || _isCleaned) return;
            await _lock.WaitAsync();
            try { if (!_disposed && !_isCleaned) await RefreshTablesAsync(); }
            catch (Exception ex) { await SetStatusAsync($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
            finally { _lock.Release(); }
        }

        #endregion

        #region 表加载

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
                    SelectedTable = null; EditableFields.Clear(); IsEditingFields = false;
                    IsLoading = false;
                    SetStatus($"加载完成，共 {userTables.Count} 个表", new SolidColorBrush(Colors.Green));
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { await SetStatusAsync($"加载失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
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

        #endregion

        #region 字段加载

        private async Task LoadTableFieldsAsync(string tableName)
        {
            if (_disposed || _isCleaned || string.IsNullOrEmpty(tableName)) return;
            CancelCurrentOperation();
            try
            {
                await SetLoadingAsync(true, $"正在加载表 '{tableName}'...");
                _cts = new CancellationTokenSource();
                var token = _cts.Token;
                var fields = await GetTableColumnsAsync(tableName, token);
                if (token.IsCancellationRequested || _disposed || _isCleaned) return;
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    var editableList = new ObservableCollection<EditableFieldInfo>();
                    for (int i = 0; i < fields.Count; i++)
                        editableList.Add(new EditableFieldInfo { Ordinal = i + 1, FieldName = fields[i].Name, IsReadOnly = !IsEditingFields });
                    EditableFields = editableList;

                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine($"📋 表名: {tableName}");
                    sb.AppendLine($"📊 字段数: {fields.Count}");
                    sb.AppendLine(new string('=', 80));
                    sb.AppendLine();
                    sb.AppendLine($"{"序号",-6} {"字段名",-25} {"数据类型",-15}{"主键",-6}");
                    sb.AppendLine(new string('-', 80));
                    foreach (var f in fields)
                        sb.AppendLine(string.Format("{0,-6} {1,-25} {2,-15}{3,-6}", f.Cid, f.Name, f.Type, f.Pk > 0 ? "✓" : ""));
                    sb.AppendLine(new string('=', 80));
                    FieldInfo = sb.ToString();

                    SetStatus($"表 '{tableName}' 加载完成，共 {fields.Count} 个字段", new SolidColorBrush(Colors.Green));
                    IsLoading = false;
                });
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { await SetStatusAsync($"加载失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
        }

        private async Task<List<ColumnInfo>> GetTableColumnsAsync(string tableName, CancellationToken token)
        {
            var columns = new List<ColumnInfo>();
            var sql = $"PRAGMA table_info(\"{tableName}\")";
            var result = await _databaseService.ExecuteQueryAsync(sql, new List<object>());
            token.ThrowIfCancellationRequested();
            if (result == null || result.Count <= 1) return columns;
            for (int i = 1; i < result.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var row = result[i];
                if (row.Count >= 6)
                    columns.Add(new ColumnInfo { Cid = Convert.ToInt32(row[0]), Name = row[1]?.ToString() ?? "", Type = row[2]?.ToString() ?? "", Pk = Convert.ToInt32(row[5]) });
            }
            return columns;
        }

        #endregion

        #region 字段编辑

        private void StartEditFields()
        {
            if (string.IsNullOrEmpty(SelectedTable)) return;
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsEditingFields = true;
                foreach (var f in EditableFields) f.IsReadOnly = false;
                SetStatus($"编辑模式：可直接修改表 '{SelectedTable}' 的字段名称", new SolidColorBrush(Colors.Orange));
            }).ConfigureAwait(false);
        }

        private async Task SaveFieldChangesAsync()
        {
            if (string.IsNullOrEmpty(SelectedTable) || !IsEditingFields) return;
            try
            {
                await SetLoadingAsync(true, "正在保存字段修改...");
                var originalFields = await GetTableColumnsAsync(SelectedTable, CancellationToken.None);
                var newNames = EditableFields.Select(f => f.FieldName).ToList();

                // 处理重名：添加序号
                var finalNames = new List<string>();
                var counts = new Dictionary<string, int>();
                foreach (var name in newNames)
                {
                    if (counts.ContainsKey(name)) { counts[name]++; finalNames.Add($"{name}_{counts[name]}"); }
                    else { counts[name] = 1; finalNames.Add(name); }
                }

                for (int i = 0; i < originalFields.Count && i < finalNames.Count; i++)
                {
                    var oldName = originalFields[i].Name;
                    var newName = finalNames[i];
                    if (oldName == newName || string.IsNullOrEmpty(newName)) continue;
                    var sql = $"ALTER TABLE \"{SelectedTable}\" RENAME COLUMN \"{oldName}\" TO \"{newName}\"";
                    await _databaseService.ExecuteNonQueryAsync(sql, new List<object>());
                }

                await RefreshTablesAsync();
                var savedTable = SelectedTable;
                await ThreadingHelper.RunOnUIThreadAsync(() => { SelectedTable = savedTable; });
                await SetStatusAsync("✅ 字段修改保存成功", new SolidColorBrush(Colors.Green));
                IsEditingFields = false;
            }
            catch (Exception ex) { await SetStatusAsync($"❌ 保存失败: {ex.Message}", new SolidColorBrush(Colors.Red)); IsEditingFields = false; }
        }

        private void CancelEditFields()
        {
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsEditingFields = false;
                foreach (var f in EditableFields) f.IsReadOnly = true;
                if (!string.IsNullOrEmpty(SelectedTable)) _ = LoadTableFieldsAsync(SelectedTable);
                SetStatus("已取消编辑", new SolidColorBrush(Color.Parse("#78909C")));
            }).ConfigureAwait(false);
        }

        private async Task CopyFieldInfoAsync()
        {
            if (string.IsNullOrEmpty(FieldInfo)) return;
            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var mainWindow = (Avalonia.Application.Current?.ApplicationLifetime as Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime)?.MainWindow;
                var topLevel = TopLevel.GetTopLevel(mainWindow);
                if (topLevel?.Clipboard != null)
                    await topLevel.Clipboard.SetTextAsync(FieldInfo);
                SetStatus("字段信息已复制", new SolidColorBrush(Colors.Green));
            });
        }

        #endregion

        #region UI辅助

        private void SetStatus(string message, IBrush color)
        {
            if (!Dispatcher.UIThread.CheckAccess()) { Dispatcher.UIThread.Post(() => SetStatus(message, color)); return; }
            StatusMessage = message; StatusColor = color; ShowStatus = !string.IsNullOrEmpty(message);
        }
        private async Task SetStatusAsync(string message, IBrush color) => await ThreadingHelper.RunOnUIThreadAsync(() => SetStatus(message, color));
        private async Task SetLoadingAsync(bool isLoading, string? message = null) => await ThreadingHelper.RunOnUIThreadAsync(() => { IsLoading = isLoading; if (message != null) SetStatus(message, new SolidColorBrush(Colors.Orange)); });
        private void CancelCurrentOperation() { try { _cts?.Cancel(); _cts?.Dispose(); _cts = null; } catch { } }

        #endregion

        #region 清理

        public void Cleanup()
        {
            if (_isCleaned) return;
            try
            {
                CancelCurrentOperation();
                _lock.Wait(TimeSpan.FromSeconds(2));
                try { _subscriptions?.Dispose(); _ = ThreadingHelper.RunOnUIThreadAsync(() => { try { TableNames?.Clear(); FilteredTableNames?.Clear(); EditableFields?.Clear(); SelectedTable = null; IsEditingFields = false; } catch { } }); _isCleaned = true; }
                finally { try { _lock.Release(); } catch { } }
            }
            catch { _isCleaned = true; }
        }

        public void Dispose()
        {
            if (_disposed) return;
            Cleanup(); _lock?.Dispose(); _cts?.Dispose();
            RefreshCommand?.Dispose(); CopyFieldInfoCommand?.Dispose(); StartEditFieldsCommand?.Dispose(); SaveFieldChangesCommand?.Dispose(); CancelEditFieldsCommand?.Dispose();
            _disposed = true; GC.SuppressFinalize(this);
        }

        #endregion

        private class ColumnInfo { public int Cid { get; set; } public string Name { get; set; } = ""; public string Type { get; set; } = ""; public int Pk { get; set; } }
    }

    public class EditableFieldInfo : ReactiveObject
    {
        private int _ordinal;
        private string _fieldName = "";
        private bool _isReadOnly = true;
        public int Ordinal { get => _ordinal; set => this.RaiseAndSetIfChanged(ref _ordinal, value); }
        public string FieldName { get => _fieldName; set => this.RaiseAndSetIfChanged(ref _fieldName, value); }
        public bool IsReadOnly { get => _isReadOnly; set => this.RaiseAndSetIfChanged(ref _isReadOnly, value); }
    }
}