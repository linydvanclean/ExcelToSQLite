using Avalonia.Controls;
using Avalonia.Media;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Threading.Tasks;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.ViewModels;

public class DataManagementViewModel : ReactiveObject, IDisposable
{
    private readonly DatabaseService _databaseService;
    private Window? _parentWindow;
    private bool _isDisposed;
    private bool _isLoading;

    private ObservableCollection<TableInfo> _tables = new();
    private ObservableCollection<TableInfo> _filteredTables = new();
    private int _totalTables;
    private int _totalRecords;
    private int _userTables;
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Color.Parse("#4CAF50"));
    private bool _showStatus = false;
    private bool _hasTables = false;
    private string _searchKeyword = string.Empty;

    private readonly HashSet<string> _systemTables = TableNames.AllowedSet;

    #region 属性

    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    public ObservableCollection<TableInfo> Tables
    {
        get => _tables;
        set => this.RaiseAndSetIfChanged(ref _tables, value);
    }

    public ObservableCollection<TableInfo> FilteredTables
    {
        get => _filteredTables;
        set => this.RaiseAndSetIfChanged(ref _filteredTables, value);
    }

    public int TotalTables
    {
        get => _totalTables;
        set => this.RaiseAndSetIfChanged(ref _totalTables, value);
    }

    public int TotalRecords
    {
        get => _totalRecords;
        set => this.RaiseAndSetIfChanged(ref _totalRecords, value);
    }

    public int UserTables
    {
        get => _userTables;
        set => this.RaiseAndSetIfChanged(ref _userTables, value);
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

    public bool HasTables
    {
        get => _hasTables;
        set => this.RaiseAndSetIfChanged(ref _hasTables, value);
    }

    /// <summary>
    /// 搜索关键字 - 用于筛选数据表
    /// </summary>
    public string SearchKeyword
    {
        get => _searchKeyword;
        set
        {
            this.RaiseAndSetIfChanged(ref _searchKeyword, value);
            ApplyFilter();
        }
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> RefreshCommand { get; }
    public ReactiveCommand<TableInfo, Unit> RenameTableCommand { get; }
    public ReactiveCommand<TableInfo, Unit> DeleteTableCommand { get; }
    public ReactiveCommand<TableInfo, Unit> ClearTableCommand { get; }
    public ReactiveCommand<Unit, Unit> DeleteAllTablesCommand { get; }
    public ReactiveCommand<Unit, Unit> ClearAllDataCommand { get; }
    public ReactiveCommand<Unit, Unit> ShowStatisticsCommand { get; }

    #endregion

    #region 公共方法

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    #endregion

    public DataManagementViewModel()
    {
        _databaseService = DatabaseService.Instance;

        RefreshCommand = ReactiveCommand.CreateFromTask(LoadDataAsync);
        RenameTableCommand = ReactiveCommand.CreateFromTask<TableInfo>(RenameTableAsync);
        DeleteTableCommand = ReactiveCommand.CreateFromTask<TableInfo>(DeleteTableAsync);
        ClearTableCommand = ReactiveCommand.CreateFromTask<TableInfo>(ClearTableAsync);
        DeleteAllTablesCommand = ReactiveCommand.CreateFromTask(DeleteAllTablesAsync,
            this.WhenAnyValue(x => x.HasTables));
        ClearAllDataCommand = ReactiveCommand.CreateFromTask(ClearAllDataAsync,
            this.WhenAnyValue(x => x.HasTables));
        ShowStatisticsCommand = ReactiveCommand.CreateFromTask(ShowStatisticsAsync);

        _ = LoadDataAsync();
    }

    #region 数据加载

    private async Task LoadDataAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
                SetStatus("正在加载数据...", new SolidColorBrush(Colors.Orange));
            });

            var tableInfos = new List<TableInfo>();
            var totalRecords = 0;
            var index = 1;

            var allTables = await _databaseService.GetAllTableNamesAsync();

            var userTables = allTables
                .Where(t => !_systemTables.Contains(t))
                .ToList();

            foreach (var tableName in userTables)
            {
                try
                {
                    var countSql = $"SELECT COUNT(*) FROM \"{tableName}\"";
                    var result = await _databaseService.ExecuteQueryAsync(countSql, new List<object>());
                    var recordCount = 0;

                    if (result != null && result.Count > 1)
                    {
                        recordCount = Convert.ToInt32(result[1][0]);
                    }

                    var tableInfo = new TableInfo
                    {
                        Index = index++,
                        TableName = tableName,
                        RecordCount = recordCount,
                        CreatedAt = DateTime.Now
                    };

                    tableInfos.Add(tableInfo);
                    totalRecords += recordCount;
                }
                catch
                {
                }
            }

            tableInfos = tableInfos.OrderBy(t => t.TableName).ToList();

            for (int i = 0; i < tableInfos.Count; i++)
            {
                tableInfos[i].Index = i + 1;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Tables = new ObservableCollection<TableInfo>(tableInfos);
                ApplyFilter();

                TotalTables = tableInfos.Count;
                TotalRecords = totalRecords;
                UserTables = tableInfos.Count;
                HasTables = tableInfos.Count > 0;

                SetStatus($"✅ 加载完成，共 {tableInfos.Count} 个用户表，{totalRecords} 条记录",
                    new SolidColorBrush(Colors.Green));
                IsLoading = false;
            });
        }
        catch (Exception ex)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                SetStatus($"❌ 加载失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                IsLoading = false;
            });
        }
    }

    /// <summary>
    /// 根据搜索关键字筛选表
    /// </summary>
    private void ApplyFilter()
    {
        if (string.IsNullOrWhiteSpace(SearchKeyword))
        {
            FilteredTables = new ObservableCollection<TableInfo>(Tables);
        }
        else
        {
            var filtered = Tables
                .Where(t => t.TableName.Contains(SearchKeyword, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // 重新编号
            for (int i = 0; i < filtered.Count; i++)
            {
                filtered[i].Index = i + 1;
            }

            FilteredTables = new ObservableCollection<TableInfo>(filtered);
        }
    }

    #endregion

    #region 表操作

    private async Task RenameTableAsync(TableInfo table)
    {
        if (table == null)
        {
            await SetStatusAsync("❌ 未选择要重命名的表", new SolidColorBrush(Colors.Red));
            return;
        }

        var newTableName = await ShowInputDialogAsync(
            $"✏️ 重命名表\n\n当前表名: {table.TableName}\n记录数: {table.RecordCount}\n\n请输入新的表名：",
            table.TableName
        );

        if (string.IsNullOrEmpty(newTableName))
        {
            await SetStatusAsync("❌ 表名不能为空", new SolidColorBrush(Colors.Red));
            return;
        }

        if (newTableName == table.TableName)
        {
            await SetStatusAsync("ℹ️ 表名未发生变化", new SolidColorBrush(Colors.Orange));
            return;
        }

        if (!IsValidTableName(newTableName, out var errorMessage))
        {
            await ShowMessageAsync($"❌ {errorMessage}");
            return;
        }

        var allTables = await _databaseService.GetAllTableNamesAsync();
        if (allTables.Any(t => t.Equals(newTableName, StringComparison.OrdinalIgnoreCase)))
        {
            await ShowMessageAsync($"❌ 表名 '{newTableName}' 已存在，请使用其他名称");
            return;
        }

        var confirm = await ShowConfirmAsync(
            $"⚠️ 确认重命名表\n\n原表名: {table.TableName}\n新表名: {newTableName}\n记录数: {table.RecordCount}\n\n确认重命名？",
            "确认重命名"
        );

        if (!confirm) return;

        try
        {
            await SetStatusAsync("正在重命名表...", new SolidColorBrush(Colors.Orange));

            var sql = $"ALTER TABLE \"{table.TableName}\" RENAME TO \"{newTableName}\"";
            await _databaseService.ExecuteNonQueryAsync(sql, new List<object>());

            await SetStatusAsync($"✅ 表已重命名: '{table.TableName}' → '{newTableName}'",
                new SolidColorBrush(Colors.Green));
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 重命名失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"重命名失败: {ex.Message}");
        }
    }

    private async Task DeleteTableAsync(TableInfo table)
    {
        if (table == null)
        {
            await SetStatusAsync("❌ 未选择要删除的表", new SolidColorBrush(Colors.Red));
            return;
        }

        var confirm = await ShowConfirmAsync(
            $"⚠️ 确认删除表\n\n表名: {table.TableName}\n记录数: {table.RecordCount}\n\n⚠️ 此操作将永久删除该表及其所有数据，不可恢复！\n\n确认删除？",
            "⚠️ 确认删除"
        );

        if (!confirm) return;

        try
        {
            await SetStatusAsync("正在删除表...", new SolidColorBrush(Colors.Orange));
            await _databaseService.DropTableAsync(table.TableName);
            await SetStatusAsync($"✅ 表 '{table.TableName}' 已删除", new SolidColorBrush(Colors.Green));
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 删除表失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"删除表失败: {ex.Message}");
        }
    }

    private async Task ClearTableAsync(TableInfo table)
    {
        if (table == null)
        {
            await SetStatusAsync("❌ 未选择要清空的表", new SolidColorBrush(Colors.Red));
            return;
        }

        var confirm = await ShowConfirmAsync(
            $"⚠️ 确认清空表记录\n\n表名: {table.TableName}\n记录数: {table.RecordCount}\n\n⚠️ 此操作将删除该表的所有数据，但保留表结构！\n\n确认清空？",
            "⚠️ 确认清空"
        );

        if (!confirm) return;

        try
        {
            await SetStatusAsync("正在清空表记录...", new SolidColorBrush(Colors.Orange));
            await _databaseService.DeleteAllDataAsync(table.TableName);
            await SetStatusAsync($"✅ 表 '{table.TableName}' 的所有记录已清空", new SolidColorBrush(Colors.Green));
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 清空记录失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"清空记录失败: {ex.Message}");
        }
    }

    private async Task DeleteAllTablesAsync()
    {
        if (FilteredTables.Count == 0)
        {
            await ShowMessageAsync("没有用户表可以删除");
            return;
        }

        var tableList = string.Join("\n", FilteredTables.Select(t => $"  - {t.TableName} ({t.RecordCount} 条记录)"));

        var confirm = await ShowConfirmAsync(
            $"⚠️⚠️⚠️ 警告：删除所有用户表 ⚠️⚠️⚠️\n\n" +
            $"即将删除以下 {FilteredTables.Count} 个表及其所有数据：\n\n" +
            $"{tableList}\n\n" +
            $"⚠️ 此操作不可恢复！\n\n" +
            $"确认删除所有表？",
            "⚠️ 确认全部删除"
        );

        if (!confirm) return;

        try
        {
            await SetStatusAsync("正在删除所有表...", new SolidColorBrush(Colors.Orange));

            var deletedCount = 0;
            var errorList = new List<string>();

            foreach (var table in FilteredTables)
            {
                try
                {
                    await _databaseService.DropTableAsync(table.TableName);
                    deletedCount++;
                }
                catch (Exception ex)
                {
                    errorList.Add($"{table.TableName}: {ex.Message}");
                }
            }

            var message = $"✅ 已删除 {deletedCount} 个表";
            if (errorList.Any())
            {
                message += $"\n⚠️ 删除失败的表:\n{string.Join("\n", errorList)}";
            }

            await SetStatusAsync(message, errorList.Any() ? new SolidColorBrush(Colors.Orange) : new SolidColorBrush(Colors.Green));
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 删除表失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"删除表失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 清除所有用户数据（清空所有用户表的数据，但保留表结构）
    /// </summary>
    private async Task ClearAllDataAsync()
    {
        if (FilteredTables.Count == 0)
        {
            await ShowMessageAsync("没有用户表可以清除数据");
            return;
        }

        var tableList = string.Join("\n", FilteredTables.Select(t => $"  - {t.TableName} ({t.RecordCount} 条记录)"));

        var confirm = await ShowConfirmAsync(
            $"⚠️⚠️ 警告：清除所有用户数据 ⚠️⚠️\n\n" +
            $"即将清空以下 {FilteredTables.Count} 个表的所有数据（保留表结构）：\n\n" +
            $"{tableList}\n\n" +
            $"⚠️ 此操作不可恢复！\n\n" +
            $"确认清除所有数据？",
            "⚠️ 确认清除所有数据"
        );

        if (!confirm) return;

        try
        {
            await SetStatusAsync("正在清除所有用户数据...", new SolidColorBrush(Colors.Orange));

            var clearedCount = 0;
            var errorList = new List<string>();

            foreach (var table in FilteredTables)
            {
                try
                {
                    await _databaseService.DeleteAllDataAsync(table.TableName);
                    clearedCount++;
                }
                catch (Exception ex)
                {
                    errorList.Add($"{table.TableName}: {ex.Message}");
                }
            }

            var message = $"✅ 已清空 {clearedCount} 个表的数据";
            if (errorList.Any())
            {
                message += $"\n⚠️ 清空失败的表:\n{string.Join("\n", errorList)}";
            }

            await SetStatusAsync(message, errorList.Any() ? new SolidColorBrush(Colors.Orange) : new SolidColorBrush(Colors.Green));
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 清除数据失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await ShowMessageAsync($"清除数据失败: {ex.Message}");
        }
    }

    #endregion

    #region 统计信息

    private async Task ShowStatisticsAsync()
    {
        var systemTableNames = string.Join("、", _systemTables.Where(t => t != "sqlite_sequence"));
        var systemTableCount = _systemTables.Count(t => t != "sqlite_sequence");

        var message = $"📊 数据库统计信息\n\n" +
                      $"📋 用户表总数: {TotalTables}\n" +
                      $"📝 总记录数: {TotalRecords}\n" +
                      $"🗄️ 系统表: {systemTableCount} ({systemTableNames})\n\n" +
                      $"📋 表详情:\n";

        foreach (var table in FilteredTables)
        {
            message += $"  - {table.TableName}: {table.RecordCount} 条记录\n";
        }

        await ShowMessageAsync(message);
    }

    #endregion

    #region 辅助方法

    private bool IsValidTableName(string tableName, out string errorMessage)
    {
        errorMessage = string.Empty;

        if (string.IsNullOrWhiteSpace(tableName))
        {
            errorMessage = "表名不能为空";
            return false;
        }

        var invalidChars = new[] { ' ', '-', '/', '\\', ':', '*', '?', '"', '<', '>', '|', '\'', ';' };
        if (tableName.IndexOfAny(invalidChars) >= 0)
        {
            errorMessage = "表名包含非法字符，请使用字母、数字和下划线";
            return false;
        }

        if (!char.IsLetter(tableName[0]))
        {
            errorMessage = "表名应以字母开头";
            return false;
        }

        if (_systemTables.Contains(tableName))
        {
            errorMessage = "表名与系统表冲突";
            return false;
        }

        return true;
    }

    private void SetStatus(string message, IBrush color)
    {
        StatusMessage = message;
        StatusColor = color;
        ShowStatus = !string.IsNullOrEmpty(message);
    }

    private async Task SetStatusAsync(string message, IBrush color)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            SetStatus(message, color);
        });
    }

    #endregion

    #region 对话框（使用标准化MessageBox）

    private Window? GetWindow()
    {
        if (_parentWindow != null) return _parentWindow;

        if (App.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
        {
            return desktop.MainWindow;
        }
        return null;
    }

    private async Task<string?> ShowInputDialogAsync(string message, string defaultText = "")
    {
        var window = GetWindow();
        if (window == null) return null;

        return await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            var tcs = new TaskCompletionSource<string?>();

            var dialog = new Window
            {
                Title = "✏️ 重命名表",
                Width = 450,
                Height = 280,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var stackPanel = new StackPanel
            {
                Margin = new Avalonia.Thickness(20),
                Spacing = 15
            };

            var textBlock = new TextBlock
            {
                Text = message,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                FontSize = 13,
                Margin = new Avalonia.Thickness(0, 0, 0, 10)
            };

            var textBox = new TextBox
            {
                Text = defaultText,
                FontSize = 13,
                Padding = new Avalonia.Thickness(10, 8)
            };
            textBox.SelectAll();
            textBox.Focus();

            var buttonPanel = new StackPanel
            {
                Orientation = Avalonia.Layout.Orientation.Horizontal,
                Spacing = 10,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right,
                Margin = new Avalonia.Thickness(0, 10, 0, 0)
            };

            var confirmButton = new Button
            {
                Content = "✅ 确认重命名",
                Width = 120,
                Background = new SolidColorBrush(Color.Parse("#2196F3")),
                Foreground = Brushes.White
            };
            confirmButton.Click += (s, e) =>
            {
                tcs.SetResult(textBox.Text);
                dialog.Close();
            };

            var cancelButton = new Button
            {
                Content = "❌ 取消",
                Width = 80
            };
            cancelButton.Click += (s, e) =>
            {
                tcs.SetResult(null);
                dialog.Close();
            };

            textBox.KeyDown += (s, e) =>
            {
                if (e.Key == Avalonia.Input.Key.Enter)
                {
                    tcs.SetResult(textBox.Text);
                    dialog.Close();
                }
            };

            buttonPanel.Children.Add(confirmButton);
            buttonPanel.Children.Add(cancelButton);

            stackPanel.Children.Add(textBlock);
            stackPanel.Children.Add(textBox);
            stackPanel.Children.Add(buttonPanel);
            dialog.Content = stackPanel;

            await dialog.ShowDialog(window);
            return await tcs.Task;
        });
    }

    private async Task ShowMessageAsync(string message)
    {
        var window = GetWindow();
        if (window != null)
        {
            await MessageBox.ShowAsync(window, message, "提示", MessageBoxButtons.OK);
        }
        else
        {
            await MessageBox.ShowAsync(message, "提示", MessageBoxButtons.OK);
        }
    }

    private async Task<bool> ShowConfirmAsync(string message, string title = "确认操作")
    {
        var window = GetWindow();
        MessageBoxResult result;
        if (window != null)
        {
            result = await MessageBox.ShowAsync(window, message, title, MessageBoxButtons.YesNo);
        }
        else
        {
            result = await MessageBox.ShowAsync(message, title, MessageBoxButtons.YesNo);
        }
        return result == MessageBoxResult.Yes;
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                RefreshCommand?.Dispose();
                RenameTableCommand?.Dispose();
                DeleteTableCommand?.Dispose();
                ClearTableCommand?.Dispose();
                DeleteAllTablesCommand?.Dispose();
                ClearAllDataCommand?.Dispose();
                ShowStatisticsCommand?.Dispose();
            }
            catch
            {
            }
            _isDisposed = true;
        }
        GC.SuppressFinalize(this);
    }

    #endregion
}