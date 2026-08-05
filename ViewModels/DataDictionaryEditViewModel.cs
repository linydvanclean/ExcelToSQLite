using ReactiveUI;
using System;
using System.Reactive;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Avalonia.Controls;
using ExcelToSQLite.Models;
using ExcelToSQLite.Helpers;
using Avalonia.Media;

namespace ExcelToSQLite.ViewModels;

public class DataDictionaryEditViewModel : ReactiveObject, IDisposable
{
    private string _name = string.Empty;
    private string _tableName = string.Empty;
    private string _description = string.Empty;
    private string _dialogTitle = "创建数据字典";
    private bool _isEditing = false;
    private bool _isSaving = false;
    private string _errorMessage = string.Empty;
    private bool _hasError = false;
    private bool _showStatus = false;
    private bool _isDisposed;
    private bool _isClosed = false;
    private IBrush _statusColor = new SolidColorBrush(Color.Parse("#C62828"));

    private Window? _dialogWindow;

    public DataDictionaryEditViewModel()
    {
        SaveCommand = ReactiveCommand.CreateFromTask(SaveAsync);
        CancelCommand = ReactiveCommand.Create(Cancel);
    }

    #region 属性

    public string Name
    {
        get => _name;
        set => this.RaiseAndSetIfChanged(ref _name, value);
    }

    public string TableName
    {
        get => _tableName;
        set => this.RaiseAndSetIfChanged(ref _tableName, value);
    }

    public string Description
    {
        get => _description;
        set => this.RaiseAndSetIfChanged(ref _description, value);
    }

    public string DialogTitle
    {
        get => _dialogTitle;
        set => this.RaiseAndSetIfChanged(ref _dialogTitle, value);
    }

    public bool IsEditing
    {
        get => _isEditing;
        set => this.RaiseAndSetIfChanged(ref _isEditing, value);
    }

    public bool IsSaving
    {
        get => _isSaving;
        set => this.RaiseAndSetIfChanged(ref _isSaving, value);
    }

    public bool IsBusy => IsSaving;

    public string ErrorMessage
    {
        get => _errorMessage;
        set => this.RaiseAndSetIfChanged(ref _errorMessage, value);
    }

    public bool HasError
    {
        get => _hasError;
        set => this.RaiseAndSetIfChanged(ref _hasError, value);
    }

    public bool ShowStatus
    {
        get => _showStatus;
        set => this.RaiseAndSetIfChanged(ref _showStatus, value);
    }

    public IBrush StatusColor
    {
        get => _statusColor;
        set => this.RaiseAndSetIfChanged(ref _statusColor, value);
    }

    #endregion

    #region 公共属性

    public DataDictionary? EditingDictionary { get; private set; }
    public Func<DataDictionary, Task<bool>>? OnSave { get; set; }
    public Action<bool>? OnClose { get; set; }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> SaveCommand { get; }
    public ReactiveCommand<Unit, Unit> CancelCommand { get; }

    #endregion

    #region 公共方法

    public void SetDialogWindow(Window window)
    {
        _dialogWindow = window;
    }

    public void LoadDictionary(DataDictionary dictionary)
    {
        if (dictionary == null)
            throw new ArgumentNullException(nameof(dictionary));

        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            EditingDictionary = dictionary;
            Name = dictionary.Name ?? string.Empty;
            TableName = dictionary.TableName ?? string.Empty;
            Description = dictionary.Description ?? string.Empty;
            IsEditing = true;
            DialogTitle = "编辑数据字典";
            ClearError();
        }).ConfigureAwait(false);
    }

    public void Reset()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Name = string.Empty;
            TableName = string.Empty;
            Description = string.Empty;
            IsEditing = false;
            EditingDictionary = null;
            DialogTitle = "创建数据字典";
            ClearError();
        }).ConfigureAwait(false);
    }

    #endregion

    #region 私有方法

    private async Task SaveAsync()
    {
        if (IsSaving || _isClosed) return;

        try
        {
            // 1. 设置保存状态
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsSaving = true;
                ClearError();
                ShowStatus = true;
                StatusColor = new SolidColorBrush(Color.Parse("#1565C0")); // 蓝色表示正在处理
            });

            // 2. 验证输入
            if (!await ValidateInputAsync())
            {
                return;
            }

            // 3. 执行保存
            if (OnSave != null)
            {
                var dictionary = BuildDictionary();
                var result = await OnSave(dictionary);

                if (result)
                {
                    // 保存成功
                    _isClosed = true;
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        IsSaving = false;
                        ShowStatus = false;
                        OnClose?.Invoke(true);
                    });
                }
                else
                {
                    // 保存失败
                    await SetErrorAsync("保存失败，请重试");
                    await ResetSavingStateAsync();
                }
            }
        }
        catch (Exception ex)
        {
            await SetErrorAsync($"保存失败: {ex.Message}");
            await ResetSavingStateAsync();
        }
    }

    /// <summary>
    /// 验证输入
    /// </summary>
    private async Task<bool> ValidateInputAsync()
    {
        // 验证名称（必填）
        if (string.IsNullOrWhiteSpace(Name))
        {
            await SetErrorAsync("请输入数据字典名称");
            await ResetSavingStateAsync();
            return false;
        }

        // 表名验证：允许为空，但如果有值必须符合规范
        if (!string.IsNullOrWhiteSpace(TableName))
        {
            var trimmedTableName = TableName.Trim();

            // 检查是否包含空格（包括全角空格）
            if (trimmedTableName.Contains(' ') || trimmedTableName.Contains('\u3000'))
            {
                await SetErrorAsync("表名不能包含空格或全角空格");
                await ResetSavingStateAsync();
                return false;
            }

            // 检查是否包含其他空白字符
            if (Regex.IsMatch(trimmedTableName, @"\s"))
            {
                await SetErrorAsync("表名不能包含空白字符（如制表符、换行符等）");
                await ResetSavingStateAsync();
                return false;
            }

            // 检查是否符合命名规范：以中文、字母或下划线开头，只能包含中文、字母、数字、下划线
            if (!Regex.IsMatch(trimmedTableName, @"^[\u4e00-\u9fa5a-zA-Z_][\u4e00-\u9fa5a-zA-Z0-9_]*$"))
            {
                await SetErrorAsync("表名必须以中文、字母或下划线开头，只能包含中文、字母、数字和下划线");
                await ResetSavingStateAsync();
                return false;
            }

            // 更新为清理后的值
            TableName = trimmedTableName;
        }

        return true;
    }

    /// <summary>
    /// 构建数据字典对象
    /// </summary>
    private DataDictionary BuildDictionary()
    {
        return new DataDictionary
        {
            Id = EditingDictionary?.Id ?? Guid.NewGuid().ToString(),
            Name = Name.Trim(),
            TableName = TableName,
            Description = Description?.Trim() ?? string.Empty,
            CreatedAt = EditingDictionary?.CreatedAt ?? DateTime.Now,
            UpdatedAt = DateTime.Now,
            CreatedBy = EditingDictionary?.CreatedBy ?? "admin",
            IsActive = true
        };
    }

    /// <summary>
    /// 重置保存状态
    /// </summary>
    private async Task ResetSavingStateAsync()
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsSaving = false;
        });
    }

    private void Cancel()
    {
        if (_isClosed) return;

        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            try
            {
                _isClosed = true;
                OnClose?.Invoke(false);
            }
            catch
            {
            }
        }).ConfigureAwait(false);
    }

    private void ClearError()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            HasError = false;
            ErrorMessage = string.Empty;
            ShowStatus = false;
        }).ConfigureAwait(false);
    }

    private async Task SetErrorAsync(string message)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            HasError = true;
            ErrorMessage = message;
            ShowStatus = true;
            StatusColor = new SolidColorBrush(Color.Parse("#C62828")); // 红色表示错误
        });
    }

    private async Task SetSuccessAsync(string message)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            HasError = false;
            ErrorMessage = message;
            ShowStatus = true;
            StatusColor = new SolidColorBrush(Color.Parse("#2E7D32")); // 绿色表示成功
        });
    }

    private async Task SetInfoAsync(string message)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            HasError = false;
            ErrorMessage = message;
            ShowStatus = true;
            StatusColor = new SolidColorBrush(Color.Parse("#0D47A1")); // 蓝色表示信息
        });
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                _isClosed = true;
                SaveCommand?.Dispose();
                CancelCommand?.Dispose();
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