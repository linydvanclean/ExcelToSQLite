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
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsSaving = true;
                ClearError();
                ShowStatus = true;
            });

            // 验证名称
            if (string.IsNullOrWhiteSpace(Name))
            {
                await SetErrorAsync("请输入数据字典名称");
                return;
            }

            // 验证表名（如果有值）
            if (!string.IsNullOrWhiteSpace(TableName) &&
                !Regex.IsMatch(TableName, @"^[a-zA-Z0-9_]+$"))
            {
                await SetErrorAsync("表名只能包含字母、数字和下划线");
                return;
            }

            if (OnSave != null)
            {
                var dictionary = new DataDictionary
                {
                    Id = EditingDictionary?.Id ?? Guid.NewGuid().ToString(),
                    Name = Name.Trim(),
                    TableName = TableName?.Trim() ?? string.Empty,
                    Description = Description?.Trim() ?? string.Empty,
                    CreatedAt = EditingDictionary?.CreatedAt ?? DateTime.Now,
                    UpdatedAt = DateTime.Now,
                    CreatedBy = EditingDictionary?.CreatedBy ?? "admin",
                    IsActive = true
                };

                var result = await OnSave(dictionary);
                
                if (result)
                {
                    _isClosed = true;
                    
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        IsSaving = false;
                        // ✅ 关闭窗口，传递 true 表示保存成功
                        OnClose?.Invoke(true);
                    });
                }
                else
                {
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        IsSaving = false;
                        HasError = true;
                        ErrorMessage = "保存失败，请重试";
                        ShowStatus = true;
                    });
                }
            }
        }
        catch (Exception ex)
        {
            await SetErrorAsync($"保存失败: {ex.Message}");
            
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsSaving = false;
            });
        }
    }

    private void Cancel()
    {
        if (_isClosed) return;
        
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            try
            {
                _isClosed = true;
                // ✅ 取消操作，传递 false
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