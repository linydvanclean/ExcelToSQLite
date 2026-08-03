using ReactiveUI;
using System;
using System.Reactive;
using System.Threading.Tasks;
using Avalonia.Media;
using ExcelToSQLite.Helpers;
using Avalonia.Threading;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 带编辑功能的 ViewModel 基类
/// </summary>
public abstract class BaseEditableViewModel<T> : BaseListViewModel<T> where T : class
{
    private bool _isLoading;
    private bool _isSaving;
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Colors.Green);
    private bool _showStatus = false;
    private string _errorMessage = string.Empty;
    private bool _hasError = false;

    protected BaseEditableViewModel()
    {
        // 初始化状态
        StatusMessage = "就绪";
        StatusColor = new SolidColorBrush(Colors.Green);
    }

    #region 属性

    /// <summary>
    /// 是否正在加载
    /// </summary>
    public new bool IsLoading  // ← 保持 new（基类存在）
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    /// <summary>
    /// 是否正在保存
    /// </summary>
    public bool IsSaving
    {
        get => _isSaving;
        set => this.RaiseAndSetIfChanged(ref _isSaving, value);
    }

    /// <summary>
    /// 是否正在执行任何操作（组合状态）
    /// </summary>
    public bool IsBusy => IsLoading || IsSaving;

    /// <summary>
    /// 状态消息
    /// </summary>
    public new string StatusMessage  // ← 保持 new（基类存在）
    {
        get => _statusMessage;
        set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
    }

    /// <summary>
    /// 状态颜色（使用 IBrush 类型）
    /// </summary>
    public IBrush StatusColor  // ← 移除 new（基类不存在）
    {
        get => _statusColor;
        set => this.RaiseAndSetIfChanged(ref _statusColor, value);
    }

    /// <summary>
    /// 是否显示状态
    /// </summary>
    public bool ShowStatus
    {
        get => _showStatus;
        set => this.RaiseAndSetIfChanged(ref _showStatus, value);
    }

    /// <summary>
    /// 错误消息
    /// </summary>
    public string ErrorMessage
    {
        get => _errorMessage;
        set => this.RaiseAndSetIfChanged(ref _errorMessage, value);
    }

    /// <summary>
    /// 是否有错误
    /// </summary>
    public bool HasError
    {
        get => _hasError;
        set => this.RaiseAndSetIfChanged(ref _hasError, value);
    }

    #endregion

    #region 状态管理

    /// <summary>
    /// 设置状态（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected new void SetStatus(string message, string color = "#2E7D32")  // ← 保持 new（基类存在）
    {
        // 确保在UI线程
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetStatus(message, color));
            return;
        }

        StatusMessage = message;
        StatusColor = Brush.Parse(color);
        ShowStatus = !string.IsNullOrEmpty(message);
        HasError = false;
    }

    /// <summary>
    /// 设置状态（异步版本 - 可在任何线程调用）
    /// </summary>
    protected async Task SetStatusAsync(string message, IBrush? color = null)  // ← 移除 new（基类不存在）
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            if (color != null)
            {
                StatusColor = color;
            }
            ShowStatus = !string.IsNullOrEmpty(message);
            HasError = false;
        });
    }

    /// <summary>
    /// 设置状态（异步版本 - 使用颜色字符串）
    /// </summary>
    protected new async Task SetStatusAsync(string message, string color)  // ← 保持 new（基类存在）
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = Brush.Parse(color);
            ShowStatus = !string.IsNullOrEmpty(message);
            HasError = false;
        });
    }

    /// <summary>
    /// 设置成功状态
    /// </summary>
    protected async Task SetSuccessStatusAsync(string message)
    {
        await SetStatusAsync(message, new SolidColorBrush(Colors.Green));
    }

    /// <summary>
    /// 设置警告状态
    /// </summary>
    protected async Task SetWarningStatusAsync(string message)
    {
        await SetStatusAsync(message, new SolidColorBrush(Colors.Orange));
    }

    /// <summary>
    /// 设置错误状态
    /// </summary>
    protected async Task SetErrorStatusAsync(string message)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = new SolidColorBrush(Colors.Red);
            ShowStatus = true;
            HasError = true;
            ErrorMessage = message;
        });
    }

    /// <summary>
    /// 设置错误状态（同步版本）
    /// </summary>
    protected void SetErrorStatus(string message)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = new SolidColorBrush(Colors.Red);
            ShowStatus = true;
            HasError = true;
            ErrorMessage = message;
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 清空状态
    /// </summary>
    protected new void ClearStatus()  // ← 保持 new（基类存在）
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = string.Empty;
            StatusColor = new SolidColorBrush(Colors.Green);
            ShowStatus = false;
            HasError = false;
            ErrorMessage = string.Empty;
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// 清空状态（异步版本）
    /// </summary>
    protected async Task ClearStatusAsync()
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = string.Empty;
            StatusColor = new SolidColorBrush(Colors.Green);
            ShowStatus = false;
            HasError = false;
            ErrorMessage = string.Empty;
        });
    }

    /// <summary>
    /// 设置加载状态
    /// </summary>
    protected async Task SetLoadingStateAsync(bool isLoading, string? message = null)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsLoading = isLoading;
            if (message != null)
            {
                StatusMessage = message;
                StatusColor = new SolidColorBrush(Colors.Orange);
                ShowStatus = true;
            }
        });
    }

    /// <summary>
    /// 设置保存状态
    /// </summary>
    protected async Task SetSavingStateAsync(bool isSaving, string? message = null)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsSaving = isSaving;
            if (message != null)
            {
                StatusMessage = message;
                StatusColor = new SolidColorBrush(Colors.Orange);
                ShowStatus = true;
            }
        });
    }

    #endregion

    #region 异常处理

    /// <summary>
    /// 处理异常（异步版本）
    /// </summary>
    protected new async Task HandleExceptionAsync(Exception ex, string? customMessage = null)  // ← 保持 new（基类存在）
    {
        var errorMsg = customMessage ?? "操作失败";
        
        await SetErrorStatusAsync($"{errorMsg}: {ex.Message}");
    }

    /// <summary>
    /// 处理异常（同步版本）
    /// </summary>
    protected new void HandleException(Exception ex, string? customMessage = null)  // ← 保持 new（基类存在）
    {
        var errorMsg = customMessage ?? "操作失败";
        
        SetErrorStatus($"{errorMsg}: {ex.Message}");
    }

    /// <summary>
    /// 安全的异步操作执行器
    /// </summary>
    protected async Task<bool> ExecuteSafeAsync(Func<Task> action, string successMessage = "操作成功", string? errorMessage = null)
    {
        try
        {
            await action();
            await SetSuccessStatusAsync(successMessage);
            return true;
        }
        catch (Exception ex)
        {
            await HandleExceptionAsync(ex, errorMessage);
            return false;
        }
    }

    /// <summary>
    /// 安全的异步操作执行器（带返回值）
    /// </summary>
    protected async Task<TResult?> ExecuteSafeAsync<TResult>(Func<Task<TResult>> action, string successMessage = "操作成功", string? errorMessage = null)
    {
        try
        {
            var result = await action();
            await SetSuccessStatusAsync(successMessage);
            return result;
        }
        catch (Exception ex)
        {
            await HandleExceptionAsync(ex, errorMessage);
            return default;
        }
    }

    #endregion

    #region 命令

    /// <summary>
    /// 刷新命令（子类实现）
    /// </summary>
    public ReactiveCommand<Unit, Unit>? RefreshCommand { get; protected set; }

    #endregion

    #region 资源清理

    /// <summary>
    /// 清理资源
    /// </summary>
    public new void Cleanup()  // ← 保持 new（基类存在）
    {
        try
        {
            ClearStatus();
        }
        catch
        {
        }
        // 注意：如果基类 Cleanup 不是 virtual，不能调用 base.Cleanup()
        // 如果基类是 virtual，请取消注释下面这行
        // base.Cleanup();
    }

    #endregion
}