using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Avalonia.Threading;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// ViewModel 基类，统一实现 ICleanupPage 和 IDisposable
/// </summary>
public abstract class BaseViewModel : ReactiveObject, ICleanupPage, IDisposable
{
    private bool _disposed = false;
    private bool _isCleaned = false;
    private readonly List<IDisposable> _disposables = new();
    private readonly object _disposablesLock = new object();
    
    private bool _isLoading = false;
    private string _statusMessage = string.Empty;

    #region 属性

    /// <summary>
    /// 是否已初始化
    /// </summary>
    protected bool IsInitialized { get; set; } = false;

    /// <summary>
    /// 是否正在加载
    /// </summary>
    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    /// <summary>
    /// 状态消息
    /// </summary>
    public string StatusMessage
    {
        get => _statusMessage;
        set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
    }

    /// <summary>
    /// 是否已清理
    /// </summary>
    public bool IsCleaned => _isCleaned;

    #endregion

    #region 资源管理

    /// <summary>
    /// 注册需要自动释放的资源（线程安全）
    /// </summary>
    protected void RegisterDisposable(IDisposable disposable)
    {
        if (disposable == null) return;
        
        lock (_disposablesLock)
        {
            if (!_disposed && !_isCleaned)
            {
                _disposables.Add(disposable);
            }
            else
            {
                // 如果已经清理，立即释放
                try
                {
                    disposable.Dispose();
                }
                catch
                {
                }
            }
        }
    }

    /// <summary>
    /// 批量注册资源
    /// </summary>
    protected void RegisterDisposables(params IDisposable[] disposables)
    {
        foreach (var disposable in disposables)
        {
            RegisterDisposable(disposable);
        }
    }

    /// <summary>
    /// 清理资源（子类重写）
    /// </summary>
    public virtual void Cleanup()
    {
        if (_isCleaned) return;

        try
        {

            // 清理所有注册的 disposable
            lock (_disposablesLock)
            {
                foreach (var disposable in _disposables)
                {
                    try
                    {
                        disposable?.Dispose();
                    }
                    catch
                    {
                    }
                }
                _disposables.Clear();
            }

            // 清空集合
            ClearCollections();

            // 清理状态
            ClearStatus();

            _isCleaned = true;
        }
        catch
        {
        }
    }

    /// <summary>
    /// 清空集合（子类重写）
    /// </summary>
    protected virtual void ClearCollections()
    {
        // 子类重写此方法清空具体的集合
    }

    /// <summary>
    /// 清空状态
    /// </summary>
    protected virtual void ClearStatus()
    {
        StatusMessage = string.Empty;
        IsLoading = false;
    }

    #endregion

    #region 状态管理

    /// <summary>
    /// 设置状态信息（同步版本 - 必须在UI线程调用）
    /// </summary>
    protected virtual void SetStatus(string message, string color = "#2E7D32")
    {
        // 确保在UI线程
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetStatus(message, color));
            return;
        }

        StatusMessage = message;
    }

    /// <summary>
    /// 设置状态信息（异步版本 - 可在任何线程调用）
    /// </summary>
    protected async Task SetStatusAsync(string message, string color = "#2E7D32")
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
        });
    }

    /// <summary>
    /// 设置加载状态
    /// </summary>
    protected void SetLoading(bool isLoading, string? message = null)
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetLoading(isLoading, message));
            return;
        }

        IsLoading = isLoading;
        if (message != null)
        {
            StatusMessage = message;
        }
    }

    /// <summary>
    /// 设置加载状态（异步版本）
    /// </summary>
    protected async Task SetLoadingAsync(bool isLoading, string? message = null)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsLoading = isLoading;
            if (message != null)
            {
                StatusMessage = message;
            }
        });
    }

    #endregion

    #region 数据加载

    /// <summary>
    /// 刷新数据（子类重写）
    /// </summary>
    public virtual async Task RefreshAsync()
    {
        try
        {
            if (IsLoading) return; // 防止重复加载

            await SetLoadingAsync(true, "正在刷新数据...");
            await LoadDataAsync();
            await SetLoadingAsync(false);
        }
        catch (Exception ex)
        {
            await SetLoadingAsync(false);
            await SetStatusAsync($"刷新失败: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// 加载数据（子类实现）
    /// </summary>
    protected abstract Task LoadDataAsync();

    /// <summary>
    /// 安全加载数据（带异常处理）
    /// </summary>
    protected async Task<bool> SafeLoadDataAsync(Func<Task> loadAction, string successMessage = "加载成功")
    {
        try
        {
            await SetLoadingAsync(true, "正在加载数据...");
            await loadAction();
            await SetLoadingAsync(false);
            await SetStatusAsync(successMessage);
            return true;
        }
        catch (Exception ex)
        {
            await SetLoadingAsync(false);
            await SetStatusAsync($"加载失败: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region 异常处理

    /// <summary>
    /// 处理异常
    /// </summary>
    protected void HandleException(Exception ex, string? customMessage = null)
    {
        var errorMsg = customMessage ?? "操作失败";
        
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetStatus($"{errorMsg}: {ex.Message}"));
        }
        else
        {
            SetStatus($"{errorMsg}: {ex.Message}");
        }
    }

    /// <summary>
    /// 处理异常（异步版本）
    /// </summary>
    protected async Task HandleExceptionAsync(Exception ex, string? customMessage = null)
    {
        var errorMsg = customMessage ?? "操作失败";
        
        await SetStatusAsync($"{errorMsg}: {ex.Message}");
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            try
            {
                if (!_isCleaned)
                {
                    Cleanup();
                }
                
                // 清理所有注册的 disposable（双重保险）
                lock (_disposablesLock)
                {
                    foreach (var disposable in _disposables)
                    {
                        try
                        {
                            disposable?.Dispose();
                        }
                        catch 
                        {
                        }
                    }
                    _disposables.Clear();
                }
            }
            catch
            {
            }
        }

        _disposed = true;
    }

    #endregion

    #region 析构函数

    ~BaseViewModel()
    {
        Dispose(false);
    }

    #endregion
}