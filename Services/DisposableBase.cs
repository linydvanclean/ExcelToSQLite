// Services/DisposableBase.cs
using System;
using System.Collections.Generic;

namespace ExcelToSQLite.Services;

/// <summary>
/// 可释放对象的基类
/// </summary>
public abstract class DisposableBase : IDisposable
{
    private bool _disposed = false;
    private readonly string _serviceName;
    private readonly List<IDisposable> _disposables = new List<IDisposable>();
    private readonly object _lock = new object();

    protected DisposableBase()
    {
        _serviceName = GetType().Name;
    }

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
            // 释放所有注册的资源
            lock (_lock)
            {
                foreach (var disposable in _disposables)
                {
                    try
                    {
                        disposable?.Dispose();
                    }
                    catch (Exception)
                    {
                        // 释放资源失败，忽略
                    }
                }
                _disposables.Clear();
            }

            // 释放其他托管资源
            DisposeManagedResources();
        }

        // 释放非托管资源
        DisposeUnmanagedResources();

        _disposed = true;
        OnDisposed();
    }

    /// <summary>
    /// 注册需要自动释放的资源
    /// </summary>
    protected void RegisterDisposable(IDisposable disposable)
    {
        if (disposable == null) return;
        
        lock (_lock)
        {
            _disposables.Add(disposable);
        }
    }

    /// <summary>
    /// 释放托管资源（子类重写）
    /// </summary>
    protected virtual void DisposeManagedResources()
    {
        // 子类可以重写此方法来释放托管资源
    }

    /// <summary>
    /// 释放非托管资源（子类重写）
    /// </summary>
    protected virtual void DisposeUnmanagedResources()
    {
        // 子类可以重写此方法来释放非托管资源
    }

    /// <summary>
    /// 对象被释放时的回调
    /// </summary>
    protected virtual void OnDisposed()
    {
        // 对象已释放
    }

    /// <summary>
    /// 检查对象是否已被释放
    /// </summary>
    protected void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(_serviceName);
        }
    }
}