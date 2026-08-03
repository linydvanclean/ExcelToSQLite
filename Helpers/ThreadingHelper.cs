using System;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Threading;

namespace ExcelToSQLite.Helpers;

public static class ThreadingHelper
{
    public static void EnsureUIThread()
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            throw new InvalidOperationException("必须在UI线程执行此操作");
        }
    }
    
    // 同步 Action - 无返回值
    public static async Task RunOnUIThreadAsync(Action action)
    {
        if (Dispatcher.UIThread.CheckAccess())
        {
            action();
        }
        else
        {
            await Dispatcher.UIThread.InvokeAsync(action);
        }
    }
    
    // 异步 Func<Task> - 无返回值，支持 async lambda
    public static async Task RunOnUIThreadAsync(Func<Task> asyncAction)
    {
        if (Dispatcher.UIThread.CheckAccess())
        {
            await asyncAction();
        }
        else
        {
            await Dispatcher.UIThread.InvokeAsync(asyncAction);
        }
    }
    
    // 同步 Func<T> - 有返回值
    public static async Task<T> RunOnUIThreadAsync<T>(Func<T> action)
    {
        if (Dispatcher.UIThread.CheckAccess())
        {
            return action();
        }
        else
        {
            return await Dispatcher.UIThread.InvokeAsync(action);
        }
    }
    
    // 异步 Func<Task<T>> - 有返回值，支持 async lambda
    public static async Task<T> RunOnUIThreadAsync<T>(Func<Task<T>> asyncAction)
    {
        if (Dispatcher.UIThread.CheckAccess())
        {
            return await asyncAction();
        }
        else
        {
            return await Dispatcher.UIThread.InvokeAsync(asyncAction);
        }
    }
}