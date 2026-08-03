using System;
using System.Threading.Tasks;
using Avalonia.Threading;
using ExcelToSQLite.Views;

namespace ExcelToSQLite.Helpers
{
    public static class ErrorDialogHelper
    {
        /// <summary>
        /// 显示错误对话框（同步方式，适用于构造函数等不能使用 await 的场景）
        /// 注意：此方法会阻塞当前线程，建议在 UI 线程调用
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="title">对话框标题，默认为"错误"</param>
        /// <param name="buttons">按钮类型，默认为 OK</param>
        /// <returns>是否成功显示了对话框</returns>
        public static bool ShowErrorSync(string message, string title = "错误", MessageBoxButtons buttons = MessageBoxButtons.OK)
        {
            try
            {
                // ✅ 使用 ThreadingHelper 确保在 UI 线程执行
                // 使用 RunOnUIThreadAsync 的同步等待方式
                var task = ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    // 注意：这里不能直接 await MessageBox.ShowAsync，因为 RunOnUIThreadAsync 的 Action 重载不支持 async
                    // 需要使用 Func<Task> 重载
                    return MessageBox.ShowAsync(message, title, buttons);
                });
                
                // 等待完成（注意：如果已在 UI 线程，GetAwaiter().GetResult() 是安全的）
                task.GetAwaiter().GetResult();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 显示错误对话框（异步方式，适用于 async 方法）
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="title">对话框标题，默认为"错误"</param>
        /// <param name="buttons">按钮类型，默认为 OK</param>
        public static async Task ShowErrorAsync(string message, string title = "错误", MessageBoxButtons buttons = MessageBoxButtons.OK)
        {
            try
            {
                await MessageBox.ShowAsync(message, title, buttons);
            }
            catch
            {
            }
        }

        /// <summary>
        /// 在 UI 线程上显示错误对话框（适用于后台线程）
        /// </summary>
        /// <param name="message">错误消息</param>
        /// <param name="title">对话框标题，默认为"错误"</param>
        /// <param name="buttons">按钮类型，默认为 OK</param>
        public static async Task ShowErrorOnUIThreadAsync(string message, string title = "错误", MessageBoxButtons buttons = MessageBoxButtons.OK)
        {
            try
            {
                // ✅ 使用 ThreadingHelper 替代直接使用 Dispatcher
                await ThreadingHelper.RunOnUIThreadAsync(async () =>
                {
                    await MessageBox.ShowAsync(message, title, buttons);
                });
            }
            catch
            {
            }
        }
    }
}