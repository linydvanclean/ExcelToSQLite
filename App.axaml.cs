using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using ExcelToSQLite.Views;
using ExcelToSQLite.Services;
using System.Text;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite;

public partial class App : Application
{
    // 使用 Lazy<T> 实现线程安全的延迟初始化
    private static readonly Lazy<IndicatorService> _indicatorServiceLazy = 
        new(() => new IndicatorService());
    private static readonly Lazy<DataDictionaryService> _dataDictionaryServiceLazy = 
        new(() => new DataDictionaryService());
    private static readonly Lazy<UserService> _userServiceLazy = 
        new(() => new UserService());
    private static readonly Lazy<TableInitializerService> _tableInitializerLazy = 
        new(() => new TableInitializerService());
    
    // 公开的静态属性
    public static IndicatorService IndicatorService => _indicatorServiceLazy.Value;
    public static DataDictionaryService DataDictionaryService => _dataDictionaryServiceLazy.Value;
    public static UserService UserService => _userServiceLazy.Value;
    public static TableInitializerService TableInitializer => _tableInitializerLazy.Value;
    
    public override void Initialize()
    {
        try
        {
            Console.WriteLine("=== App.Initialize 开始 ===");
            
            // 注册编码
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("  ✅ CodePagesEncodingProvider 注册成功");

            // 设置控制台编码
            try { Console.OutputEncoding = Encoding.UTF8; } catch { }
            try { Console.InputEncoding = Encoding.UTF8; } catch { }

            AvaloniaXamlLoader.Load(this);
            Console.WriteLine("=== App.Initialize 完成 ===");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"=== App.Initialize 异常: {ex.Message} ===");
            Console.WriteLine($"堆栈: {ex.StackTrace}");
            throw;
        }
    }

    public override void OnFrameworkInitializationCompleted()
    {
        try
        {
            Console.WriteLine("=== OnFrameworkInitializationCompleted 开始 ===");
            
            // 1. 注册全局异常处理
            RegisterGlobalExceptionHandlers();
            
            // 2. 初始化数据库（同步等待）
            Console.WriteLine("=== 开始初始化数据库 ===");
            try
            {
                // 使用 Task.Run 避免阻塞 UI 线程
                var initTask = Task.Run(async () => 
                {
                    await InitializeDatabaseAsync();
                });
                initTask.Wait(TimeSpan.FromSeconds(30)); // 30秒超时
                Console.WriteLine("=== 数据库初始化完成 ===");
            }
            catch (AggregateException aggEx)
            {
                foreach (var ex in aggEx.InnerExceptions)
                {
                    Console.WriteLine($"数据库初始化异常: {ex.Message}");
                    Console.WriteLine($"堆栈: {ex.StackTrace}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据库初始化失败: {ex.Message}");
                Console.WriteLine($"堆栈: {ex.StackTrace}");
            }
            
            // 3. 初始化服务
            InitializeServices();

            // 4. 创建主窗口
            if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                Console.WriteLine("=== 创建登录窗口 ===");
                
                try
                {
                    var loginWindow = new LoginWindow();
                    loginWindow.SetAppIcon();
                    desktop.MainWindow = loginWindow;
                    Console.WriteLine("=== 登录窗口创建成功 ===");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"创建登录窗口失败: {ex.Message}");
                    Console.WriteLine($"堆栈: {ex.StackTrace}");
                    throw;
                }
                
                desktop.Exit += OnApplicationExit;
                desktop.ShutdownRequested += OnShutdownRequested;
            }
            else
            {
                Console.WriteLine($"警告: ApplicationLifetime 不是 IClassicDesktopStyleApplicationLifetime，类型为: {ApplicationLifetime?.GetType().Name}");
            }

            base.OnFrameworkInitializationCompleted();
            Console.WriteLine("=== OnFrameworkInitializationCompleted 完成 ===");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"=== OnFrameworkInitializationCompleted 异常 ===");
            Console.WriteLine($"异常: {ex.Message}");
            Console.WriteLine($"堆栈: {ex.StackTrace}");
            
            // 显示错误信息并等待用户按键
            Console.WriteLine("按任意键退出...");
            Console.ReadKey();
            Environment.Exit(1);
        }
    }
    
    /// <summary>
    /// 注册全局异常处理
    /// </summary>
    private void RegisterGlobalExceptionHandlers()
    {
        try
        {
            Console.WriteLine("=== 注册全局异常处理器 ===");
            
            AppDomain.CurrentDomain.UnhandledException += OnCurrentDomainUnhandledException;
            Console.WriteLine("  ✅ AppDomain.CurrentDomain.UnhandledException 已注册");
            
            TaskScheduler.UnobservedTaskException += OnTaskSchedulerUnobservedTaskException;
            Console.WriteLine("  ✅ TaskScheduler.UnobservedTaskException 已注册");
            
            try
            {
                Dispatcher.UIThread.UnhandledException += OnDispatcherUnhandledException;
                Console.WriteLine("  ✅ Dispatcher.UIThread.UnhandledException 已注册");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ⚠️ 注册 Dispatcher 异常处理器失败: {ex.Message}");
            }
            
            Console.WriteLine("=== 全局异常处理器注册完成 ===");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"=== 注册全局异常处理器失败: {ex.Message} ===");
        }
    }
    
    /// <summary>
    /// 初始化数据库（异步）
    /// </summary>
    private async Task InitializeDatabaseAsync()
    {
        try
        {
            Console.WriteLine("  📌 正在初始化数据库表...");
            // 等待数据库初始化完成
            await TableInitializer.InitializeAllTablesAsync();
            // 获取已创建的表
            var tables = await TableInitializer.GetAllTableNamesAsync();
            Console.WriteLine("  ✅ 数据库初始化完成");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ❌ 数据库初始化失败: {ex.Message}");
            Console.WriteLine($"  堆栈: {ex.StackTrace}");
            throw;
        }
    }
    
    /// <summary>
    /// 初始化所有服务
    /// </summary>
    private void InitializeServices()
    {
        try
        {
            Console.WriteLine("=== 初始化 Services ===");
            
            // 触发 Lazy 初始化
            var indicator = IndicatorService;
            var dataDict = DataDictionaryService;
            var user = UserService;
            
            Console.WriteLine($"  ✅ IndicatorService 已初始化");
            Console.WriteLine($"  ✅ DataDictionaryService 已初始化");
            Console.WriteLine($"  ✅ UserService 已初始化");
            
            Console.WriteLine("=== Services 初始化完成 ===");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"=== Services 初始化失败: {ex.Message} ===");
            throw;
        }
    }

    /// <summary>
    /// 检查异常是否为 IME/输入法相关
    /// </summary>
    private bool IsImeRelatedException(Exception? ex)
    {
        if (ex == null) return false;
    
        var message = ex.Message ?? "";
        var stackTrace = ex.StackTrace ?? "";
    
        // 检查消息
        if (message.Contains("IME", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Fcitx", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("DBus", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("InputMethod", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("输入法", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    
        // 检查堆栈
        if (stackTrace.Contains("DBusIme", StringComparison.OrdinalIgnoreCase) ||
            stackTrace.Contains("FcitxX11TextInputMethod", StringComparison.OrdinalIgnoreCase) ||
            stackTrace.Contains("Avalonia.FreeDesktop", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    
        // 检查内部异常
        if (ex.InnerException != null && IsImeRelatedException(ex.InnerException))
        {
            return true;
        }
    
        return false;
    }

    /// <summary>
    /// 检查异常是否为文件对话框相关（统信 UOS 系统 Bug）
    /// </summary>
    private bool IsFileDialogRelatedException(Exception? ex)
    {
        if (ex == null) return false;
    
        var message = ex.Message ?? "";
        var stackTrace = ex.StackTrace ?? "";
    
        // 检查消息
        if (message.Contains("Timer", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("dialog", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("QBasicTimer", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("dde-file", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("StorageProvider", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("FilePicker", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    
        // 检查堆栈
        if (stackTrace.Contains("dde-file-dialog", StringComparison.OrdinalIgnoreCase) ||
            stackTrace.Contains("dde-select-dialog", StringComparison.OrdinalIgnoreCase) ||
            stackTrace.Contains("com.deepin.filemanager", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    
        // 检查内部异常
        if (ex.InnerException != null && IsFileDialogRelatedException(ex.InnerException))
        {
            return true;
        }
    
        return false;
    }

    /// <summary>
    /// 应用程序域未处理异常
    /// </summary>
    private void OnCurrentDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        var exception = e.ExceptionObject as Exception;
        
        // ✅ 如果是 IME 相关异常，静默处理，不崩溃
        if (IsImeRelatedException(exception))
        {
            Console.WriteLine($"[App] ⚠️ IME 异常已被忽略 (AppDomain): {exception?.Message}");
            return;
        }
        
        // ✅ 如果是文件对话框相关异常，静默处理
        if (IsFileDialogRelatedException(exception))
        {
            Console.WriteLine($"[App] ⚠️ 文件对话框异常已被忽略 (AppDomain): {exception?.Message}");
            return;
        }
        
        Console.WriteLine("╔═══════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║          ❌ 应用程序域未处理异常 ❌                          ║");
        Console.WriteLine("╚═══════════════════════════════════════════════════════════════╝");
        Console.WriteLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        Console.WriteLine($"异常: {exception?.Message}");
        Console.WriteLine($"堆栈: {exception?.StackTrace}");
        Console.WriteLine($"终止: {e.IsTerminating}");
        
        if (exception?.InnerException != null)
        {
            Console.WriteLine($"内部异常: {exception.InnerException.Message}");
            Console.WriteLine($"内部堆栈: {exception.InnerException.StackTrace}");
        }
        
        Console.WriteLine("═══════════════════════════════════════════════════════════════");
        
        LogExceptionToFile(exception, "AppDomainUnhandled");
    }

    /// <summary>
    /// Task 未观察异常
    /// </summary>
    private void OnTaskSchedulerUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        var exception = e.Exception;
        
        // ✅ 如果是 IME 相关异常，静默处理，不崩溃
        if (IsImeRelatedException(exception))
        {
            Console.WriteLine($"[App] ⚠️ IME 异常已被忽略 (Task): {exception?.Message}");
            e.SetObserved();
            return;
        }
        
        // ✅ 如果是文件对话框相关异常，静默处理
        if (IsFileDialogRelatedException(exception))
        {
            Console.WriteLine($"[App] ⚠️ 文件对话框异常已被忽略 (Task): {exception?.Message}");
            e.SetObserved();
            return;
        }
        
        Console.WriteLine("╔═══════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║          ⚠️ Task 未观察异常 ⚠️                             ║");
        Console.WriteLine("╚═══════════════════════════════════════════════════════════════╝");
        Console.WriteLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        Console.WriteLine($"异常: {e.Exception.Message}");
        Console.WriteLine($"堆栈: {e.Exception.StackTrace}");
        
        foreach (var innerEx in e.Exception.InnerExceptions)
        {
            Console.WriteLine($"内部异常: {innerEx.Message}");
            Console.WriteLine($"内部堆栈: {innerEx.StackTrace}");
        }
        
        Console.WriteLine("═══════════════════════════════════════════════════════════════");
        
        LogExceptionToFile(e.Exception, "TaskUnobserved");
        e.SetObserved();
    }

    /// <summary>
    /// UI 线程未处理异常
    /// </summary>
    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        var exception = e.Exception;
        var stackTrace = exception?.StackTrace ?? "";
        var message = exception?.Message ?? "";
    
        // ✅ 检测 IME 相关异常
        if (message.Contains("IME", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("Fcitx", StringComparison.OrdinalIgnoreCase) ||
            message.Contains("InputMethod", StringComparison.OrdinalIgnoreCase) ||
            stackTrace.Contains("DBusIme", StringComparison.OrdinalIgnoreCase) ||
            stackTrace.Contains("FcitxX11TextInputMethod", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"[App] ⚠️ IME 异常已忽略: {message}");
            e.Handled = true;
            return;
        }
    
        // ✅ 检测文件对话框相关异常
        if (stackTrace.Contains("dde-select-dialog") ||
            stackTrace.Contains("dde-file-dialog") ||
            message.Contains("Timer is active"))
        {
            Console.WriteLine($"[App] ⚠️ 文件对话框异常已忽略: {message}");
            e.Handled = true;
            return;
        }
    
        // ✅ 检测死锁
        if (stackTrace.Contains("pthread_cond_timedwait") ||
            stackTrace.Contains("g_async_queue_timeout_pop"))
        {
            Console.WriteLine($"[App] ⚠️ 死锁异常已忽略: {message}");
            e.Handled = true;
            return;
        }
    
        // ✅ 检测 SkiaSharp 渲染异常
        if (stackTrace.Contains("libSkiaSharp") ||
            stackTrace.Contains("SkiaSharp"))
        {
            Console.WriteLine($"[App] ⚠️ 渲染异常已忽略: {message}");
            e.Handled = true;
            return;
        }
    
        // 其他异常正常处理
        Console.WriteLine($"❌ UI 线程异常: {message}");
        Console.WriteLine(stackTrace);
        LogExceptionToFile(exception, "DispatcherUnhandled");
        e.Handled = true;
    }

    /// <summary>
    /// 应用程序退出事件
    /// </summary>
    private void OnApplicationExit(object? sender, EventArgs e)
    {
        Console.WriteLine("╔═══════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║          👋 应用程序正常退出 👋                             ║");
        Console.WriteLine("╚═══════════════════════════════════════════════════════════════╝");
        Console.WriteLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        
        DisposeServices();
        Console.WriteLine("═══════════════════════════════════════════════════════════════");
    }
    
    private void DisposeServices()
    {
        try
        {
            Console.WriteLine("=== 释放 Services ===");
            
            if (_indicatorServiceLazy.IsValueCreated)
            {
                (_indicatorServiceLazy.Value as IDisposable)?.Dispose();
                Console.WriteLine("  ✅ IndicatorService 已释放");
            }
            
            if (_dataDictionaryServiceLazy.IsValueCreated)
            {
                (_dataDictionaryServiceLazy.Value as IDisposable)?.Dispose();
                Console.WriteLine("  ✅ DataDictionaryService 已释放");
            }
            
            if (_userServiceLazy.IsValueCreated)
            {
                (_userServiceLazy.Value as IDisposable)?.Dispose();
                Console.WriteLine("  ✅ UserService 已释放");
            }
            
            DatabaseService.Instance.Dispose();
            Console.WriteLine("  ✅ DatabaseService 已释放");
            Console.WriteLine("=== Services 释放完成 ===");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"=== 释放 Services 失败: {ex.Message} ===");
        }
    }

    private void OnShutdownRequested(object? sender, ShutdownRequestedEventArgs e)
    {
        Console.WriteLine($"📌 应用程序关闭请求: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
    }

    private void ShowErrorDialog(Exception ex)
    {
        Console.WriteLine($"显示错误对话框: {ex.Message}");
    }

    private void ShowCrashDialog(Exception? ex)
    {
        Console.WriteLine($"显示崩溃对话框: {ex?.Message}");
    }

    private void LogExceptionToFile(Exception? exception, string type)
    {
        try
        {
            var logDir = Path.Combine(AppContext.BaseDirectory, "logs");
            Directory.CreateDirectory(logDir);

            var logFile = Path.Combine(logDir, $"error_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            
            var content = $"""
                ╔═══════════════════════════════════════════════════════════════╗
                ║ 异常类型: {type}
                ║ 时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}
                ║ 异常: {exception?.Message}
                ║ 堆栈: {exception?.StackTrace}
                ║ 内部异常: {exception?.InnerException?.Message}
                ║ 内部堆栈: {exception?.InnerException?.StackTrace}
                ║ 应用程序: {System.Reflection.Assembly.GetEntryAssembly()?.GetName().Name}
                ║ 版本: {System.Reflection.Assembly.GetEntryAssembly()?.GetName().Version}
                ║ 操作系统: {Environment.OSVersion}
                ║ 64位: {Environment.Is64BitProcess}
                ╚═══════════════════════════════════════════════════════════════╝
                """;
            
            File.WriteAllText(logFile, content);
            Console.WriteLine($"日志已保存: {logFile}");
            
            // 删除超过30天的日志
            DeleteOldLogs(logDir, 30);
        }
        catch (Exception logEx)
        {
            Console.WriteLine($"写入日志文件失败: {logEx.Message}");
        }
    }

    private void DeleteOldLogs(string logDir, int daysToKeep)
    {
        try
        {
            var oldLogs = Directory.GetFiles(logDir, "error_*.log");
            foreach (var file in oldLogs)
            {
                if (DateTime.Now - File.GetCreationTime(file) > TimeSpan.FromDays(daysToKeep))
                {
                    try { File.Delete(file); } catch { }
                }
            }
        }
        catch { }
    }
}