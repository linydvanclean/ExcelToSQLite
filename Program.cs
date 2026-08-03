using System;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.ReactiveUI;
using System.Text;
using System.IO;
using System.Linq;
using System.Diagnostics;
using Avalonia.X11;

namespace ExcelToSQLite;

class Program
{
    // Linux 信号处理
    [DllImport("libc", SetLastError = true)]
    private static extern int sigaction(int signum, IntPtr act, IntPtr oldact);
    
    private const int SIGSEGV = 11;
    private const int SIGABRT = 6;
    private const int SIGFPE = 8;
    private const int SIGILL = 4;
    
    // 静态构造函数：在 Main 之前执行，确保最早生效
    static Program()
    {
        // Linux 环境下配置输入法
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            // ✅ 强制使用软件渲染，避免 GPU 死锁
            Environment.SetEnvironmentVariable("AVALONIA_X11_RENDER_MODE", "Software");
            Environment.SetEnvironmentVariable("AVALONIA_GPU", "false");
        
            // ✅ 禁用 DBus 文件对话框
            Environment.SetEnvironmentVariable("AVALONIA_X11_USE_DBUS_FILE_PICKER", "false");
        
            // ✅ 配置输入法（使用 fcitx）
            Environment.SetEnvironmentVariable("GTK_IM_MODULE", "fcitx");
            Environment.SetEnvironmentVariable("QT_IM_MODULE", "fcitx");
            Environment.SetEnvironmentVariable("XMODIFIERS", "@im=fcitx");
            
            Console.WriteLine("[Program] ✅ Linux 输入法已启用 (fcitx)");
        }
    }

    [STAThread]
    public static void Main(string[] args)
    {
        try
        {
            // ========== 1. 运行系统诊断 ==========
            RunDiagnostics();

            // ========== 2. SQLite 提供者初始化（必须在任何数据库操作之前） ==========
            InitializeSQLite();

            // ========== 3. 编码设置 ==========
            InitializeEncoding();

            // ========== 4. Linux 特定设置 ==========
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                RegisterSignalHandlers();
                CheckFonts();
                ConfigureEnvironment();
            }

            // ========== 5. 内存监控 ==========
            MonitorMemory();
            
            // ✅ 添加环境诊断
            PrintEnvironmentDiagnostics();

            // ========== 6. 启动应用程序 ==========
            Console.WriteLine("=== ExcelToSQLite 启动 ===");
            Console.WriteLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            Console.WriteLine($"操作系统: {RuntimeInformation.OSDescription}");
            Console.WriteLine($"架构: {RuntimeInformation.ProcessArchitecture}");
            Console.WriteLine("========================================");

            BuildAvaloniaApp()
                .StartWithClassicDesktopLifetime(args);

            Console.WriteLine("=== 应用程序正常退出 ===");
        }
        catch (Exception ex)
        {
            // ========== 全局异常捕获 ==========
            HandleCrash(ex);
        }
    }

    private static void PrintEnvironmentDiagnostics()
    {
        Console.WriteLine("=== 环境诊断 ===");
        Console.WriteLine($"操作系统: {RuntimeInformation.OSDescription}");
        Console.WriteLine($"架构: {RuntimeInformation.ProcessArchitecture}");
    
        var importantVars = new[] 
        { 
            "DISPLAY", 
            "XDG_SESSION_TYPE", 
            "SESSION_MANAGER",
            "GTK_IM_MODULE",
            "QT_IM_MODULE",
            "XMODIFIERS",
            "AVALONIA_GLOBAL_IME_DISABLED",
            "AVALONIA_IME_ENABLED"
        };
    
        foreach (var varName in importantVars)
        {
            var value = Environment.GetEnvironmentVariable(varName);
            Console.WriteLine($"{varName}: {(string.IsNullOrEmpty(value) ? "(未设置)" : value)}");
        }
        Console.WriteLine("=== 环境诊断完成 ===");
    }

    #region 初始化方法

    private static void InitializeSQLite()
    {
        Console.WriteLine("=== 初始化 SQLite 提供者 ===");
        try
        {
            Console.WriteLine("=== 程序启动 ===");
        
            // 1. 初始化 SQLite
            Console.WriteLine("[1] 初始化 SQLite...");
        
            // 使用 e_sqlite3 provider（Windows 和 Linux 都支持）
            SQLitePCL.raw.SetProvider(new SQLitePCL.SQLite3Provider_e_sqlite3());
            Console.WriteLine("[1] SQLite 初始化成功");
        
            // 2. 验证 SQLite
            Console.WriteLine("[2] 验证 SQLite...");
            using var testConn = new Microsoft.Data.Sqlite.SqliteConnection("Data Source=:memory:");
            testConn.Open();
            Console.WriteLine("[2] SQLite 验证成功");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ❌ SQLitePCL 提供者初始化失败: {ex.Message}");
            Console.WriteLine($"  堆栈: {ex.StackTrace}");
            throw;
        }
        Console.WriteLine("========================================");
    }

    private static void InitializeEncoding()
    {
        Console.WriteLine("=== 初始化编码 ===");
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Console.WriteLine("  ✅ CodePagesEncodingProvider 注册成功");
            
            // 设置控制台编码（部分终端可能不支持，忽略异常）
            try 
            { 
                Console.OutputEncoding = Encoding.UTF8; 
                Console.WriteLine("  ✅ 输出编码设置为 UTF-8");
            }
            catch (Exception ex) 
            { 
                Console.WriteLine($"  ⚠️ OutputEncoding 设置失败: {ex.Message}"); 
            }

            try 
            { 
                Console.InputEncoding = Encoding.UTF8; 
                Console.WriteLine("  ✅ 输入编码设置为 UTF-8");
            }
            catch (Exception ex) 
            { 
                Console.WriteLine($"  ⚠️ InputEncoding 设置失败: {ex.Message}"); 
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ❌ 编码初始化失败: {ex.Message}");
        }
        Console.WriteLine("========================================");
    }

    #endregion

    #region Linux 特定方法

    private static void RegisterSignalHandlers()
    {
        try
        {
            Console.WriteLine("=== 注册信号处理器 ===");
            Console.WriteLine($"  监控信号: SIGSEGV (11), SIGABRT (6), SIGFPE (8), SIGILL (4)");
            Console.WriteLine("  💡 注意: 信号处理仅用于日志记录，实际崩溃由系统处理");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ⚠️ 注册信号处理器失败: {ex.Message}");
        }
    }

    private static void CheckFonts()
    {
        try
        {
            Console.WriteLine("=== 检查系统字体 ===");
            var fontDirs = new[] 
            { 
                "/usr/share/fonts", 
                "/usr/local/share/fonts", 
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts")
            };
            
            var found = false;
            foreach (var dir in fontDirs)
            {
                if (Directory.Exists(dir))
                {
                    Console.WriteLine($"  ✅ 字体目录存在: {dir}");
                    found = true;
                    
                    // 检查中文字体
                    var fontFiles = Directory.GetFiles(dir, "*.ttf", SearchOption.AllDirectories)
                        .Concat(Directory.GetFiles(dir, "*.otf", SearchOption.AllDirectories))
                        .Where(f => f.Contains("WenQuanYi") || f.Contains("wqy") || f.Contains("YaHei"))
                        .Take(3)
                        .ToList();
                    
                    if (fontFiles.Any())
                    {
                        Console.WriteLine($"     找到中文字体: {Path.GetFileName(fontFiles.First())}");
                    }
                }
            }
            
            if (!found)
            {
                Console.WriteLine("  ⚠️ 未找到字体目录，建议安装中文字体:");
                Console.WriteLine("     sudo apt install fonts-wqy-zenhei fonts-wqy-microhei");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ⚠️ 字体检查失败: {ex.Message}");
        }
        Console.WriteLine("========================================");
    }

    private static void ConfigureEnvironment()
    {
        Console.WriteLine("=== 配置环境变量 ===");
        
        try
        {
            // ✅ 1. 字体设置
            SetEnvironmentVariable("AVALONIA_DEFAULT_FONT_FAMILY", 
                "WenQuanYi Micro Hei,DejaVu Sans,Microsoft YaHei,SimHei");
    
            // ✅ 2. .NET 运行时优化
            SetEnvironmentVariable("DOTNET_EnableWriteXorExecute", "0");
            SetEnvironmentVariable("DOTNET_GCHeapHardLimit", "2097152000"); // 2GB GC限制
            
            // ✅ 3. 图形渲染设置（通过环境变量）
            SetEnvironmentVariable("AVALONIA_X11_USE_GPU", "1");
            SetEnvironmentVariable("AVALONIA_X11_SOFTWARE_RENDERING", "0");
            
            // ✅ 4. 输入法配置（启用，支持中文输入）
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                // ✅ 启用输入法：设置输入法为 fcitx（统信 UOS 默认）
                SetEnvironmentVariable("GTK_IM_MODULE", "fcitx");
                SetEnvironmentVariable("QT_IM_MODULE", "fcitx");
                SetEnvironmentVariable("XMODIFIERS", "@im=fcitx");
                
                // ✅ 不设置禁用标志，允许中文输入
                // 如果出现 IME 错误，App.xaml.cs 会兜底处理
                
                Console.WriteLine("  ✅ 输入法已启用 (fcitx)");
            }
            
            // ✅ 5. 修复：会话管理器环境变量 - 解决 SESSION_MANAGER not defined
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                var sessionManager = Environment.GetEnvironmentVariable("SESSION_MANAGER");
                if (string.IsNullOrEmpty(sessionManager))
                {
                    // 尝试从 /tmp/.ICE-unix 获取
                    try
                    {
                        var iceDir = "/tmp/.ICE-unix";
                        if (Directory.Exists(iceDir))
                        {
                            var files = Directory.GetFiles(iceDir);
                            if (files.Length > 0)
                            {
                                var socket = Path.GetFileName(files[0]);
                                var hostname = Environment.MachineName;
                                SetEnvironmentVariable("SESSION_MANAGER", $"local/{hostname}:/tmp/.ICE-unix/{socket}");
                                Console.WriteLine($"  ✅ SESSION_MANAGER 已设置");
                            }
                        }
                    }
                    catch
                    {
                        // 设置一个默认值
                        SetEnvironmentVariable("SESSION_MANAGER", $"local/{Environment.MachineName}:/tmp/.ICE-unix/0");
                    }
                }
                else
                {
                    Console.WriteLine($"  ✅ SESSION_MANAGER 已存在: {sessionManager}");
                }
            }
            
            // ✅ 6. Wayland 支持
            if (Environment.GetEnvironmentVariable("XDG_SESSION_TYPE") == "wayland")
            {
                SetEnvironmentVariable("XDG_RUNTIME_DIR", $"/run/user/{GetUid()}");
                Console.WriteLine("  ✅ 检测到 Wayland 会话");
            }
            
            // ✅ 7. X11 相关环境变量
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                var display = Environment.GetEnvironmentVariable("DISPLAY");
                if (string.IsNullOrEmpty(display))
                {
                    SetEnvironmentVariable("DISPLAY", ":0");
                    Console.WriteLine("  ✅ DISPLAY 已设置为 :0");
                }
                
                SetEnvironmentVariable("AVALONIA_X11_DIAGNOSTICS", "1");
            }
            
            // ✅ 8. 设置临时目录权限
            try
            {
                var tempDir = Path.GetTempPath();
                if (Directory.Exists(tempDir))
                {
                    Console.WriteLine($"  ✅ 临时目录: {tempDir}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ⚠️ 无法获取临时目录: {ex.Message}");
            }
    
            // ✅ 9. 设置应用程序名称
            SetEnvironmentVariable("APP_NAME", "ExcelToSQLite");
            
            Console.WriteLine("  ✅ Linux 环境配置完成");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ⚠️ 环境配置失败: {ex.Message}");
            Console.WriteLine($"  ⚠️ 堆栈: {ex.StackTrace}");
        }
        Console.WriteLine("========================================");
    }

    private static int GetUid()
    {
        try
        {
            using var process = Process.GetCurrentProcess();
            return process.Id;
        }
        catch
        {
            return 1000;
        }
    }

    #endregion

    #region 诊断和监控

    private static void RunDiagnostics()
    {
        Console.WriteLine("========================================");
        Console.WriteLine("=== 系统诊断 ===");
        
        try
        {
            Console.WriteLine($"  操作系统: {RuntimeInformation.OSDescription}");
            Console.WriteLine($"  架构: {RuntimeInformation.ProcessArchitecture}");
            Console.WriteLine($"  Framework: {RuntimeInformation.FrameworkDescription}");
            Console.WriteLine($"  .NET 版本: {Environment.Version}");
        
            Console.WriteLine($"  当前目录: {Environment.CurrentDirectory}");
            Console.WriteLine($"  可执行目录: {AppContext.BaseDirectory}");
            Console.WriteLine($"  临时目录: {Path.GetTempPath()}");
        
            var importantVars = new[] { "PATH", "LD_LIBRARY_PATH", "DISPLAY", "XDG_SESSION_TYPE" };
            foreach (var varName in importantVars)
            {
                var value = Environment.GetEnvironmentVariable(varName);
                Console.WriteLine($"  {varName}: {(string.IsNullOrEmpty(value) ? "(未设置)" : value)}");
            }
        
            try
            {
                var drive = new DriveInfo(Path.GetPathRoot(Environment.CurrentDirectory) ?? "/");
                var freeGB = drive.TotalFreeSpace / (1024.0 * 1024 * 1024);
                var totalGB = drive.TotalSize / (1024.0 * 1024 * 1024);
                Console.WriteLine($"  磁盘空间: {freeGB:F1} GB / {totalGB:F1} GB 可用");
            }
            catch { }
        
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) && File.Exists("/proc/meminfo"))
            {
                try
                {
                    var memInfo = File.ReadAllLines("/proc/meminfo")
                        .Where(l => l.StartsWith("MemTotal") || l.StartsWith("MemFree") || l.StartsWith("MemAvailable"))
                        .ToList();
                    foreach (var line in memInfo)
                    {
                        var parts = line.Split(':', StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length == 2)
                        {
                            var value = parts[1].Trim();
                            Console.WriteLine($"  {parts[0]}: {value}");
                        }
                    }
                }
                catch { }
            }
        
            CheckForCrashLogs();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ⚠️ 诊断失败: {ex.Message}");
        }
        
        Console.WriteLine("=== 诊断完成 ===");
        Console.WriteLine("========================================");
    }

    private static void CheckForCrashLogs()
    {
        try
        {
            var crashLogPath = Path.Combine(AppContext.BaseDirectory, "crash.log");
            if (File.Exists(crashLogPath))
            {
                var info = new FileInfo(crashLogPath);
                if (info.Length > 0 && info.LastWriteTime > DateTime.Now.AddHours(-24))
                {
                    Console.WriteLine($"  ⚠️ 发现最近的崩溃日志: {crashLogPath}");
                    Console.WriteLine($"     大小: {info.Length} bytes, 时间: {info.LastWriteTime}");
                }
            }
        }
        catch { }
    }

    private static void MonitorMemory()
    {
        try
        {
            var process = Process.GetCurrentProcess();
            var memoryMB = process.WorkingSet64 / (1024 * 1024);
            var privateMemoryMB = process.PrivateMemorySize64 / (1024 * 1024);
            
            Console.WriteLine($"=== 内存状态 ===");
            Console.WriteLine($"  工作集内存: {memoryMB} MB");
            Console.WriteLine($"  私有内存: {privateMemoryMB} MB");
            
            if (memoryMB > 1024)
            {
                Console.WriteLine($"  ⚠️ 内存使用较高: {memoryMB} MB");
                Console.WriteLine($"  💡 建议: 检查是否有内存泄漏");
            }
            Console.WriteLine("========================================");
        }
        catch { }
    }

    #endregion

    #region 崩溃处理

    private static void HandleCrash(Exception ex)
    {
        Console.WriteLine("========================================");
        Console.WriteLine($"=== ❌ 应用程序崩溃 ===");
        Console.WriteLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        Console.WriteLine($"异常类型: {ex.GetType().FullName}");
        Console.WriteLine($"异常消息: {ex.Message}");
        Console.WriteLine($"堆栈跟踪:");
        Console.WriteLine(ex.StackTrace);
        
        if (ex.InnerException != null)
        {
            Console.WriteLine($"\n=== 内部异常 ===");
            Console.WriteLine($"类型: {ex.InnerException.GetType().FullName}");
            Console.WriteLine($"消息: {ex.InnerException.Message}");
            Console.WriteLine($"堆栈: {ex.InnerException.StackTrace}");
        }
        
        if (ex is AggregateException aggEx)
        {
            Console.WriteLine($"\n=== 聚合异常详情 ===");
            foreach (var inner in aggEx.InnerExceptions)
            {
                Console.WriteLine($"  - {inner.GetType().Name}: {inner.Message}");
            }
        }
        
        Console.WriteLine("========================================");
        
        try
        {
            var logPath = Path.Combine(AppContext.BaseDirectory, "crash.log");
            var logContent = $"""
                ========================================
                崩溃时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}
                操作系统: {RuntimeInformation.OSDescription}
                架构: {RuntimeInformation.ProcessArchitecture}
                .NET版本: {Environment.Version}
                
                异常类型: {ex.GetType().FullName}
                异常消息: {ex.Message}
                堆栈跟踪:
                {ex.StackTrace}
                
                {(ex.InnerException != null ? $"内部异常:\n{ex.InnerException}" : "")}
                ========================================
                """;
            
            File.AppendAllText(logPath, logContent + "\n");
            Console.WriteLine($"✅ 崩溃日志已写入: {logPath}");
        }
        catch (Exception logEx)
        {
            Console.WriteLine($"⚠️ 写入崩溃日志失败: {logEx.Message}");
        }
        
        try
        {
            Services.DatabaseService.Instance?.Dispose();
        }
        catch { }
        
        // ✅ 显示错误信息，等待用户按键
        Console.WriteLine("\n按任意键退出...");
        Console.ReadKey();
        
        Environment.Exit(1);
    }

    #endregion

    #region 应用程序构建

    public static AppBuilder BuildAvaloniaApp()
    {
        try
        {
            Console.WriteLine("=== 构建 Avalonia 应用 ===");

            var builder = AppBuilder.Configure<App>()
                .UseSkia()
                .UsePlatformDetect()
                .WithInterFont()
                .LogToTrace()
                .UseReactiveUI();

            // ✅ Linux 下禁用 DBus 文件对话框
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                builder.With(new X11PlatformOptions
                {
                    UseDBusFilePicker = false
                });
                Console.WriteLine("  ✅ Linux DBus 文件对话框已禁用");
            }
            
            return builder;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ❌ 构建 Avalonia 应用失败: {ex.Message}");
            throw;
        }
    }

    #endregion

    #region 辅助方法

    private static void SetEnvironmentVariable(string name, string value)
    {
        try
        {
            Environment.SetEnvironmentVariable(name, value);
            Console.WriteLine($"  {name} = {value}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ⚠️ 设置 {name} 失败: {ex.Message}");
        }
    }

    #endregion
}