using System;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using Avalonia.Controls;
using System.Linq;
using System.Threading;

namespace ExcelToSQLite.Helpers;

public static class FileDialogHelper
{
    /// <summary>
    /// 保存文件（异步）
    /// </summary>
    public static async Task<string?> SaveFileAsync(Window? parent, string content, string title, string extension, string fileTypeName)
    {
        try
        {
            var filePath = await GetSaveFilePathAsync(parent, title, $"{fileTypeName.Replace("文件", "").Trim()}_{DateTime.Now:yyyyMMdd_HHmmss}", extension, fileTypeName);

            if (string.IsNullOrEmpty(filePath))
                return null;

            await File.WriteAllTextAsync(filePath, content);
            return filePath;
        }
        catch
        {
            throw;
        }
    }

    /// <summary>
    /// 打开文件（支持单个扩展名）- 返回文件路径
    /// </summary>
    public static async Task<string?> OpenFileAsync(Window? parent, string title, string extension, string fileTypeName)
    {
        return await OpenFileAsync(parent, title, new[] { extension }, fileTypeName);
    }

    /// <summary>
    /// 打开文件（支持多个扩展名）- 返回文件路径
    /// </summary>
    public static async Task<string?> OpenFileAsync(Window? parent, string title, string[] extensions, string fileTypeName)
    {
        // ✅ 检查父窗口是否有效
        if (parent == null)
        {
            return null;
        }

        // ✅ 检查父窗口是否可见
        if (!parent.IsVisible || parent.WindowState == WindowState.Minimized)
        {
            return null;
        }

        var storageProvider = parent.StorageProvider;
        if (storageProvider == null)
        {
            return null;
        }

        var patterns = new List<string>();
        foreach (var ext in extensions)
        {
            var cleanExt = ext.TrimStart('.');
            patterns.Add($"*.{cleanExt}");
        }

        var options = new Avalonia.Platform.Storage.FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = false,
            FileTypeFilter = new List<Avalonia.Platform.Storage.FilePickerFileType>
            {
                new Avalonia.Platform.Storage.FilePickerFileType($"{fileTypeName} ({string.Join(", ", patterns)})")
                {
                    Patterns = patterns
                },
                new Avalonia.Platform.Storage.FilePickerFileType("所有文件 (*.*)")
                {
                    Patterns = new List<string> { "*" }
                }
            }
        };

        try
        {
            // ✅ 使用超时机制，避免死锁
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

            var files = await storageProvider.OpenFilePickerAsync(options)
                .WaitAsync(cts.Token)
                .ConfigureAwait(false);

            if (files != null && files.Count > 0)
            {
                var file = files[0];
                var localPath = file.Path?.LocalPath ?? file.Name;
                return localPath;
            }
        }
        catch (TimeoutException)
        {
            return null;
        }
        catch (OperationCanceledException)
        {
            return null;
        }
        catch
        {
            return null;
        }

        return null;
    }

    /// <summary>
    /// 打开文件并读取内容（单个扩展名）
    /// </summary>
    public static async Task<string?> OpenFileAndReadContentAsync(Window? parent, string title, string extension, string fileTypeName)
    {
        var filePath = await OpenFileAsync(parent, title, extension, fileTypeName);
        if (string.IsNullOrEmpty(filePath))
            return null;

        try
        {
            return await File.ReadAllTextAsync(filePath);
        }
        catch
        {
            throw;
        }
    }

    /// <summary>
    /// 打开文件并读取内容（支持多个扩展名）
    /// </summary>
    public static async Task<string?> OpenFileAndReadContentAsync(Window? parent, string title, string[] extensions, string fileTypeName)
    {
        var filePath = await OpenFileAsync(parent, title, extensions, fileTypeName);
        if (string.IsNullOrEmpty(filePath))
            return null;

        try
        {
            return await File.ReadAllTextAsync(filePath);
        }
        catch
        {
            throw;
        }
    }

    /// <summary>
    /// 获取保存文件路径
    /// </summary>
    public static async Task<string?> GetSaveFilePathAsync(Window? parent, string title, string defaultFileName, string extension, string fileTypeName)
    {
        // ✅ 检查父窗口是否有效
        if (parent == null)
        {
            return null;
        }

        if (!parent.IsVisible || parent.WindowState == WindowState.Minimized)
        {
            return null;
        }

        var storageProvider = parent.StorageProvider;
        if (storageProvider == null)
        {
            return null;
        }

        extension = extension.TrimStart('.');
        var fileNameWithoutExt = Path.GetFileNameWithoutExtension(defaultFileName);

        // 如果文件名为空或只有扩展名，生成默认名称
        if (string.IsNullOrEmpty(fileNameWithoutExt))
        {
            fileNameWithoutExt = $"{fileTypeName.Replace("文件", "").Trim()}_{DateTime.Now:yyyyMMdd_HHmmss}";
        }

        var options = new Avalonia.Platform.Storage.FilePickerSaveOptions
        {
            Title = title,
            DefaultExtension = extension,
            FileTypeChoices = new List<Avalonia.Platform.Storage.FilePickerFileType>
            {
                new Avalonia.Platform.Storage.FilePickerFileType($"{fileTypeName} (*.{extension})")
                {
                    Patterns = new List<string> { $"*.{extension}" }
                },
                new Avalonia.Platform.Storage.FilePickerFileType("所有文件 (*.*)")
                {
                    Patterns = new List<string> { "*" }
                }
            },
            SuggestedFileName = $"{fileNameWithoutExt}.{extension}",
            ShowOverwritePrompt = true
        };

        try
        {
            // ✅ 使用超时机制，避免死锁
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

            var file = await storageProvider.SaveFilePickerAsync(options)
                .WaitAsync(cts.Token)
                .ConfigureAwait(false);

            var localPath = file?.Path?.LocalPath ?? file?.Name;
            return localPath;
        }
        catch (TimeoutException)
        {
            return null;
        }
        catch (OperationCanceledException)
        {
            return null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 打开指定扩展名的文件并读取内容
    /// </summary>
    public static async Task<string?> OpenFileAndReadContentWithExtensionAsync(Window? parent,
        string title,
        string extension,
        string fileTypeDescription = "")
    {
        if (string.IsNullOrEmpty(fileTypeDescription))
        {
            fileTypeDescription = $"{extension.ToUpper()} 文件 (*.{extension})";
        }

        return await OpenFileAndReadContentAsync(parent, title,
            new[] { extension },
            fileTypeDescription);
    }

    #region 专用方法

    /// <summary>
    /// 打开 Excel 文件
    /// </summary>
    public static async Task<string?> OpenExcelFileAsync(Window? parent, string title = "选择Excel文件")
    {
        return await OpenFileAsync(parent, title, new[] { "xlsx", "xls" }, "Excel 文件");
    }

    /// <summary>
    /// 打开 CSV 文件
    /// </summary>
    public static async Task<string?> OpenCsvFileAsync(Window? parent, string title = "选择CSV文件")
    {
        return await OpenFileAsync(parent, title, new[] { "csv" }, "CSV 文件");
    }

    /// <summary>
    /// 打开 JSON 文件
    /// </summary>
    public static async Task<string?> OpenJsonFileAsync(Window? parent, string title = "选择JSON文件")
    {
        return await OpenFileAsync(parent, title, new[] { "json" }, "JSON 文件");
    }

    /// <summary>
    /// 打开 .van 指标文件
    /// </summary>
    public static async Task<string?> OpenVanFileAndReadContentAsync(Window? parent,
        string title = "选择指标文件")
    {
        return await OpenFileAndReadContentWithExtensionAsync(
            parent,
            title,
            "van",
            "指标文件 (*.van)");
    }

    #endregion
}