using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reactive;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelToSQLite.Helpers;

namespace ExcelToSQLite.ViewModels;

public class AttendanceViewModel : ReactiveObject, ICleanupPage, IDisposable
{
    private readonly AttendanceParserService? _parserService;
    private readonly DatabaseService? _databaseService;
    private readonly AnalysisBatchService? _batchService;
    private readonly DataDictionaryService? _dictionaryService;
    private readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
    private readonly CompositeDisposable _subscriptions = new CompositeDisposable();
    private CancellationTokenSource? _cts;

    private Window? _parentWindow;
    private bool _disposed;
    private bool _isCleaned;
    private bool _isLoading;
    private string _filePath = string.Empty;
    private string _tableName = string.Empty;
    private ObservableCollection<AttendanceRecord> _attendanceRecords = new();
    private ObservableCollection<ObservableCollection<object>> _previewData = new();
    private string _previewInfo = string.Empty;
    private bool _canImport;
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Colors.Transparent);
    private bool _showStatus = false;
    private int _totalRecords = 0;
    private int _employeeCount = 0;
    private string _debugInfo = string.Empty;
    private bool _isParsing = false;
    private string _tableNameValidation = string.Empty;
    private IBrush _validationColor = new SolidColorBrush(Colors.Green);
    private bool _showValidation = false;
    private string _tableNameInfo = string.Empty;
    private ObservableCollection<string> _selectedFiles = new();
    private string? _lastOpenDirectory;

    private ObservableCollection<AnalysisBatch> _batches = new();
    private AnalysisBatch? _selectedBatch;
    private string _batchInfo = string.Empty;

    private ObservableCollection<DataDictionary> _dataDictionaries = new();
    private DataDictionary? _selectedDictionary;
    private string _dictionaryInfo = string.Empty;
    private bool _useDictionary = false;

    // 新增：当前解析的原始数据（用于调试）
    private string _rawDataInfo = string.Empty;
    private string _parsedTableType = string.Empty;

    // 声明为可空类型以消除 CS8618 警告
    public ReactiveCommand<Unit, Unit>? DownloadTemplateCommand { get; private set; }

    public AttendanceViewModel()
    {
        try
        {
            _parserService = new AttendanceParserService();
            _databaseService = DatabaseService.Instance;
            _batchService = new AnalysisBatchService();
            _dictionaryService = new DataDictionaryService();

            SelectFileCommand = ReactiveCommand.CreateFromTask(SelectFileAsync);
            ImportCommand = ReactiveCommand.CreateFromTask(ImportAsync, this.WhenAnyValue(x => x.CanImport));
            ClearCommand = ReactiveCommand.Create(Clear);
            RefreshBatchesCommand = ReactiveCommand.CreateFromTask(LoadBatchesAsync);
            RefreshDictionariesCommand = ReactiveCommand.CreateFromTask(LoadDictionariesAsync);
            GenerateTableNameCommand = ReactiveCommand.Create(GenerateTableName);

            // 初始化 DownloadTemplateCommand
            DownloadTemplateCommand = ReactiveCommand.CreateFromTask(DownloadTemplateAsync);
            DownloadTemplateCommand.ThrownExceptions.ObserveOn(RxApp.MainThreadScheduler).Subscribe(ex =>
            {
                SetStatus($"下载模板失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            }).DisposeWith(_subscriptions);

            SelectFileCommand.ThrownExceptions.ObserveOn(RxApp.MainThreadScheduler).Subscribe(ex =>
            {
                SetStatus($"选择文件失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            }).DisposeWith(_subscriptions);
            ImportCommand.ThrownExceptions.ObserveOn(RxApp.MainThreadScheduler).Subscribe(ex =>
            {
                SetStatus($"导入失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            }).DisposeWith(_subscriptions);

            this.WhenAnyValue(x => x.SelectedDictionary).Subscribe(_ =>
            {
                if (UseDictionary && SelectedDictionary != null && !string.IsNullOrEmpty(FilePath)) GenerateTableName();
            }).DisposeWith(_subscriptions);
            this.WhenAnyValue(x => x.UseDictionary).Subscribe(_ =>
            {
                if (UseDictionary && SelectedDictionary != null && !string.IsNullOrEmpty(FilePath)) GenerateTableName();
                else if (!UseDictionary && !string.IsNullOrEmpty(FilePath)) GenerateTableName();
            }).DisposeWith(_subscriptions);

            CanImport = false;
            _ = InitializeAsync();
        }
        catch (Exception ex)
        {
            SetStatusSafe($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    // 添加下载模板方法
    private async Task DownloadTemplateAsync()
    {
        if (_disposed || _isCleaned) return;

        try
        {
            // 获取模板文件的完整路径
            string templateFileName = "Resources/Templates/考勤打卡样表模板.xlsx";
            string templatePath = Path.Combine(AppContext.BaseDirectory, templateFileName);

            // 检查文件是否存在
            if (!File.Exists(templatePath))
            {
                // 如果输出目录不存在，尝试从项目源目录复制
                string sourcePath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Resources", "Templates", "考勤打卡样表模板.xlsx");
                if (File.Exists(sourcePath))
                {
                    // 确保目标目录存在
                    string? targetDir = Path.GetDirectoryName(templatePath);
                    if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                        Directory.CreateDirectory(targetDir);

                    // 复制文件到输出目录
                    File.Copy(sourcePath, templatePath, true);
                }
                else
                {
                    await ShowMessageBoxAsync($"❌ 未找到模板文件: 考勤打卡样表模板.xlsx\n\n请确保文件已添加到 Resources/Templates/ 目录");
                    return;
                }
            }

            // 让用户选择保存位置
            var window = GetWindow();
            if (window == null)
            {
                await ShowMessageBoxAsync("无法获取窗口句柄");
                return;
            }

            var storageProvider = window.StorageProvider;
            if (storageProvider == null)
            {
                await ShowMessageBoxAsync("无法获取文件存储服务");
                return;
            }

            // 设置保存选项
            var options = new Avalonia.Platform.Storage.FilePickerSaveOptions
            {
                Title = "保存样表模板",
                SuggestedFileName = "考勤打卡样表模板.xlsx",
                FileTypeChoices = new List<Avalonia.Platform.Storage.FilePickerFileType>
                {
                    new("Excel 文件 (*.xlsx)") { Patterns = new List<string> { "*.xlsx" } }
                },
                DefaultExtension = ".xlsx"
            };

            // 打开保存对话框
            var file = await storageProvider.SaveFilePickerAsync(options);
            if (file == null)
            {
                // 用户取消了保存
                return;
            }

            // 获取保存路径
            var savePath = file.Path?.LocalPath ?? file.Name;
            if (string.IsNullOrEmpty(savePath))
            {
                await ShowMessageBoxAsync("获取保存路径失败");
                return;
            }

            // 确保目标目录存在
            string? saveDir = Path.GetDirectoryName(savePath);
            if (!string.IsNullOrEmpty(saveDir) && !Directory.Exists(saveDir))
                Directory.CreateDirectory(saveDir);

            // 复制模板到目标位置
            File.Copy(templatePath, savePath, true);

            SetStatus($"✅ 模板已保存到: {Path.GetFileName(savePath)}", new SolidColorBrush(Colors.Green));

            // 询问是否打开所在文件夹
            var openFolder = await ShowConfirmDialogAsync($"模板已保存成功！\n\n是否打开文件所在文件夹？");
            if (openFolder)
            {
                // 打开文件夹
                System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{savePath}\"");
            }
        }
        catch (Exception ex)
        {
            await ShowMessageBoxAsync($"下载模板失败: {ex.Message}");
            SetStatus($"❌ 下载失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #region 公共方法
    public void SetParentWindow(Window window) => _parentWindow = window;
    #endregion

    #region 属性
    public string FilePath { get => _filePath; set => this.RaiseAndSetIfChanged(ref _filePath, value); }
    public string TableName { get => _tableName; set { this.RaiseAndSetIfChanged(ref _tableName, value); ValidateTableName(); } }
    public string TableNameInfo { get => _tableNameInfo; set => this.RaiseAndSetIfChanged(ref _tableNameInfo, value); }
    public ObservableCollection<AttendanceRecord> AttendanceRecords { get => _attendanceRecords; set => this.RaiseAndSetIfChanged(ref _attendanceRecords, value); }
    public ObservableCollection<ObservableCollection<object>> PreviewData { get => _previewData; set => this.RaiseAndSetIfChanged(ref _previewData, value); }
    public string PreviewInfo { get => _previewInfo; set => this.RaiseAndSetIfChanged(ref _previewInfo, value); }
    public bool CanImport { get => _canImport; set => this.RaiseAndSetIfChanged(ref _canImport, value); }
    public string StatusMessage { get => _statusMessage; set => this.RaiseAndSetIfChanged(ref _statusMessage, value); }
    public IBrush StatusColor { get => _statusColor; set => this.RaiseAndSetIfChanged(ref _statusColor, value); }
    public bool ShowStatus { get => _showStatus; set => this.RaiseAndSetIfChanged(ref _showStatus, value); }
    public int TotalRecords { get => _totalRecords; set => this.RaiseAndSetIfChanged(ref _totalRecords, value); }
    public int EmployeeCount { get => _employeeCount; set => this.RaiseAndSetIfChanged(ref _employeeCount, value); }
    public string DebugInfo { get => _debugInfo; set => this.RaiseAndSetIfChanged(ref _debugInfo, value); }
    public bool IsParsing { get => _isParsing; set => this.RaiseAndSetIfChanged(ref _isParsing, value); }
    public bool IsLoading { get => _isLoading; set => this.RaiseAndSetIfChanged(ref _isLoading, value); }
    public string TableNameValidation { get => _tableNameValidation; set => this.RaiseAndSetIfChanged(ref _tableNameValidation, value); }
    public IBrush ValidationColor { get => _validationColor; set => this.RaiseAndSetIfChanged(ref _validationColor, value); }
    public bool ShowValidation { get => _showValidation; set => this.RaiseAndSetIfChanged(ref _showValidation, value); }
    public ObservableCollection<string> SelectedFiles { get => _selectedFiles; set => this.RaiseAndSetIfChanged(ref _selectedFiles, value); }
    public ObservableCollection<AnalysisBatch> Batches { get => _batches; set => this.RaiseAndSetIfChanged(ref _batches, value); }
    public AnalysisBatch? SelectedBatch
    {
        get => _selectedBatch;
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedBatch, value);
            if (value != null) { BatchInfo = $"已选批次: {value.Name} (创建: {value.CreatedAt:yyyy-MM-dd HH:mm})"; if (!string.IsNullOrEmpty(FilePath)) GenerateTableName(); }
            else { BatchInfo = "请选择批次"; }
        }
    }
    public string BatchInfo { get => _batchInfo; set => this.RaiseAndSetIfChanged(ref _batchInfo, value); }
    public ObservableCollection<DataDictionary> DataDictionaries { get => _dataDictionaries; set => this.RaiseAndSetIfChanged(ref _dataDictionaries, value); }
    public DataDictionary? SelectedDictionary
    {
        get => _selectedDictionary;
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedDictionary, value);
            if (value != null) { DictionaryInfo = $"已选字典: {value.Name} (表名: {(string.IsNullOrEmpty(value.TableName) ? "使用文件名" : value.TableName)})"; if (UseDictionary && !string.IsNullOrEmpty(FilePath)) GenerateTableName(); }
            else { DictionaryInfo = "未选择数据字典"; }
        }
    }
    public string DictionaryInfo { get => _dictionaryInfo; set => this.RaiseAndSetIfChanged(ref _dictionaryInfo, value); }
    public bool UseDictionary { get => _useDictionary; set { this.RaiseAndSetIfChanged(ref _useDictionary, value); if (!string.IsNullOrEmpty(FilePath)) GenerateTableName(); } }
    public string RawDataInfo { get => _rawDataInfo; set => this.RaiseAndSetIfChanged(ref _rawDataInfo, value); }
    public string ParsedTableType { get => _parsedTableType; set => this.RaiseAndSetIfChanged(ref _parsedTableType, value); }
    #endregion

    #region 命令
    public ReactiveCommand<Unit, Unit> SelectFileCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> ImportCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> ClearCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> RefreshBatchesCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> RefreshDictionariesCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> GenerateTableNameCommand { get; } = null!;
    #endregion

    #region 初始化
    private async Task InitializeAsync()
    {
        if (_disposed || _isCleaned) return;
        await _lock.WaitAsync();
        try
        {
            if (_disposed || _isCleaned) return;
            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            await Task.WhenAll(LoadBatchesAsync(), LoadDictionariesAsync());
            if (!token.IsCancellationRequested && !_disposed && !_isCleaned) SetStatus("初始化完成", new SolidColorBrush(Colors.Green));
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { await SetStatusAsync($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
        finally { _lock.Release(); }
    }
    #endregion

    #region 数据加载
    private async Task LoadBatchesAsync()
    {
        if (_disposed || _isCleaned || _batchService == null) return;
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() => { if (!_disposed && !_isCleaned) IsLoading = true; });
            var list = await _batchService.GetAllAsync(1000);
            var sortedList = list.OrderByDescending(b => b.CreatedAt).ToList();
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                Batches = new ObservableCollection<AnalysisBatch>(sortedList);
                if (Batches.Count > 0) { SelectedBatch = Batches[0]; BatchInfo = $"共 {Batches.Count} 个批次，已选择最新批次"; }
                else { SelectedBatch = null; BatchInfo = "暂无批次，请先创建批次"; }
                IsLoading = false;
            });
        }
        catch (Exception ex) { await SetStatusAsync($"加载批次失败: {ex.Message}", new SolidColorBrush(Colors.Red)); await ThreadingHelper.RunOnUIThreadAsync(() => { if (!_disposed && !_isCleaned) IsLoading = false; }); }
    }

    private async Task LoadDictionariesAsync()
    {
        if (_disposed || _isCleaned || _dictionaryService == null) return;
        try
        {
            await _dictionaryService.InitializeAsync();
            var list = await _dictionaryService.GetAllAsync();
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                DataDictionaries = new ObservableCollection<DataDictionary>(list);
                if (DataDictionaries.Count > 0) { SelectedDictionary = DataDictionaries[0]; DictionaryInfo = $"已加载 {DataDictionaries.Count} 个数据字典"; }
                else { DictionaryInfo = "暂无数据字典，请先在数据字典页面创建"; SelectedDictionary = null; }
            });
        }
        catch (Exception ex) { await SetStatusAsync($"加载数据字典失败: {ex.Message}", new SolidColorBrush(Colors.Red)); }
    }
    #endregion

    #region 文件选择（支持多文件）
    private async Task SelectFileAsync()
    {
        if (_disposed || _isCleaned) return;
        CancelCurrentOperation();
        try
        {
            var window = GetWindow();
            if (window == null) { await ShowMessageBoxAsync("无法获取窗口句柄"); return; }
            var filePaths = await OpenMultipleExcelFilesAsync(window, "选择出勤记录Excel文件（可多选）");
            if (filePaths != null && filePaths.Count > 0)
            {
                if (filePaths.Count > 0) _lastOpenDirectory = Path.GetDirectoryName(filePaths[0]);
                var primaryFilePath = filePaths[0];
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (_disposed || _isCleaned) return;
                    FilePath = primaryFilePath;
                    SelectedFiles = new ObservableCollection<string>(filePaths.Select(f => Path.GetFileName(f) ?? f));
                    GenerateTableName();
                });
                await ParseAndPreviewAsync(primaryFilePath);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { await ShowMessageBoxAsync($"选择文件失败: {ex.Message}"); }
    }

    private async Task<List<string>?> OpenMultipleExcelFilesAsync(Window? parent, string title)
    {
        if (parent == null || !parent.IsVisible || parent.WindowState == WindowState.Minimized) return null;
        var storageProvider = parent.StorageProvider;
        if (storageProvider == null) return null;
        var options = new Avalonia.Platform.Storage.FilePickerOpenOptions
        {
            Title = title,
            AllowMultiple = true,
            FileTypeFilter = new List<Avalonia.Platform.Storage.FilePickerFileType>
            {
                new("Excel 文件 (*.xlsx, *.xls)") { Patterns = new List<string> { "*.xlsx", "*.xls" } },
                new("所有文件 (*.*)") { Patterns = new List<string> { "*" } }
            }
        };
        try
        {
            using var cts2 = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var files = await storageProvider.OpenFilePickerAsync(options).WaitAsync(cts2.Token);
            if (files != null && files.Count > 0) return files.Select(f => f.Path?.LocalPath ?? f.Name).ToList();
        }
        catch { }
        return null;
    }
    #endregion

    #region 解析和预览（增强版）
    private async Task ParseAndPreviewAsync(string filePath)
    {
        if (_disposed || _isCleaned || _parserService == null) return;
        CancelCurrentOperation();
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                IsParsing = true;
                IsLoading = true;
                SetStatus("正在解析出勤记录...", new SolidColorBrush(Colors.Orange));
                PreviewInfo = "⏳ 解析中，请稍候...";
                ParsedTableType = "检测中...";
                RawDataInfo = "读取文件中...";
                // 清空之前的数据
                AttendanceRecords.Clear();
                PreviewData.Clear();
                CanImport = false;
            });

            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            // 记录开始时间
            var startTime = DateTime.Now;

            // 解析数据
            var records = await _parserService.ParseAttendanceAsync(filePath);

            // 记录结束时间
            var endTime = DateTime.Now;
            var parseDuration = (endTime - startTime).TotalSeconds;

            if (token.IsCancellationRequested || _disposed || _isCleaned) return;

            // 更新调试信息
            var fileInfo = new FileInfo(filePath);
            var rawInfo = $"文件: {fileInfo.Name}, 大小: {fileInfo.Length / 1024}KB, 解析耗时: {parseDuration:F2}秒";

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed || _isCleaned) return;
                RawDataInfo = rawInfo;

                if (records == null)
                {
                    ParsedTableType = "❌ 解析返回空数据";
                    PreviewInfo = "❌ 解析失败，未获取到任何数据";
                    SetStatus("❌ 解析失败: 未获取到数据", new SolidColorBrush(Colors.Red));
                    IsParsing = false;
                    IsLoading = false;
                    return;
                }

                // 更新记录
                AttendanceRecords = new ObservableCollection<AttendanceRecord>(records);
                TotalRecords = records.Count;
                EmployeeCount = records.Select(r => r.EmployeeId).Distinct().Count();

                // 判断是否成功解析
                if (records.Count > 0)
                {
                    ParsedTableType = $"✅ 解析成功 ({records.Count} 条记录)";
                    PreviewInfo = $"✅ 解析完成！共 {records.Count} 条打卡记录，{EmployeeCount} 名员工";
                    CanImport = true;
                    SetStatus($"✅ 解析成功: {records.Count} 条记录, {EmployeeCount} 名员工", new SolidColorBrush(Colors.Green));
                }
                else
                {
                    ParsedTableType = "⚠️ 解析完成但未找到打卡记录";
                    PreviewInfo = "⚠️ 解析完成，但未找到打卡记录。请检查文件格式是否正确。";
                    CanImport = false;
                    SetStatus("⚠️ 未找到打卡记录", new SolidColorBrush(Colors.Orange));
                }

                IsParsing = false;
                IsLoading = false;
            });

            // 生成预览数据（如果不为空）
            if (records.Count > 0 && !token.IsCancellationRequested)
            {
                await GeneratePreviewDataAsync(records, token);
            }

            // 如果解析成功，自动生成表名
            if (records.Count > 0 && !token.IsCancellationRequested)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    if (!_disposed && !_isCleaned)
                    {
                        GenerateTableName();
                        ValidateTableName();
                    }
                });
            }
        }
        catch (OperationCanceledException)
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed && !_isCleaned)
                {
                    PreviewInfo = "⏹ 解析已取消";
                    IsParsing = false;
                    IsLoading = false;
                }
            });
        }
        catch (Exception ex)
        {
            var errorMsg = $"解析失败: {ex.Message}";
            if (ex.InnerException != null)
                errorMsg += $"\n内部错误: {ex.InnerException.Message}";

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed && !_isCleaned)
                {
                    PreviewInfo = $"❌ {errorMsg}";
                    ParsedTableType = "❌ 解析失败";
                    SetStatus($"❌ 解析失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                    CanImport = false;
                    IsParsing = false;
                    IsLoading = false;
                }
            });

            await ShowMessageBoxAsync($"解析失败，请检查文件格式:\n{errorMsg}");
        }
    }

    private async Task GeneratePreviewDataAsync(List<AttendanceRecord> records, CancellationToken token)
    {
        await Task.Run(async () =>
        {
            if (token.IsCancellationRequested) return;
            var previewData = new ObservableCollection<ObservableCollection<object>>();
            var header = new ObservableCollection<object> { "序号", "工号", "姓名", "部门", "打卡时间" };
            previewData.Add(header);
            int maxPreview = Math.Min(records.Count, 100);
            for (int i = 0; i < maxPreview; i++)
            {
                if (token.IsCancellationRequested) break;
                var record = records[i];
                previewData.Add(new ObservableCollection<object>
                {
                    i + 1,
                    record.EmployeeId ?? string.Empty,
                    record.EmployeeName ?? string.Empty,
                    record.Department ?? string.Empty,
                    record.CheckTime.ToString("yyyy-MM-dd HH:mm:ss")
                });
            }
            if (records.Count > 100 && !token.IsCancellationRequested)
            {
                previewData.Add(new ObservableCollection<object> { "...", $"共 {records.Count} 条记录，仅显示前100条", string.Empty, string.Empty, string.Empty });
            }
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed && !_isCleaned && !token.IsCancellationRequested)
                    PreviewData = previewData;
            });
        }, token);
    }
    #endregion

    #region 取消操作
    private void CancelCurrentOperation() { try { _cts?.Cancel(); _cts?.Dispose(); _cts = null; } catch { } }
    #endregion

    #region 表名处理
    private void GenerateTableName()
    {
        if (_disposed || _isCleaned) return;
        if (string.IsNullOrEmpty(FilePath) || SelectedBatch == null)
        {
            TableName = string.Empty;
            TableNameInfo = "请选择文件和批次后自动生成表名";
            return;
        }
        var fileName = Path.GetFileNameWithoutExtension(FilePath);
        var tablePrefix = SelectedBatch.TablePrefix;
        string generatedName;
        if (UseDictionary && SelectedDictionary != null)
            generatedName = PublicEvent.GetFormatFilename(tablePrefix, fileName, SelectedDictionary.TableName);
        else
            generatedName = PublicEvent.GetFormatFilename(tablePrefix, fileName);
        TableName = generatedName;
        ValidateTableName();
    }

    private void ValidateTableName()
    {
        if (_disposed || _isCleaned) return;
        if (string.IsNullOrEmpty(TableName))
        {
            TableNameValidation = "⚠️ 表名不能为空";
            ValidationColor = new SolidColorBrush(Colors.Orange);
            ShowValidation = true;
            return;
        }
        var invalidChars = new[] { ' ', '-', '/', '\\', ':', '*', '?', '"', '<', '>', '|', '\'', ';' };
        if (TableName.IndexOfAny(invalidChars) >= 0)
        {
            TableNameValidation = "❌ 表名包含非法字符，请使用字母、数字和下划线";
            ValidationColor = new SolidColorBrush(Colors.Red);
            ShowValidation = true;
            return;
        }
        if (!char.IsLetter(TableName[0]))
        {
            TableNameValidation = "⚠️ 表名应以字母开头";
            ValidationColor = new SolidColorBrush(Colors.Orange);
            ShowValidation = true;
            return;
        }
        if (TableName.Length > 60)
        {
            TableNameValidation = $"⚠️ 表名过长 ({TableName.Length}/60)，建议缩短";
            ValidationColor = new SolidColorBrush(Colors.Orange);
            ShowValidation = true;
            return;
        }
        TableNameValidation = "✅ 表名有效";
        ValidationColor = new SolidColorBrush(Colors.Green);
        ShowValidation = true;
    }
    #endregion

    #region 导入（增强版）
    private async Task ImportAsync()
    {
        if (_disposed || _isCleaned || IsLoading || _databaseService == null || _parserService == null) return;
        try
        {
            if (string.IsNullOrEmpty(FilePath)) { await ShowMessageBoxAsync("请先选择Excel文件"); return; }
            if (SelectedBatch == null) { await ShowMessageBoxAsync("请先选择分析批次"); return; }
            if (string.IsNullOrEmpty(TableName)) { await ShowMessageBoxAsync("请输入表名"); return; }

            // 检查是否有数据
            if (AttendanceRecords.Count == 0)
            {
                var retry = await ShowConfirmDialogAsync("⚠️ 当前没有加载数据，是否重新解析文件？");
                if (retry)
                {
                    await ParseAndPreviewAsync(FilePath);
                    if (AttendanceRecords.Count == 0)
                    {
                        await ShowMessageBoxAsync("❌ 重新解析后仍未找到数据，请检查文件格式");
                        return;
                    }
                }
                else
                {
                    return;
                }
            }

            // 收集所有要导入的文件
            var filesToImport = new List<string>();
            if (SelectedFiles.Count > 1)
            {
                var firstDir = Path.GetDirectoryName(FilePath);
                filesToImport.Add(FilePath);
                foreach (var fn in SelectedFiles.Skip(1))
                {
                    var fullPath = Path.Combine(firstDir ?? "", fn);
                    if (File.Exists(fullPath)) filesToImport.Add(fullPath);
                }
            }
            else { filesToImport.Add(FilePath); }

            var confirm = await ShowConfirmDialogAsync(
                $"📊 确认导入出勤记录\n\n" +
                $"📁 文件数: {filesToImport.Count}\n" +
                $"📋 批次: {SelectedBatch.Name}\n" +
                $"📝 表名: {TableName}\n" +
                $"📊 当前记录: {AttendanceRecords.Count} 条\n" +
                $"👥 员工数: {EmployeeCount}\n\n" +
                $"确认继续导入？");
            if (!confirm) return;

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed && !_isCleaned)
                {
                    IsLoading = true;
                    SetStatus("正在导入数据...", new SolidColorBrush(Colors.Orange));
                }
            });

            int totalImportedRows = 0;
            var targetTables = new HashSet<string>();
            var importResults = new List<string>();

            // 基准表名
            string baseTableName = TableName;
            int? baseColCount = null;
            int tableSuffix = 1;
            var columnSigMap = new Dictionary<int, string>();
            var allErrors = new List<string>();

            for (int idx = 0; idx < filesToImport.Count; idx++)
            {
                try
                {
                    var filePath = filesToImport[idx];
                    await SetStatusAsync($"正在解析: {Path.GetFileName(filePath)}...", new SolidColorBrush(Colors.Orange));

                    var records = await _parserService.ParseAttendanceAsync(filePath);

                    if (records == null || records.Count == 0)
                    {
                        importResults.Add($"⚠️ {Path.GetFileName(filePath)}: 无数据可导入");
                        continue;
                    }

                    var columns = new List<string> { "EmployeeId", "EmployeeName", "Department", "CheckTime", "DayOfMonth", "CreatedAt" };
                    int currentColCount = columns.Count;

                    // 决定目标表名
                    string targetTableName;
                    if (idx == 0)
                    {
                        targetTableName = baseTableName;
                        baseColCount = currentColCount;
                        columnSigMap[currentColCount] = baseTableName;
                    }
                    else if (currentColCount == baseColCount)
                    {
                        targetTableName = baseTableName;
                    }
                    else if (columnSigMap.TryGetValue(currentColCount, out var existing))
                    {
                        targetTableName = existing;
                    }
                    else
                    {
                        targetTableName = $"{baseTableName}_{tableSuffix}";
                        tableSuffix++;
                        columnSigMap[currentColCount] = targetTableName;
                    }

                    await _databaseService.CreateAttendanceTableAsync(targetTableName, columns);
                    var rows = records.Select(r => new List<object>
                    {
                        r.EmployeeId ?? "",
                        r.EmployeeName ?? "",
                        r.Department ?? "",
                        r.CheckTime.ToString("yyyy-MM-dd HH:mm:ss"),
                        r.DayOfMonth,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    }).ToList();

                    await _databaseService.InsertDataAsync(targetTableName, columns, rows);

                    totalImportedRows += rows.Count;
                    targetTables.Add(targetTableName);
                    importResults.Add($"✅ {Path.GetFileName(filePath)} → {targetTableName}: {rows.Count} 条, 👥 {records.Select(r2 => r2.EmployeeId).Distinct().Count()} 人");
                }
                catch (Exception ex)
                {
                    allErrors.Add($"❌ {Path.GetFileName(filesToImport[idx])}: {ex.Message}");
                }
            }

            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                if (_disposed || _isCleaned) return;

                var resultMsg = $"✅ 导入完成!\n\n" +
                               $"📁 文件数: {filesToImport.Count}\n" +
                               $"📝 表数: {targetTables.Count}\n" +
                               $"🕐 总记录: {totalImportedRows} 条\n\n";

                if (importResults.Count > 0)
                    resultMsg += $"详情:\n{string.Join("\n", importResults)}\n";

                if (allErrors.Count > 0)
                    resultMsg += $"\n⚠️ 错误:\n{string.Join("\n", allErrors)}";

                await ShowMessageBoxAsync(resultMsg);
                SetStatus($"✅ 导入成功: {totalImportedRows} 条记录 ({targetTables.Count}个表)", new SolidColorBrush(Colors.Green));
                CanImport = false;
                IsLoading = false;
            });
        }
        catch (Exception ex)
        {
            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                if (!_disposed && !_isCleaned)
                {
                    await ShowMessageBoxAsync($"导入失败: {ex.Message}\n\n{ex.StackTrace}");
                    IsLoading = false;
                }
            });
        }
    }
    #endregion

    #region UI辅助
    private void Clear()
    {
        if (_disposed || _isCleaned) return;
        _ = ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (_disposed || _isCleaned) return;
            FilePath = string.Empty;
            TableName = string.Empty;
            TableNameInfo = string.Empty;
            TableNameValidation = string.Empty;
            ShowValidation = false;
            UseDictionary = false;
            SelectedFiles.Clear();
            AttendanceRecords.Clear();
            PreviewData.Clear();
            PreviewInfo = string.Empty;
            TotalRecords = 0;
            EmployeeCount = 0;
            CanImport = false;
            IsLoading = false;
            IsParsing = false;
            ParsedTableType = string.Empty;
            RawDataInfo = string.Empty;
            if (SelectedBatch != null) BatchInfo = $"已选批次: {SelectedBatch.Name}";
            SetStatus(string.Empty, new SolidColorBrush(Colors.Transparent));
            ShowStatus = false;
        });
    }

    private void SetStatus(string message, IBrush color)
    {
        if (!Dispatcher.UIThread.CheckAccess())
        {
            Dispatcher.UIThread.Post(() => SetStatus(message, color));
            return;
        }
        StatusMessage = message;
        StatusColor = color;
        ShowStatus = !string.IsNullOrEmpty(message);
    }

    private void SetStatusSafe(string message, IBrush color) => _ = ThreadingHelper.RunOnUIThreadAsync(() => SetStatus(message, color));
    private async Task SetStatusAsync(string message, IBrush color) => await ThreadingHelper.RunOnUIThreadAsync(() => SetStatus(message, color));

    private Window? GetWindow()
    {
        if (_parentWindow != null) return _parentWindow;
        if (App.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            return desktop.MainWindow ?? desktop.Windows.FirstOrDefault();
        return null;
    }

    private async Task ShowMessageBoxAsync(string message)
    {
        if (_disposed || _isCleaned) return;
        var window = GetWindow();
        if (window != null) await MessageBox.ShowAsync(window, message, "提示", MessageBoxButtons.OK);
        else await MessageBox.ShowAsync(message, "提示", MessageBoxButtons.OK);
    }

    private async Task<bool> ShowConfirmDialogAsync(string message)
    {
        if (_disposed || _isCleaned) return false;
        var window = GetWindow();
        if (window != null)
        {
            var r = await MessageBox.ShowAsync(window, message, "确认导入", MessageBoxButtons.YesNo);
            return r == MessageBoxResult.Yes;
        }
        else
        {
            var r = await MessageBox.ShowAsync(message, "确认导入", MessageBoxButtons.YesNo);
            return r == MessageBoxResult.Yes;
        }
    }
    #endregion

    #region 清理
    public void Cleanup()
    {
        if (_isCleaned) return;
        try
        {
            CancelCurrentOperation();
            _lock.Wait(TimeSpan.FromSeconds(2));
            try
            {
                _subscriptions?.Dispose();
                _ = ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    try
                    {
                        AttendanceRecords?.Clear();
                        PreviewData?.Clear();
                        Batches?.Clear();
                        DataDictionaries?.Clear();
                        SelectedFiles?.Clear();
                        SelectedBatch = null;
                        SelectedDictionary = null;
                        StatusMessage = string.Empty;
                        ShowStatus = false;
                        FilePath = string.Empty;
                        TableName = string.Empty;
                        TableNameInfo = string.Empty;
                        TableNameValidation = string.Empty;
                        ShowValidation = false;
                        PreviewInfo = string.Empty;
                        CanImport = false;
                        UseDictionary = false;
                        IsLoading = false;
                        IsParsing = false;
                        TotalRecords = 0;
                        EmployeeCount = 0;
                        DebugInfo = string.Empty;
                        ParsedTableType = string.Empty;
                        RawDataInfo = string.Empty;
                    }
                    catch { }
                });
                _isCleaned = true;
            }
            finally { try { _lock.Release(); } catch { } }
        }
        catch { _isCleaned = true; }
    }

    public void Dispose()
    {
        if (_disposed) return;
        Cleanup();
        _lock?.Dispose();
        _cts?.Dispose();
        SelectFileCommand?.Dispose();
        ImportCommand?.Dispose();
        ClearCommand?.Dispose();
        RefreshBatchesCommand?.Dispose();
        RefreshDictionariesCommand?.Dispose();
        GenerateTableNameCommand?.Dispose();
        DownloadTemplateCommand?.Dispose();
        _disposed = true;
        GC.SuppressFinalize(this);
    }
    #endregion
}