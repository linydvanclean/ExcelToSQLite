using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ExcelToSQLite.Views;
using ExcelToSQLite.Helpers;
using Avalonia.Media;
using Avalonia;
using Avalonia.Input;
using Avalonia.Layout;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 表选择项 - 用于数据库表勾选列表
/// </summary>
public class TableSelection : ReactiveObject
{
    private string _tableName = string.Empty;
    private bool _isSelected;

    public string TableName
    {
        get => _tableName;
        set => this.RaiseAndSetIfChanged(ref _tableName, value);
    }

    public bool IsSelected
    {
        get => _isSelected;
        set => this.RaiseAndSetIfChanged(ref _isSelected, value);
    }
}

/// <summary>
/// AI寻模 - 智能SQL生成 ViewModel
/// </summary>
public class CreateAiModelViewModel : ReactiveObject, ICleanupPage, IDisposable
{
    private readonly DatabaseService _databaseService;
    private readonly IndicatorService _indicatorService;
    private readonly AppConfigService _configService;
    private readonly DeepSeekService _deepSeekService;
    private Window? _parentWindow;
    private bool _disposed;
    private bool _isCleaned;
    private bool _isInitialized;
    private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);

    #region 表单属性

    private string _indicatorName = string.Empty;
    private string _indicatorCategory = string.Empty;
    private string _indicatorDescription = string.Empty;
    private string _indicatorSqlStatement = string.Empty;
    private string _indicatorSqlDetailData = string.Empty;
    private int _selectedTabIndex = 0;
    private string _defaultCategory = "其他";
    private ObservableCollection<string> _categoryOptions = new();

    public string IndicatorName
    {
        get => _indicatorName;
        set => this.RaiseAndSetIfChanged(ref _indicatorName, value);
    }

    public string IndicatorCategory
    {
        get => _indicatorCategory;
        set => this.RaiseAndSetIfChanged(ref _indicatorCategory, value);
    }

    public string IndicatorDescription
    {
        get => _indicatorDescription;
        set => this.RaiseAndSetIfChanged(ref _indicatorDescription, value);
    }

    public string IndicatorSqlStatement
    {
        get => _indicatorSqlStatement;
        set => this.RaiseAndSetIfChanged(ref _indicatorSqlStatement, value);
    }

    public string IndicatorSqlDetailData
    {
        get => _indicatorSqlDetailData;
        set => this.RaiseAndSetIfChanged(ref _indicatorSqlDetailData, value);
    }

    public int SelectedTabIndex
    {
        get => _selectedTabIndex;
        set => this.RaiseAndSetIfChanged(ref _selectedTabIndex, value);
    }

    public ObservableCollection<string> CategoryOptions
    {
        get => _categoryOptions;
        set => this.RaiseAndSetIfChanged(ref _categoryOptions, value);
    }

    #endregion
    
    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    #region 表选择属性

    private ObservableCollection<TableSelection> _tables = new();

    public ObservableCollection<TableSelection> Tables
    {
        get => _tables;
        set => this.RaiseAndSetIfChanged(ref _tables, value);
    }

    public bool HasSelectedTables => Tables.Any(t => t.IsSelected);

    public bool IsAllTablesSelected
    {
        get => Tables.Count > 0 && Tables.All(t => t.IsSelected);
        set
        {
            if (value)
                SelectAllTables();
            else
                ClearTableSelection();
        }
    }

    #endregion

    #region 状态属性

    private string _statusMessage = "就绪";
    private IBrush _statusColor = new SolidColorBrush(Color.Parse("#78909C"));
    private bool _isAnalyzing;
    private bool _isSaving;
    private bool _showStatus;
    private ObservableCollection<string> _progressLog = new();
    private string _combinedPrompt = string.Empty;

    public string StatusMessage
    {
        get => _statusMessage;
        set => this.RaiseAndSetIfChanged(ref _statusMessage, value);
    }

    public IBrush StatusColor
    {
        get => _statusColor;
        set => this.RaiseAndSetIfChanged(ref _statusColor, value);
    }

    public bool ShowStatus
    {
        get => _showStatus;
        set => this.RaiseAndSetIfChanged(ref _showStatus, value);
    }

    public bool IsAnalyzing
    {
        get => _isAnalyzing;
        set => this.RaiseAndSetIfChanged(ref _isAnalyzing, value);
    }

    public bool IsSaving
    {
        get => _isSaving;
        set => this.RaiseAndSetIfChanged(ref _isSaving, value);
    }

    public bool IsBusy => IsAnalyzing || IsSaving;

    public ObservableCollection<string> ProgressLog
    {
        get => _progressLog;
        set => this.RaiseAndSetIfChanged(ref _progressLog, value);
    }

    public string CombinedPrompt
    {
        get => _combinedPrompt;
        set => this.RaiseAndSetIfChanged(ref _combinedPrompt, value);
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> SubmitToAiCommand { get; }
    public ReactiveCommand<Unit, Unit> SaveCommand { get; }
    public ReactiveCommand<Unit, Unit> NewIndicatorCommand { get; }
    public ReactiveCommand<Unit, Unit> PreviewCommand { get; }
    public ReactiveCommand<Unit, Unit> SelectAllTablesCommand { get; }
    public ReactiveCommand<Unit, Unit> ClearTableSelectionCommand { get; }
    public ReactiveCommand<Unit, Unit> RefreshTablesCommand { get; }
    public ReactiveCommand<Unit, Unit> CopyPromptCommand { get; }

    #endregion

    public CreateAiModelViewModel(Window? parentWindow = null)
    {
        try
        {
            
            _parentWindow = parentWindow;
            _databaseService = DatabaseService.Instance;
            _indicatorService = new IndicatorService();
            _configService = new AppConfigService();
            _deepSeekService = new DeepSeekService();

            LoadDefaultCategories();
            _defaultCategory = CategoryOptions.FirstOrDefault() ?? "其他";
            IndicatorCategory = _defaultCategory;

            // 创建命令时添加异常处理
            SubmitToAiCommand = ReactiveCommand.CreateFromTask(
                SubmitToAiAsync,
                outputScheduler: RxApp.MainThreadScheduler);
            
            SaveCommand = ReactiveCommand.CreateFromTask(
                SaveAsync,
                outputScheduler: RxApp.MainThreadScheduler);
            
            NewIndicatorCommand = ReactiveCommand.Create(
                NewIndicator,
                outputScheduler: RxApp.MainThreadScheduler);
            
            PreviewCommand = ReactiveCommand.CreateFromTask(
                PreviewDataAsync,
                outputScheduler: RxApp.MainThreadScheduler);
            
            SelectAllTablesCommand = ReactiveCommand.Create(
                SelectAllTables,
                outputScheduler: RxApp.MainThreadScheduler);
            
            ClearTableSelectionCommand = ReactiveCommand.Create(
                ClearTableSelection,
                outputScheduler: RxApp.MainThreadScheduler);
            
            RefreshTablesCommand = ReactiveCommand.CreateFromTask(
                RefreshTablesAsync,
                outputScheduler: RxApp.MainThreadScheduler);
            
            CopyPromptCommand = ReactiveCommand.CreateFromTask(
                CopyPromptAsync,
                outputScheduler: RxApp.MainThreadScheduler);

            // 添加全局异常处理
            SubmitToAiCommand.ThrownExceptions.Subscribe(ex => 
                HandleCommandException("提交AI", ex));
            SaveCommand.ThrownExceptions.Subscribe(ex => 
                HandleCommandException("保存", ex));
            PreviewCommand.ThrownExceptions.Subscribe(ex => 
                HandleCommandException("预览", ex));
            RefreshTablesCommand.ThrownExceptions.Subscribe(ex => 
                HandleCommandException("刷新表", ex));
            CopyPromptCommand.ThrownExceptions.Subscribe(ex => 
                HandleCommandException("复制提示词", ex));

            _ = InitializeAsync();
            
        }
        catch
        {
            throw;
        }
    }

    #region 异常处理

    private void HandleCommandException(string commandName, Exception ex)
    {
        _ = SetStatusAsync($"❌ {commandName}失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        _ = AddProgressAsync($"❌ {commandName}异常: {ex.Message}");
    }

    #endregion

    #region 异步初始化

    private async Task InitializeAsync()
    {
        if (_isInitialized) return;
        
        await _initLock.WaitAsync();
        try
        {
            if (_isInitialized) return;
            
            
            await LoadCategoriesFromConfigAsync();
            await RefreshTablesAsync();
            
            _isInitialized = true;
            
        }
        catch (Exception ex)
        {
            await SetStatusAsync("初始化失败: " + ex.Message, new SolidColorBrush(Colors.Orange));
        }
        finally
        {
            _initLock.Release();
        }
    }

    #endregion

    #region 分类加载

    private void LoadDefaultCategories()
    {
        try
        {
            CategoryOptions.Clear();
            var defaultCategories = _configService.GetDefaultCategories();
            foreach (var category in defaultCategories)
            {
                CategoryOptions.Add(category);
            }
        }
        catch
        {
        }
    }

    private async Task LoadCategoriesFromConfigAsync()
    {
        try
        {
            var categories = await _configService.GetCategoriesAsync();
            if (categories != null && categories.Count > 0)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    CategoryOptions.Clear();
                    foreach (var category in categories)
                    {
                        CategoryOptions.Add(category);
                    }
                    _defaultCategory = CategoryOptions.FirstOrDefault() ?? "其他";
                    if (!CategoryOptions.Contains(IndicatorCategory))
                    {
                        IndicatorCategory = _defaultCategory;
                    }
                });
            }
        }
        catch
        {
        }
    }

    #endregion

    #region 表选择

    public async Task RefreshTablesAsync()
    {
        try
        {
            if (_disposed) return;
            
            await AddProgressAsync("⏳ 正在加载数据库表列表...");
            
            var tableNames = await _databaseService.GetAllTableNamesAsync();

            var userTables = tableNames
                .Where(t => !t.StartsWith("sqlite_", StringComparison.OrdinalIgnoreCase))
                .Where(t => !Models.TableNames.AllowedSet.Contains(t))
                .OrderBy(t => t)
                .ToList();

            var selections = userTables.Select(t => new TableSelection
            {
                TableName = t,
                IsSelected = false
            }).ToList();

            if (_disposed) return;
            
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Tables = new ObservableCollection<TableSelection>(selections);
            });
            
            await AddProgressAsync($"✅ 加载完成，共 {Tables.Count} 个表");
        }
        catch (Exception ex)
        {
            await AddProgressAsync($"❌ 加载表列表失败: {ex.Message}");
            await SetStatusAsync($"加载表列表失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private void SelectAllTables()
    {
        try
        {
            if (_disposed) return;
            
            foreach (var t in Tables) 
                t.IsSelected = true;
            
            this.RaisePropertyChanged(nameof(HasSelectedTables));
            this.RaisePropertyChanged(nameof(IsAllTablesSelected));
        }
        catch
        {
        }
    }

    private void ClearTableSelection()
    {
        try
        {
            if (_disposed) return;
            
            foreach (var t in Tables) 
                t.IsSelected = false;
            
            this.RaisePropertyChanged(nameof(HasSelectedTables));
            this.RaisePropertyChanged(nameof(IsAllTablesSelected));
        }
        catch
        {
        }
    }

    #endregion

    #region 生成合成提示词

    private async Task CopyPromptAsync()
    {
        try
        {
            if (_disposed) return;
            
            if (string.IsNullOrWhiteSpace(IndicatorDescription))
            {
                await SetStatusAsync("请先在【指标描述】标签页中输入需求描述", new SolidColorBrush(Colors.Orange));
                return;
            }

            var selectedTables = Tables.Where(t => t.IsSelected).ToList();
            if (selectedTables.Count == 0)
            {
                await SetStatusAsync("请至少勾选一个数据库表", new SolidColorBrush(Colors.Orange));
                return;
            }

            // 1. 获取表结构和字段分析
            await AddProgressAsync("⏳ 正在分析表结构...");
            var schemaInfo = await BuildSchemaInfoAsync(selectedTables);
            var fieldAnalysis = await AnalyzeTableFieldsAsync(selectedTables);
            
            await AddProgressAsync($"✅ 表结构分析完成，共识别 {fieldAnalysis.TotalFields} 个字段");
            await AddProgressAsync($"   📊 数值字段: {fieldAnalysis.NumericFields} 个");
            await AddProgressAsync($"   💰 金额字段: {fieldAnalysis.AmountFields} 个");
            await AddProgressAsync($"   📅 日期字段: {fieldAnalysis.DateFields} 个");

            // 2. 构建提示词（仅用于网页版提问）
            var prompt = BuildPromptForWeb(IndicatorDescription, schemaInfo, selectedTables, fieldAnalysis);

            // 3. 保存提示词
            CombinedPrompt = prompt;

            await SetStatusAsync("✅ 提示词已生成，可复制到DeepSeek网页端使用", new SolidColorBrush(Colors.Green));
            
            // 4. 显示提示词对话框
            await ShowPromptDialogAsync(prompt);
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"生成提示词失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await AddProgressAsync($"❌ 错误: {ex.Message}");
        }
    }

    /// <summary>
    /// 构建用于网页版提问的提示词（包含角色设定）
    /// </summary>
    private string BuildPromptForWeb(
        string description, 
        string schemaInfo, 
        List<TableSelection> selectedTables,
        FieldAnalysis? fieldAnalysis = null)
    {
        var sb = new StringBuilder();
        
        // 1. 角色设定（重要！）
        sb.AppendLine("你是一个专业的SQLite数据库专家，拥有丰富的SQL编写经验。");
        sb.AppendLine("你擅长根据业务需求生成高质量、高性能的SQL查询语句。");
        sb.AppendLine();
        
        // 2. 任务说明
        sb.AppendLine("请根据以下业务需求和表结构，生成符合SQLite规范的SQL查询语句。");
        sb.AppendLine();
        
        // 3. 业务需求
        sb.AppendLine("【业务需求】");
        sb.AppendLine(description);
        sb.AppendLine();
        
        // 4. 表结构信息
        sb.AppendLine("【数据库表结构】");
        sb.AppendLine(schemaInfo);
        sb.AppendLine();
        
        // 5. 字段分析
        if (fieldAnalysis != null && fieldAnalysis.TotalFields > 0)
        {
            sb.AppendLine("【字段分析】");
            if (fieldAnalysis.AmountFieldNames.Count > 0)
            {
                sb.AppendLine($"- 金额字段：{string.Join("、", fieldAnalysis.AmountFieldNames)}（需去除千分位分隔符）");
            }
            if (fieldAnalysis.DateFieldNames.Count > 0)
            {
                sb.AppendLine($"- 日期字段：{string.Join("、", fieldAnalysis.DateFieldNames)}（需统一日期格式）");
            }
            if (fieldAnalysis.NumericFieldNames.Count > 0)
            {
                sb.AppendLine($"- 数值字段：{string.Join("、", fieldAnalysis.NumericFieldNames)}");
            }
            sb.AppendLine();
        }
        
        // 6. SQL生成要求
        sb.AppendLine("【SQL生成要求】");
        sb.AppendLine("请生成两个SQL查询语句：");
        sb.AppendLine();
        sb.AppendLine("1. **统计SQL**：使用聚合函数（COUNT、SUM、AVG、MAX、MIN等）进行数据汇总统计");
        sb.AppendLine("2. **明细SQL**：返回完整的明细数据，包含必要的筛选和排序条件");
        sb.AppendLine();
        sb.AppendLine("两个SQL必须逻辑一致，统计结果应与明细数据相匹配。");
        sb.AppendLine();
        
        // 7. 数据格式处理要求
        sb.AppendLine("【数据格式处理要求】");
        sb.AppendLine("1. **金额字段处理**：");
        sb.AppendLine("   - 去除千分位分隔符（如 8,888,888.88 → 8888888.88）");
        sb.AppendLine("   - 使用 ROUND(字段, 2) 保留两位小数");
        sb.AppendLine("   - 使用 CAST 或 REPLACE 函数处理文本金额");
        sb.AppendLine("2. **日期字段处理**：");
        sb.AppendLine("   - 统一转换为 'YYYY-MM-DD' 格式");
        sb.AppendLine("   - 使用 date() 或 datetime() 函数");
        sb.AppendLine("   - 处理空值和无效日期（使用 COALESCE）");
        sb.AppendLine("3. **NULL值处理**：");
        sb.AppendLine("   - 使用 COALESCE 或 IFNULL 函数处理NULL值");
        sb.AppendLine("4. **表名和字段名**：");
        sb.AppendLine("   - 使用双引号包裹（如 \"table_name\"）");
        sb.AppendLine("   - 避免与SQL关键字冲突");
        sb.AppendLine();
        
        // 8. SQLite语法规范
        sb.AppendLine("【SQLite语法规范】");
        sb.AppendLine("1. 所有SQL必须严格符合SQLite语法规范");
        sb.AppendLine("2. 使用 SQLite 支持的函数和操作符");
        sb.AppendLine("3. 考虑查询性能，合理使用索引");
        sb.AppendLine("4. 使用 EXPLAIN 可验证查询计划");
        sb.AppendLine();
        
        // 9. 输出格式要求
        sb.AppendLine("【输出格式】");
        sb.AppendLine("请清晰标注两个SQL，格式如下：");
        sb.AppendLine();
        sb.AppendLine("-- ====== 统计SQL ======");
        sb.AppendLine("SELECT ...");
        sb.AppendLine();
        sb.AppendLine("-- ====== 明细SQL ======");
        sb.AppendLine("SELECT ...");
        sb.AppendLine();
        
        // 10. 质量要求
        sb.AppendLine("【质量要求】");
        sb.AppendLine("1. SQL语句应该能够直接执行，不能有语法错误");
        sb.AppendLine("2. 金额计算时注意精度，避免数据丢失");
        sb.AppendLine("3. 日期比较时注意时区和格式");
        sb.AppendLine("4. 如果涉及多表关联，使用明确的 JOIN 条件");
        sb.AppendLine("5. 添加必要的注释说明SQL逻辑");
        sb.AppendLine();
        
        sb.AppendLine("请开始生成SQL：");
        
        return sb.ToString();
    }

    /// <summary>
    /// 显示提示词对话框
    /// </summary>
    private async Task ShowPromptDialogAsync(string prompt)
{
    try
    {
        if (_disposed) return;

        var window = _parentWindow ?? GetMainWindow();
        if (window == null || !window.IsVisible || window.WindowState == WindowState.Minimized)
        {
            return;
        }

        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            try
            {
                if (!window.IsVisible || window.WindowState == WindowState.Minimized)
                {
                    return;
                }

                // 先创建文本框
                var textBox = new TextBox
                {
                    Text = prompt,
                    AcceptsReturn = true,
                    TextWrapping = TextWrapping.Wrap,
                    FontFamily = new FontFamily("Consolas, Courier New, monospace"),
                    FontSize = 12,
                    MinHeight = 350,
                    IsReadOnly = false
                };

                // 创建滚动视图 (移除 Margin)
                var scrollViewer = new ScrollViewer
                {
                    Content = textBox,
                    MaxHeight = 420
                };

                // 创建窗口
                var dialog = new Window
                {
                    Title = "📋 合成提示词 - 可复制到DeepSeek网页端使用",
                    Width = 850,
                    Height = 650,
                    MinWidth = 700,
                    MinHeight = 550,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    CanResize = true,
                    Icon = IconHelper.GetAppIcon(),
                    Background = new SolidColorBrush(Color.Parse("#F5F5F5"))  // 窗口背景色
                };

                // === 使用 Grid 布局 ===
                var grid = new Grid
                {
                    Margin = new Thickness(16),
                    Background = new SolidColorBrush(Color.Parse("#F5F5F5"))  // 浅灰色背景
                };

                // 定义行：信息 | 间距 | 文本框 | 间距 | 按钮
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });      // Row 0: 信息
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(5) });    // Row 1: 间距 (5px)
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Row 2: 文本框
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });    // Row 3: 间距 (8px)
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });      // Row 4: 按钮

                // === Row 0: 信息区域 ===
                var infoPanel = new StackPanel
                {
                    Spacing = 8
                };

                var infoText = new TextBlock
                {
                    Text = "📋 以下是将提交给DeepSeek网页版的完整提示词，复制后可直接在网页端提问使用：",
                    FontSize = 13,
                    Foreground = new SolidColorBrush(Color.Parse("#546E7A")),
                    TextWrapping = TextWrapping.Wrap
                };
                infoPanel.Children.Add(infoText);

                var tipsText = new TextBlock
                {
                    Text = "💡 提示：复制后到 DeepSeek 网页版（chat.deepseek.com）粘贴提问，AI将根据提示词生成SQL",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.Parse("#78909C")),
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyle.Italic
                };
                infoPanel.Children.Add(tipsText);

                Grid.SetRow(infoPanel, 0);
                grid.Children.Add(infoPanel);

                // === Row 2: 文本框区域 ===
                Grid.SetRow(scrollViewer, 2);
                grid.Children.Add(scrollViewer);

                // === Row 4: 按钮区域 ===
                var buttonPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 10,
                    HorizontalAlignment = HorizontalAlignment.Center
                };

                var copyButton = new Button
                {
                    Content = "📋 复制到剪贴板",
                    Background = new SolidColorBrush(Color.Parse("#2196F3")),
                    Foreground = Brushes.White,
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(16, 10),
                    FontWeight = FontWeight.SemiBold,
                    MinWidth = 160,
                    Cursor = new Cursor(StandardCursorType.Hand)
                };
                copyButton.Click += (s, e) =>
                {
                    try
                    {
                        textBox.SelectAll();
                        textBox.Focus();
                        _ = SetStatusAsync("✅ 文本已全选，请按 Ctrl+C 复制到剪贴板", new SolidColorBrush(Colors.Green));
                    }
                    catch (Exception ex)
                    {
                        _ = SetStatusAsync($"⚠️ 请手动全选复制: {ex.Message}", new SolidColorBrush(Colors.Orange));
                    }
                };
                buttonPanel.Children.Add(copyButton);

                var closeButton = new Button
                {
                    Content = "✕ 关闭",
                    Background = new SolidColorBrush(Color.Parse("#78909C")),
                    Foreground = Brushes.White,
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(16, 10),
                    FontWeight = FontWeight.SemiBold,
                    MinWidth = 120,
                    Cursor = new Cursor(StandardCursorType.Hand)
                };
                closeButton.Click += (s, e) => dialog.Close();
                buttonPanel.Children.Add(closeButton);

                Grid.SetRow(buttonPanel, 4);
                grid.Children.Add(buttonPanel);

                dialog.Content = grid;

                // 窗口大小变化时动态调整
                dialog.SizeChanged += (s, e) =>
                {
                    var availableHeight = dialog.Bounds.Height - 200;
                    scrollViewer.MaxHeight = Math.Max(200, Math.Min(600, availableHeight));
                };

                await dialog.ShowDialog(window);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ShowPromptDialog error: {ex.Message}");
            }
        });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"ShowPromptDialog error: {ex.Message}");
    }
}

    #endregion

    #region 提交AI

    private async Task SubmitToAiAsync()
    {
        if (IsAnalyzing || _disposed) return;

        try
        {
            // 1. 验证输入
            if (string.IsNullOrWhiteSpace(IndicatorName))
            {
                await SetStatusAsync("请先输入指标名称", new SolidColorBrush(Colors.Red));
                return;
            }
            
            if (string.IsNullOrWhiteSpace(IndicatorDescription))
            {
                await SetStatusAsync("请在【指标描述】标签页中输入自然语言描述", new SolidColorBrush(Colors.Red));
                return;
            }

            var selectedTables = Tables.Where(t => t.IsSelected).ToList();
            if (selectedTables.Count == 0)
            {
                await SetStatusAsync("请至少勾选一个数据库表", new SolidColorBrush(Colors.Red));
                return;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed) return;
                IsAnalyzing = true;
                ProgressLog.Clear();
                ShowStatus = true;
                StatusMessage = "正在准备AI分析...";
                StatusColor = new SolidColorBrush(Colors.Orange);
            });

            // 2. 获取API配置
            var config = await _configService.GetConfigAsync();
            var apiKey = config.DeepSeekApiKey;
            var apiEndpoint = config.DeepSeekApiEndpoint;
            var model = config.DeepSeekModel;

            if (string.IsNullOrWhiteSpace(apiKey) || apiKey == "你的DeepSekk API Key")
            {
                await SetStatusAsync("请先在系统配置中设置DeepSeek API Key", new SolidColorBrush(Colors.Red));
                await ThreadingHelper.RunOnUIThreadAsync(() => { if (!_disposed) IsAnalyzing = false; });
                return;
            }
            
            if (!DeepSeekApiKeyValidator.IsValid(apiKey))
            {
                await SetStatusAsync("DeepSeek API Key格式无效，请检查设置", new SolidColorBrush(Colors.Red));
                await ThreadingHelper.RunOnUIThreadAsync(() => { if (!_disposed) IsAnalyzing = false; });
                return;
            }

            // 3. 获取表结构和字段分析
            await AddProgressAsync($"⏳ 正在分析 {selectedTables.Count} 个表的结构...");
            
            var schemaInfo = await BuildSchemaInfoAsync(selectedTables);
            var fieldAnalysis = await AnalyzeTableFieldsAsync(selectedTables);
            
            await AddProgressAsync($"✅ 表结构分析完成，共识别 {fieldAnalysis.TotalFields} 个字段");
            await AddProgressAsync($"   📊 数值字段: {fieldAnalysis.NumericFields} 个");
            await AddProgressAsync($"   💰 金额字段: {fieldAnalysis.AmountFields} 个");
            await AddProgressAsync($"   📅 日期字段: {fieldAnalysis.DateFields} 个");

            // 4. 构建增强提示词（不含角色信息，由System Prompt提供）
            var prompt = BuildEnhancedPrompt(IndicatorDescription, schemaInfo, selectedTables, fieldAnalysis);
            
            CombinedPrompt = prompt;
            await AddProgressAsync($"📝 提示词已构建（长度: {prompt.Length} 字符）");
            
            if (prompt.Length > 500)
            {
                await AddProgressAsync($"📝 提示词预览: {prompt.Substring(0, 500)}...");
            }
            else
            {
                await AddProgressAsync($"📝 提示词: {prompt}");
            }

            // 5. 调用AI（带重试逻辑）
            var maxRetries = 3;
            DeepSeekService.AiSqlResponse? aiResponse = null;
            var currentPrompt = prompt;
            var retryReasons = new List<string>();

            for (int retry = 0; retry <= maxRetries; retry++)
            {
                if (_disposed) return;
                
                try
                {
                    if (retry > 0)
                    {
                        await AddProgressAsync($"⏳ 第 {retry} 次重试，优化提示词...");
                        currentPrompt = BuildRetryPrompt(prompt, retryReasons);
                        await AddProgressAsync($"📝 重试提示词已优化（长度: {currentPrompt.Length} 字符）");
                    }
                    
                    aiResponse = await _deepSeekService.GenerateSqlWithFallbackAsync(
                        currentPrompt,
                        apiKey);

                    if (aiResponse == null)
                    {
                        await AddProgressAsync("❌ AI返回了空响应");
                        retryReasons.Add("AI返回空响应");
                        continue;
                    }

                    await AddProgressAsync($"✅ SQL生成成功");
                    
                    if (!string.IsNullOrWhiteSpace(aiResponse.Explanation))
                    {
                        await AddProgressAsync($"📝 AI说明: {aiResponse.Explanation}");
                    }

                    // 6. 自动填充SQL
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        if (_disposed) return;
                        if (!string.IsNullOrWhiteSpace(aiResponse.SqlStatement))
                        {
                            IndicatorSqlStatement = aiResponse.SqlStatement;
                        }
                        if (!string.IsNullOrWhiteSpace(aiResponse.SqlDetailData))
                        {
                            IndicatorSqlDetailData = aiResponse.SqlDetailData;
                        }
                    });

                    // 7. 验证SQL
                    await AddProgressAsync("⏳ 正在验证生成的SQL...");
                    var validationResults = await ValidateSqlsAsync(
                        aiResponse.SqlStatement, 
                        aiResponse.SqlDetailData
                    );

                    // 8. 处理验证结果
                    if (validationResults.AllValid)
                    {
                        await SetStatusAsync("✅ SQL生成并验证成功！可切换到SQL标签页查看和编辑", new SolidColorBrush(Colors.Green));
                        await AddProgressAsync("🎉 全部SQL验证通过，分析完成！");
                        
                        await AddProgressAsync($"📊 统计SQL行数: {validationResults.StatRowCount} 行");
                        await AddProgressAsync($"📊 明细SQL行数: {validationResults.DetailRowCount} 行");
                        return;
                    }
                    else
                    {
                        var errorMsg = new StringBuilder();
                        if (!validationResults.StatValid)
                        {
                            errorMsg.AppendLine($"- 统计SQL错误: {validationResults.StatErrorMessage}");
                            retryReasons.Add($"统计SQL: {validationResults.StatErrorMessage}");
                        }
                        if (!validationResults.DetailValid)
                        {
                            errorMsg.AppendLine($"- 明细SQL错误: {validationResults.DetailErrorMessage}");
                            retryReasons.Add($"明细SQL: {validationResults.DetailErrorMessage}");
                        }
                        
                        await AddProgressAsync($"⚠️ SQL验证失败: {errorMsg}");
                        
                        if (retry < maxRetries)
                        {
                            await AddProgressAsync($"🔄 将进行第 {retry + 1} 次重试...");
                            currentPrompt = BuildCorrectionPrompt(prompt, errorMsg.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    await AddProgressAsync($"❌ API调用失败: {ex.Message}");
                    retryReasons.Add($"异常: {ex.Message}");
                    
                    if (retry >= maxRetries)
                    {
                        await SetStatusAsync($"❌ AI分析失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                        return;
                    }
                    
                    await Task.Delay(1000 * (retry + 1));
                }
            }

            if (aiResponse != null)
            {
                await SetStatusAsync("⚠️ SQL已生成但验证未完全通过，请手动检查SQL标签页", new SolidColorBrush(Colors.Orange));
                await AddProgressAsync("⚠️ 达到最大重试次数，请手动检查并修正SQL");
            }
            else
            {
                await SetStatusAsync("❌ AI分析失败，请检查API配置或重试", new SolidColorBrush(Colors.Red));
            }
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 分析过程出错: {ex.Message}", new SolidColorBrush(Colors.Red));
            await AddProgressAsync($"❌ 异常: {ex.Message}");
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed) IsAnalyzing = false;
            });
        }
    }

    #endregion

    #region 辅助方法 - 提示词构建（用于API调用）

/// <summary>
/// 构建增强版提示词（用于API调用）
/// </summary>
private string BuildEnhancedPrompt(
    string description, 
    string schemaInfo, 
    List<TableSelection> selectedTables,
    FieldAnalysis? fieldAnalysis = null)
{
    var sb = new StringBuilder();
    
    // 1. 业务需求
    sb.AppendLine("【业务需求】");
    sb.AppendLine(description);
    sb.AppendLine();
    
    // 2. 重要的数据说明
    sb.AppendLine("【重要说明】");
    sb.AppendLine("1. 数据库中所有字段均为 TEXT 类型，包括数值和日期字段");
    sb.AppendLine("2. 金额字段（注册资本、投资总额等）存储为文本，可能包含千分位分隔符");
    sb.AppendLine("3. 日期字段存储为文本，格式可能不一致（如：2024-01-01、2024/01/01等）");
    sb.AppendLine("4. 处理这些字段时必须进行数据清洗和类型转换");
    sb.AppendLine();
    
    // 3. 表结构信息
    sb.AppendLine("【数据库表结构】");
    sb.AppendLine(schemaInfo);
    sb.AppendLine();
    
    // 4. 字段分析（加强版）
    if (fieldAnalysis != null && fieldAnalysis.TotalFields > 0)
    {
        sb.AppendLine("【字段类型分析】");
        sb.AppendLine($"- 总字段数：{fieldAnalysis.TotalFields}");
        
        if (fieldAnalysis.NumericFieldNames.Count > 0)
        {
            sb.AppendLine($"- 数值字段（需数值处理）：{string.Join("、", fieldAnalysis.NumericFieldNames)}");
        }
        
        if (fieldAnalysis.AmountFieldNames.Count > 0)
        {
            sb.AppendLine($"- 金额字段（需去除千分位）：{string.Join("、", fieldAnalysis.AmountFieldNames)}");
        }
        
        if (fieldAnalysis.DateFieldNames.Count > 0)
        {
            sb.AppendLine($"- 日期字段（需统一格式）：{string.Join("、", fieldAnalysis.DateFieldNames)}");
        }
        
        // 特殊字段提示
        var specialFields = new[] { "纳税人状态", "有效标志", "登记注册类型", "行业" };
        var foundSpecial = specialFields.Where(f => 
            fieldAnalysis.AllFieldNames.Contains(f)).ToList();
        
        if (foundSpecial.Count > 0)
        {
            sb.AppendLine($"- 分类/状态字段：{string.Join("、", foundSpecial)}（可用于分组和筛选）");
        }
        
        sb.AppendLine();
    }
    
    // 5. 数据清洗示例
    sb.AppendLine("【数据清洗示例】");
    sb.AppendLine("-- 金额字段清洗示例：");
    sb.AppendLine("SELECT");
    sb.AppendLine("  CAST(REPLACE(REPLACE(\"注册资本\", ',', ''), ' ', '') AS REAL) AS 注册资本_数值");
    sb.AppendLine("FROM \"sw_dj_nsrxx\";");
    sb.AppendLine();
    sb.AppendLine("-- 日期字段清洗示例：");
    sb.AppendLine("SELECT");
    sb.AppendLine("  date(REPLACE(\"开业设立日期\", '/', '-')) AS 开业设立日期_标准");
    sb.AppendLine("FROM \"sw_dj_nsrxx\";");
    sb.AppendLine();
    
    // 6. 业务需求细化
    sb.AppendLine("【业务需求细化】");
    sb.AppendLine("请根据上述业务需求，生成包含以下内容的SQL：");
    sb.AppendLine("1. 使用 CTE（公用表表达式）进行数据预处理和清洗");
    sb.AppendLine("2. 对金额字段进行数值转换和精度处理");
    sb.AppendLine("3. 对日期字段进行格式统一");
    sb.AppendLine("4. 提供详细的统计指标和维度");
    sb.AppendLine("5. 添加充分的注释说明");
    sb.AppendLine();
    
    // 7. SQL生成要求
    sb.AppendLine("【SQL生成要求】");
    sb.AppendLine("请生成两个SQL查询语句：");
    sb.AppendLine("1. **统计SQL**：使用聚合函数（COUNT、SUM、AVG、MAX、MIN等）进行数据汇总统计");
    sb.AppendLine("2. **明细SQL**：返回完整的明细数据，包含必要的筛选和排序条件");
    sb.AppendLine();
    sb.AppendLine("两个SQL必须逻辑一致，统计结果应与明细数据相匹配。");
    sb.AppendLine();
    
    // 8. 数据格式处理要求
    sb.AppendLine("【数据格式处理要求】");
    sb.AppendLine("1. **金额字段处理**：");
    sb.AppendLine("   - 去除千分位分隔符（如 8,888,888.88 → 8888888.88）");
    sb.AppendLine("   - 使用 ROUND(字段, 2) 保留两位小数");
    sb.AppendLine("   - 使用 CAST 或 REPLACE 函数处理文本金额");
    sb.AppendLine("2. **日期字段处理**：");
    sb.AppendLine("   - 统一转换为 'YYYY-MM-DD' 格式");
    sb.AppendLine("   - 使用 date() 或 datetime() 函数");
    sb.AppendLine("   - 处理空值和无效日期（使用 COALESCE）");
    sb.AppendLine("3. **NULL值处理**：");
    sb.AppendLine("   - 使用 COALESCE 或 IFNULL 函数处理NULL值");
    sb.AppendLine("4. **表名和字段名**：");
    sb.AppendLine("   - 使用双引号包裹（如 \"table_name\"）");
    sb.AppendLine("   - 避免与SQL关键字冲突");
    sb.AppendLine();
    
    // 9. SQLite语法规范
    sb.AppendLine("【SQLite语法规范】");
    sb.AppendLine("1. 所有SQL必须严格符合SQLite语法规范");
    sb.AppendLine("2. 使用 SQLite 支持的函数和操作符");
    sb.AppendLine("3. 考虑查询性能，合理使用索引");
    sb.AppendLine("4. 使用 EXPLAIN 可验证查询计划");
    sb.AppendLine();
    
    // 10. 输出格式要求
    sb.AppendLine("【输出格式】");
    sb.AppendLine("请严格按照以下JSON格式返回，不要包含任何其他文字：");
    sb.AppendLine(@"
{
    ""sql_statement"": ""统计SQL语句"",
    ""sql_detaildata"": ""明细SQL语句"",
    ""explanation"": ""SQL实现思路和关键逻辑说明""
}");
    
    sb.AppendLine();
    sb.AppendLine("请开始生成SQL：");
    
    return sb.ToString();
}

    /// <summary>
    /// 构建重试提示词
    /// </summary>
    private string BuildRetryPrompt(string originalPrompt, List<string> errorReasons)
    {
        var sb = new StringBuilder();
        sb.AppendLine(originalPrompt);
        sb.AppendLine();
        sb.AppendLine("【上次生成的问题】");
        sb.AppendLine("之前的SQL存在以下问题，请修正：");
        
        for (int i = 0; i < errorReasons.Count; i++)
        {
            sb.AppendLine($"{i + 1}. {errorReasons[i]}");
        }
        
        sb.AppendLine();
        sb.AppendLine("请重新生成正确的SQL，特别注意：");
        sb.AppendLine("1. 确保语法正确");
        sb.AppendLine("2. 正确处理数据类型转换");
        sb.AppendLine("3. 使用正确的函数和操作符");
        
        return sb.ToString();
    }

    /// <summary>
    /// 构建修正提示词
    /// </summary>
    private string BuildCorrectionPrompt(string originalPrompt, string errorMessage)
    {
        var sb = new StringBuilder();
        sb.AppendLine(originalPrompt);
        sb.AppendLine();
        sb.AppendLine("【SQL验证错误】");
        sb.AppendLine("生成的SQL验证失败，具体错误如下：");
        sb.AppendLine(errorMessage);
        sb.AppendLine();
        sb.AppendLine("请修正这些错误，重新生成正确的SQL。");
        
        return sb.ToString();
    }

    #endregion

    #region 辅助方法 - 表结构分析

    /// <summary>
    /// 构建表结构信息
    /// </summary>
    private async Task<string> BuildSchemaInfoAsync(List<TableSelection> tables)
    {
        var sb = new StringBuilder();
        
        foreach (var table in tables)
        {
            try
            {
                var schemaSql = $"PRAGMA table_info(\"{table.TableName}\")";
                var schemaData = await _databaseService.ExecuteQueryAsync(schemaSql, new List<object>());

                sb.AppendLine($"### 表名：{table.TableName}");
                sb.AppendLine();
                sb.AppendLine("| 字段名 | 数据类型 | 是否为空 | 默认值 |");
                sb.AppendLine("|--------|----------|----------|--------|");
                
                if (schemaData != null && schemaData.Count > 1)
                {
                    for (int i = 1; i < schemaData.Count; i++)
                    {
                        var row = schemaData[i];
                        if (row.Count >= 6)
                        {
                            var colName = row[1]?.ToString() ?? "";
                            var colType = row[2]?.ToString() ?? "TEXT";
                            var notNull = row[3]?.ToString() == "1" ? "NOT NULL" : "NULL";
                            var defaultValue = row[4]?.ToString() ?? "";
                            sb.AppendLine($"| {colName} | {colType} | {notNull} | {defaultValue} |");
                        }
                    }
                }
                sb.AppendLine();
            }
            catch (Exception ex)
            {
                sb.AppendLine($"### 表名：{table.TableName}");
                sb.AppendLine($"（获取结构失败: {ex.Message}）");
                sb.AppendLine();
            }
        }
        
        return sb.ToString();
    }

    /// <summary>
    /// 分析表字段
    /// </summary>
   private async Task<FieldAnalysis> AnalyzeTableFieldsAsync(List<TableSelection> tables)
{
    var analysis = new FieldAnalysis();
    var amountKeywords = new[] { "注册资本", "投资总额", "金额", "价格", "费用", "收入", "支出", "利润", "总额", "单价", "总价" };
    var dateKeywords = new[] { "日期", "时间", "年月", "日", "date", "time", "year", "month", "day", "created", "updated" };
    var numericKeywords = new[] { "Id", "人数", "数量", "比例", "percent", "number", "count" };

    foreach (var table in tables)
    {
        try
        {
            var schemaSql = $"PRAGMA table_info(\"{table.TableName}\")";
            var schemaData = await _databaseService.ExecuteQueryAsync(schemaSql, new List<object>());

            if (schemaData != null && schemaData.Count > 1)
            {
                for (int i = 1; i < schemaData.Count; i++)
                {
                    var row = schemaData[i];
                    if (row.Count >= 3)
                    {
                        var colName = row[1]?.ToString() ?? "";
                        var colType = row[2]?.ToString() ?? "TEXT";
                        
                        analysis.TotalFields++;
                        analysis.AllFieldNames.Add(colName);
                        
                        // 检查是否为金额字段（基于字段名，而非类型）
                        if (amountKeywords.Any(k => colName.Contains(k, StringComparison.OrdinalIgnoreCase)))
                        {
                            analysis.AmountFields++;
                            analysis.AmountFieldNames.Add(colName);
                            analysis.NumericFields++;
                            analysis.NumericFieldNames.Add(colName);
                        }
                        // 检查是否为日期字段
                        else if (dateKeywords.Any(k => colName.Contains(k, StringComparison.OrdinalIgnoreCase)))
                        {
                            analysis.DateFields++;
                            analysis.DateFieldNames.Add(colName);
                        }
                        // 检查是否为数值字段
                        else if (numericKeywords.Any(k => colName.Contains(k, StringComparison.OrdinalIgnoreCase)))
                        {
                            analysis.NumericFields++;
                            analysis.NumericFieldNames.Add(colName);
                        }
                    }
                }
            }
        }
        catch
        {
        }
    }
    
    return analysis;
}

    /// <summary>
    /// 字段分析结果
    /// </summary>
    public class FieldAnalysis
    {
        public int TotalFields { get; set; }
        public int NumericFields { get; set; }
        public int AmountFields { get; set; }
        public int DateFields { get; set; }
        public List<string> NumericFieldNames { get; set; } = new List<string>();
        public List<string> AmountFieldNames { get; set; } = new List<string>();
        public List<string> DateFieldNames { get; set; } = new List<string>();
        
         // 新增：所有字段名列表
        public List<string> AllFieldNames { get; set; } = new List<string>();
    }

    #endregion

    #region 辅助方法 - SQL验证

    /// <summary>
    /// SQL验证结果
    /// </summary>
    public class SqlValidationResult
    {
        public bool StatValid { get; set; }
        public bool DetailValid { get; set; }
        public bool AllValid => StatValid && DetailValid;
        public string StatErrorMessage { get; set; } = string.Empty;
        public string DetailErrorMessage { get; set; } = string.Empty;
        public int StatRowCount { get; set; }
        public int DetailRowCount { get; set; }
    }

    /// <summary>
    /// 验证SQL语句
    /// </summary>
    private async Task<SqlValidationResult> ValidateSqlsAsync(string? statSql, string? detailSql)
    {
        var result = new SqlValidationResult();
        
        if (!string.IsNullOrWhiteSpace(statSql))
        {
            try
            {
                var testSql = $"EXPLAIN {statSql}";
                var data = await _databaseService.ExecuteQueryAsync(testSql, new List<object>());
                result.StatValid = true;
                result.StatRowCount = data.Count > 0 ? data.Count - 1 : 0;
            }
            catch (Exception ex)
            {
                result.StatValid = false;
                result.StatErrorMessage = ex.Message;
            }
        }
        else
        {
            result.StatValid = false;
            result.StatErrorMessage = "统计SQL为空";
        }
        
        if (!string.IsNullOrWhiteSpace(detailSql))
        {
            try
            {
                var testSql = $"EXPLAIN {detailSql}";
                var data = await _databaseService.ExecuteQueryAsync(testSql, new List<object>());
                result.DetailValid = true;
                result.DetailRowCount = data.Count > 0 ? data.Count - 1 : 0;
            }
            catch (Exception ex)
            {
                result.DetailValid = false;
                result.DetailErrorMessage = ex.Message;
            }
        }
        else
        {
            result.DetailValid = false;
            result.DetailErrorMessage = "明细SQL为空";
        }
        
        return result;
    }

    #endregion

    #region 保存

    private async Task SaveAsync()
    {
        if (IsSaving || _disposed) return;

        try
        {
            if (string.IsNullOrWhiteSpace(IndicatorName))
            {
                await SetStatusAsync("请输入指标名称", new SolidColorBrush(Colors.Red));
                return;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed) return;
                IsSaving = true;
                ShowStatus = true;
                StatusMessage = "正在保存指标...";
                StatusColor = new SolidColorBrush(Colors.Orange);
            });

            var indicator = new Indicator
            {
                Id = Guid.NewGuid().ToString(),
                Name = IndicatorName.Trim(),
                SqlStatement = IndicatorSqlStatement?.Trim() ?? string.Empty,
                SqlDetailData = IndicatorSqlDetailData?.Trim() ?? string.Empty,
                Description = IndicatorDescription?.Trim() ?? string.Empty,
                Category = string.IsNullOrWhiteSpace(IndicatorCategory) ? _defaultCategory : IndicatorCategory,
                CreatedAt = DateTime.Now,
                UpdatedAt = DateTime.Now,
                CreatedBy = "admin",
                IsActive = true
            };

            await _indicatorService.AddAsync(indicator);
            
            await SetStatusAsync($"✅ 指标 '{indicator.Name}' 保存成功！", new SolidColorBrush(Colors.Green));
            await AddProgressAsync($"✅ 指标已保存到数据库 (ID: {indicator.Id})");
            
            await ShowSaveSuccessDialogAsync(indicator.Name);
        }
        catch (Exception ex)
        {
            await SetStatusAsync($"❌ 保存失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            await AddProgressAsync($"❌ 保存失败: {ex.Message}");
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (!_disposed) IsSaving = false;
            });
        }
    }

    private async Task ShowSaveSuccessDialogAsync(string indicatorName)
    {
        try
        {
            if (_disposed) return;
        
            var window = _parentWindow ?? GetMainWindow();
            if (window == null) return;

            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                try
                {
                    if (!window.IsVisible || window.WindowState == WindowState.Minimized) return;

                    var dialog = new Window
                    {
                        Title = "✅ 保存成功",
                        Width = 400,
                        Height = 200,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                        CanResize = false,
                        ShowInTaskbar = false
                    };

                    var okButton = new Button
                    {
                        Content = "确定",
                        Background = new SolidColorBrush(Color.Parse("#4CAF50")),
                        Foreground = Brushes.White,
                        CornerRadius = new Avalonia.CornerRadius(6),
                        Padding = new Avalonia.Thickness(24, 10),
                        FontWeight = Avalonia.Media.FontWeight.SemiBold,
                        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                        Width = 120,
                        Height = 38
                    };
                    okButton.Click += (s, e) => dialog.Close();

                    dialog.Content = new StackPanel
                    {
                        Margin = new Avalonia.Thickness(20),
                        Spacing = 16,
                        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                        VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                        Children =
                        {
                            new TextBlock
                            {
                                Text = "✅ 指标保存成功！",
                                FontSize = 18,
                                FontWeight = Avalonia.Media.FontWeight.Bold,
                                Foreground = new SolidColorBrush(Color.Parse("#2E7D32")),
                                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center
                            },
                            new TextBlock
                            {
                                Text = $"指标名称：{indicatorName}",
                                FontSize = 14,
                                Foreground = new SolidColorBrush(Color.Parse("#546E7A")),
                                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                                TextAlignment = Avalonia.Media.TextAlignment.Center
                            },
                            okButton
                        }
                    };

                    await dialog.ShowDialog(window);
                }
                catch
                {
                }
            });
        }
        catch
        {
        }
    }

    #endregion

    #region 新建指标

    private void NewIndicator()
    {
        try
        {
            if (_disposed) return;
            
            IndicatorName = string.Empty;
            IndicatorDescription = string.Empty;
            IndicatorSqlStatement = string.Empty;
            IndicatorSqlDetailData = string.Empty;
            IndicatorCategory = _defaultCategory;
            ClearTableSelection();
            
            ProgressLog.Clear();
            
            ShowStatus = true;
            StatusMessage = "📝 已清空表单，可以创建新指标";
            StatusColor = new SolidColorBrush(Color.Parse("#78909C"));
            
            CombinedPrompt = string.Empty;
        }
        catch
        {
        }
    }

    #endregion

    #region 数据预览（核心修复）

    private async Task PreviewDataAsync()
{
    try
    {
        if (_disposed) return;
        
        var sql = SelectedTabIndex switch
        {
            1 => IndicatorSqlStatement,
            2 => IndicatorSqlDetailData,
            _ => string.Empty
        };

        var sqlLabel = SelectedTabIndex switch
        {
            1 => "统计SQL",
            2 => "明细SQL",
            _ => string.Empty
        };

        if (SelectedTabIndex == 0)
        {
            await SetStatusAsync("请在【统计SQL】或【明细SQL】标签页中预览", new SolidColorBrush(Colors.Orange));
            return;
        }

        if (string.IsNullOrWhiteSpace(sql))
        {
            await SetStatusAsync($"【{sqlLabel}】为空，无法预览", new SolidColorBrush(Colors.Orange));
            return;
        }

        var window = _parentWindow ?? GetMainWindow();
        if (window == null) 
        {
            return;
        }

        await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            try
            {
                if (_disposed) return;
                if (!window.IsVisible || window.WindowState == WindowState.Minimized) 
                {
                    return;
                }

                var previewViewModel = new DetailDataViewModel(sql, $"{IndicatorName}_{sqlLabel}", IndicatorCategory);
                var detailView = new DetailDataView 
                { 
                    DataContext = previewViewModel 
                };

                previewViewModel.SetParentWindow(window);
                detailView.SetParentWindow(window);

                var previewWindow = new Window
                {
                    Title = $"数据预览 - {IndicatorName} ({sqlLabel})",
                    Width = 1200,
                    Height = 700,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Content = detailView,
                    CanResize = true,
                    MinWidth = 800,
                    MinHeight = 500,
                    Icon = IconHelper.GetAppIcon()
                };

                // ✅ 移除 Closed 事件中的 Dispose
                // 让 GC 自动回收

                await previewWindow.ShowDialog(window);
            }
            catch (Exception ex)
            {
                await SetStatusAsync($"预览失败: {ex.Message}", new SolidColorBrush(Colors.Red));
            }
        });
    }
    catch (Exception ex)
    {
        await SetStatusAsync($"预览失败: {ex.Message}", new SolidColorBrush(Colors.Red));
    }
}

    /// <summary>
    /// 显示 Interaction 未注册的错误提示
    /// </summary>
    private async Task ShowInteractionErrorDialogAsync()
    {
        try
        {
            if (_disposed) return;
            
            var window = _parentWindow ?? GetMainWindow();
            if (window == null) return;

            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                try
                {
                    if (!window.IsVisible || window.WindowState == WindowState.Minimized) return;

                    var dialog = new Window
                    {
                        Title = "⚠️ 预览错误",
                        Width = 500,
                        Height = 300,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                        CanResize = false,
                        ShowInTaskbar = false
                    };

                    var closeButton = new Button
                    {
                        Content = "确定",
                        Background = new SolidColorBrush(Color.Parse("#2196F3")),
                        Foreground = Brushes.White,
                        CornerRadius = new Avalonia.CornerRadius(6),
                        Padding = new Avalonia.Thickness(24, 10),
                        FontWeight = Avalonia.Media.FontWeight.SemiBold,
                        HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                        Width = 120,
                        Height = 38
                    };
                    closeButton.Click += (s, e) => dialog.Close();

                    dialog.Content = new StackPanel
                    {
                        Margin = new Avalonia.Thickness(20),
                        Spacing = 16,
                        Children =
                        {
                            new TextBlock
                            {
                                Text = "⚠️ 预览功能需要交互注册",
                                FontSize = 16,
                                FontWeight = Avalonia.Media.FontWeight.Bold,
                                Foreground = new SolidColorBrush(Colors.Orange)
                            },
                            new TextBlock
                            {
                                Text = "数据预览窗口需要注册交互处理器。\n\n" +
                                       "请确保 DetailDataView 中正确注册了所有 Interaction。\n\n" +
                                       "如果问题持续出现，请联系开发人员检查 DetailDataViewModel 的 Interaction 注册。",
                                FontSize = 13,
                                Foreground = new SolidColorBrush(Color.Parse("#546E7A")),
                                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                                TextAlignment = Avalonia.Media.TextAlignment.Left
                            },
                            closeButton
                        }
                    };

                    await dialog.ShowDialog(window);
                }
                catch
                {
                }
            });
        }
        catch
        {
        }
    }

    #endregion

    #region 辅助方法

    private async Task SetStatusAsync(string message, IBrush color)
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed) return;
                StatusMessage = message;
                StatusColor = color;
                ShowStatus = true;
            });
        }
        catch
        {
        }
    }

    private async Task AddProgressAsync(string message)
    {
        try
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                if (_disposed) return;
                ProgressLog.Add($"[{timestamp}] {message}");
            });
        }
        catch
        {
        }
    }

    private Window? GetMainWindow()
    {
        try
        {
            if (Avalonia.Application.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            {
                return desktop.MainWindow;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 清理

    public void Cleanup()
    {
        if (_isCleaned) return;

        try
        {
            
            ProgressLog.Clear();
            Tables.Clear();
            CategoryOptions.Clear();
            
            _isCleaned = true;
        }
        catch
        {
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        
        try
        {
            
            _disposed = true;
            _isCleaned = true;
            
            _initLock?.Dispose();
            
            SubmitToAiCommand?.Dispose();
            SaveCommand?.Dispose();
            NewIndicatorCommand?.Dispose();
            PreviewCommand?.Dispose();
            SelectAllTablesCommand?.Dispose();
            ClearTableSelectionCommand?.Dispose();
            RefreshTablesCommand?.Dispose();
            CopyPromptCommand?.Dispose();
            
        }
        catch
        {
        }
    }

    #endregion
}