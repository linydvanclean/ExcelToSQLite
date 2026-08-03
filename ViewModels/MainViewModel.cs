using Avalonia.Controls;
using Avalonia.Layout;
using ExcelToSQLite.Views;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Reactive;
using System.Threading.Tasks;
using MenuItemModel = ExcelToSQLite.Models.MenuItem;
using ExcelToSQLite.Services;
using ExcelToSQLite.Models;
using ExcelToSQLite.Helpers;
using Avalonia.Media;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 支持页面刷新（用于缓存页面切换时重新加载数据）
/// </summary>
public interface IRefreshablePage
{
    Task RefreshAsync();
}

/// <summary>
/// 支持资源清理
/// </summary>
public interface ICleanupPage
{
    void Cleanup();
}

public class MainViewModel : ReactiveObject, IDisposable
{
    private readonly Window _parentWindow;
    private readonly string _username;
    private bool _isDisposed;
    private bool _isBusy;

    private object? _currentView;
    private bool _showWelcome = true;
    private MenuItemModel? _selectedMenuItem;
    private AppConfig _appConfigs = new AppConfig();
    private string _systemTitle;
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Colors.Green);

    public ObservableCollection<MenuItemModel> MenuItems { get; } = new();

    #region 属性

    public bool IsBusy
    {
        get => _isBusy;
        set => this.RaiseAndSetIfChanged(ref _isBusy, value);
    }

    public object? CurrentView
    {
        get => _currentView;
        set => this.RaiseAndSetIfChanged(ref _currentView, value);
    }

    public bool ShowWelcome
    {
        get => _showWelcome;
        set => this.RaiseAndSetIfChanged(ref _showWelcome, value);
    }

    public AppConfig AppConfigs
    {
        get => _appConfigs;
        set => this.RaiseAndSetIfChanged(ref _appConfigs, value);
    }

    public MenuItemModel? SelectedMenuItem
    {
        get => _selectedMenuItem;
        set => this.RaiseAndSetIfChanged(ref _selectedMenuItem, value);
    }

    public string SystemTitle
    {
        get => _systemTitle;
        private set => this.RaiseAndSetIfChanged(ref _systemTitle, value);
    }

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

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> StartCommand { get; }
    public ReactiveCommand<string, Unit> NavigateCommand { get; }
    public ReactiveCommand<Unit, Unit> LogoutCommand { get; }
    public ReactiveCommand<MenuItemModel, Unit> ToggleMenuCommand { get; }

    #endregion

    public MainViewModel(Window parentWindow, 
        string username = "admin",
        string defaultPage = "AnalysisBatch")
    {
        _parentWindow = parentWindow ?? throw new ArgumentNullException(nameof(parentWindow));
        _username = username ?? throw new ArgumentNullException(nameof(username));
        _systemTitle = "数据智能分析系统 @VanClean";

        InitializeMenuItems();
        
        NavigateCommand = ReactiveCommand.CreateFromTask<string>(NavigateAsync);
        LogoutCommand = ReactiveCommand.CreateFromTask(LogoutAsync);
        ToggleMenuCommand = ReactiveCommand.Create<MenuItemModel>(ToggleMenu);
        
        StartCommand = ReactiveCommand.Create(() =>
        {
            if (!ShowWelcome) return;
            ShowWelcome = false;
            _ = NavigateAsync(defaultPage);
        });

        _ = InitializeSystemAsync();
    }

    #region 初始化

    private async Task InitializeSystemAsync()
    {
        try
        {
            var settings = new AppConfigService();
            AppConfigs = await settings.GetConfigAsync();
            SystemTitle = $"{AppConfigs.SystemName} @VanClean";
            StatusMessage = "系统初始化完成";
            StatusColor = new SolidColorBrush(Colors.Green);
        }
        catch
        {
            StatusMessage = "系统初始化失败，使用默认配置";
            StatusColor = new SolidColorBrush(Colors.Orange);
        }
    }

    #endregion

    #region 菜单初始化

    private void InitializeMenuItems()
    {
        // 系统管理
        var systemMenu = new MenuItemModel
        {
            Header = "系统管理",
            Icon = "⚙️",  // 齿轮 - 系统设置
            IsExpanded = false
        };
        systemMenu.Children.Add(new MenuItemModel 
        { 
            Header = "欢迎使用", 
            Icon = "👋",  // 挥手 - 欢迎
            CommandParameter = "Welcome" 
        });
        systemMenu.Children.Add(new MenuItemModel 
        { 
            Header = "系统配置", 
            Icon = "🔧",  // 扳手 - 配置
            CommandParameter = "Settings" 
        });
        systemMenu.Children.Add(new MenuItemModel
        {
            Header = "修改密码",
            Icon = "🔒",  // 锁 - 安全
            CommandParameter = "ChangePassword"
        });
        
        // 数据管理
        var dataMenu = new MenuItemModel
        {
            Header = "数据管理",
            Icon = "🗄️",  // 文件柜 - 数据存储
            IsExpanded = false
        };
        dataMenu.Children.Add(new MenuItemModel 
        { 
            Header = "数据库管理", 
            Icon = "💿",  // 光盘 - 数据库
            CommandParameter = "Database" 
        });
        dataMenu.Children.Add(new MenuItemModel
        {
            Header = "表字段查看",
            Icon = "🔍",  // 放大镜 - 查看
            CommandParameter = "TableFields"
        });
        dataMenu.Children.Add(new MenuItemModel
        {
            Header = "表字段管理",
            Icon = "✏️",  // 铅笔 - 编辑管理
            CommandParameter = "TableFieldManagement"
        });
        dataMenu.Children.Add(new MenuItemModel
        {
            Header = "数据预览",
            Icon = "📋",  // 剪贴板 - 数据预览
            CommandParameter = "DbTablesView"
        });
        
        // 分析管理
        var analysisAdminMenu = new MenuItemModel
        {
            Header = "分析管理",
            Icon = "🧠",  // 大脑 - 分析
            IsExpanded = false
        };
        analysisAdminMenu.Children.Add(new MenuItemModel 
        { 
            Header = "数据字典", 
            Icon = "📚",  // 多本书 - 字典
            CommandParameter = "DataDictionary" 
        });
        analysisAdminMenu.Children.Add(new MenuItemModel 
        { 
            Header = "指标管理", 
            Icon = "📊",  // 柱状图 - 指标
            CommandParameter = "IndicatorManagement" 
        });
        analysisAdminMenu.Children.Add(new MenuItemModel 
        { 
            Header = "AI建模", 
            Icon = "🤖",  // 机器人 - AI
            CommandParameter = "CreateAiModel" 
        });
        analysisAdminMenu.Children.Add(new MenuItemModel 
        { 
            Header = "分析批次", 
            Icon = "🏷️",  // 标签 - 批次
            CommandParameter = "AnalysisBatch" 
        });

        // 数据导入
        var dataImportMenu = new MenuItemModel
        {
            Header = "数据导入",
            Icon = "📥",  // 下载箭头 - 导入
            IsExpanded = false
        };
        dataImportMenu.Children.Add(new MenuItemModel 
        { 
            Header = "Excel导入", 
            Icon = "📈",  // 折线图 - Excel数据
            CommandParameter = "Excel" 
        });
        dataImportMenu.Children.Add(new MenuItemModel 
        { 
            Header = "金三数据导入", 
            Icon = "🏛️",  // 建筑 - 金三
            CommandParameter = "TaxExcel" 
        });
        dataImportMenu.Children.Add(new MenuItemModel 
        { 
            Header = "JSON导入", 
            Icon = "📄",  // 文档 - JSON
            CommandParameter = "Json" 
        });
        dataImportMenu.Children.Add(new MenuItemModel 
        { 
            Header = "考勤数据专项导入", 
            Icon = "👤",  // 单人 - 考勤
            CommandParameter = "Attendance" 
        });
        dataImportMenu.Children.Add(new MenuItemModel 
        { 
            Header = "加油卡数据专项导入", 
            Icon = "⛽",  // 加油泵 - 加油卡
            CommandParameter = "FuelCard" 
        });
        
        // 数据分析
        var analysisMenu = new MenuItemModel
        {
            Header = "数据分析",
            Icon = "📉",  // 下降趋势图 - 分析
            IsExpanded = false
        };
        analysisMenu.Children.Add(new MenuItemModel 
        { 
            Header = "扫描结果", 
            Icon = "📃",  // 卷纸文档 - 扫描结果
            CommandParameter = "ScanResults" 
        });

        MenuItems.Add(systemMenu);
        MenuItems.Add(dataMenu);
        MenuItems.Add(analysisAdminMenu);
        MenuItems.Add(dataImportMenu);
        MenuItems.Add(analysisMenu);
    }

    #endregion

    #region 菜单操作

    private void ToggleMenu(MenuItemModel menuItem)
    {
        if (menuItem == null) return;
        
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (menuItem.Children.Count > 0)
            {
                menuItem.IsExpanded = !menuItem.IsExpanded;
            }
            else if (!string.IsNullOrEmpty(menuItem.CommandParameter))
            {
                _ = NavigateAsync(menuItem.CommandParameter);
            }
        }).ConfigureAwait(false);
    }

    #endregion

    #region 导航

    private async Task NavigateAsync(string page)
    {
        if (IsBusy) return;

        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsBusy = true;
                ShowWelcome = false;
            });

            if (page == "Welcome")
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    ShowWelcome = true;
                    CleanupCurrentView();
                    CurrentView = null;
                    IsBusy = false;
                });
                return;
            }

            // 清理上一个页面
            CleanupCurrentView();

            // 创建新页面
            UserControl? newView = null;
            try
            {
                newView = page switch
                {
                    "Settings" => new ConfigEditView(_parentWindow),
                    "ChangePassword" => new ChangePasswordDialog(_parentWindow, _username),
                    "Database" => new DataManagementView(_parentWindow),
                    "TableFields" => new TableFieldView(_parentWindow),
                    "TableFieldManagement" => new TableFieldManagementView(_parentWindow),
                    "DbTablesView" => new DbTablesView(_parentWindow),
                    "DataDictionary" => new DataDictionaryView(_parentWindow),
                    "IndicatorManagement" => new IndicatorManagementView(_parentWindow),
                    "AnalysisBatch" => new AnalysisBatchView(_parentWindow),
                    "Excel" => new ExcelDataView(_parentWindow),
                    "TaxExcel" => new TaxExcelDataView(_parentWindow),
                    "Json" => new JsonDataView(_parentWindow),
                    "Attendance" => new AttendanceView(_parentWindow),
                    "FuelCard" => new FuelCardView(_parentWindow),
                    "ScanResults" => new ScanResultView(_parentWindow),
                    "CreateAiModel" => new CreateAiModelView(_parentWindow),
                    _ => null
                };
            }
            catch (Exception ex)
            {
                newView = null;
                StatusMessage = $"创建页面失败: {ex.Message}";
                StatusColor = new SolidColorBrush(Colors.Red);
            }

            if (newView != null)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    CurrentView = newView;
                    StatusMessage = $"已切换到: {page}";
                    StatusColor = new SolidColorBrush(Colors.Green);
                });

                // 如果页面实现了 IRefreshablePage，加载数据
                if (newView.DataContext is IRefreshablePage refreshable)
                {
                    try
                    {
                        await refreshable.RefreshAsync();
                    }
                    catch (Exception ex)
                    {
                        StatusMessage = $"加载数据失败: {ex.Message}";
                        StatusColor = new SolidColorBrush(Colors.Orange);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"导航失败: {ex.Message}";
            StatusColor = new SolidColorBrush(Colors.Red);
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsBusy = false;
            });
        }
    }

    #endregion

    #region 资源清理

    private void CleanupCurrentView()
    {
        if (CurrentView is not UserControl currentView) return;

        try
        {
            // 1. 如果实现了 ICleanupPage 接口，调用清理方法
            if (currentView.DataContext is ICleanupPage cleanupPage)
            {
                cleanupPage.Cleanup();
            }

            // 2. 如果实现了 IDisposable，释放资源
            if (currentView.DataContext is IDisposable disposable)
            {
                disposable.Dispose();
            }

            // 3. 清理 DataContext 引用
            currentView.DataContext = null;

            // 4. 从父容器移除（如果有）
            if (currentView.Parent is Panel panel)
            {
                panel.Children.Remove(currentView);
            }
        }
        catch
        {
        }
    }

    #endregion

    #region 登出

    private async Task LogoutAsync()
    {
        var result = await MessageBox.ShowAsync(
            _parentWindow,
            "确定要退出登录吗？",
            "确认退出",
            MessageBoxButtons.YesNo
        );

        if (result == MessageBoxResult.Yes)
        {
            try
            {
                // 退出时清理所有资源
                CleanupCurrentView();
                
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    var loginWindow = new LoginWindow();
                    loginWindow.Show();
                    _parentWindow.Close();
                });
            }
            catch (Exception ex)
            {
                StatusMessage = $"登出失败: {ex.Message}";
                StatusColor = new SolidColorBrush(Colors.Red);
            }
        }
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                StartCommand?.Dispose();
                NavigateCommand?.Dispose();
                LogoutCommand?.Dispose();
                ToggleMenuCommand?.Dispose();
                
                // 清理当前视图
                CleanupCurrentView();
                
                // 清空菜单项
                MenuItems.Clear();
            }
            catch
            {
            }
            _isDisposed = true;
        }
        GC.SuppressFinalize(this);
    }

    #endregion
}