using Avalonia.Controls;
using ExcelToSQLite.Services;
using ExcelToSQLite.Views;
using ExcelToSQLite.Helpers;
using ReactiveUI;
using System;
using System.Reactive;
using System.Threading.Tasks;
using Avalonia.Media;

namespace ExcelToSQLite.ViewModels;

public class LoginViewModel : ReactiveObject, IDisposable
{
    private readonly UserService _userService;
    private readonly Window _parentWindow;
    private bool _isDisposed;

    private string _username = "admin";
    private string _password = "123456";
    private string _errorMessage = string.Empty;
    private bool _hasError;
    private bool _isLoading;
    private bool _showError;
    

    public LoginViewModel(Window parentWindow)
    {
        _userService = new UserService();
        _parentWindow = parentWindow ?? throw new ArgumentNullException(nameof(parentWindow));
        
        LoginCommand = ReactiveCommand.CreateFromTask(LoginAsync);
        
        // 初始化错误状态
        ClearError();
    }

    #region 属性
    
    public double ErrorOpacity => ShowError ? 1.0 : 0.0;
    public bool ShowError
    {
        get => _showError;
        set 
        {
            this.RaiseAndSetIfChanged(ref _showError, value);
            this.RaisePropertyChanged(nameof(ErrorOpacity));
        }
    }
    public string Username
    {
        get => _username;
        set
        {
            this.RaiseAndSetIfChanged(ref _username, value);
            ClearError();
        }
    }

    public string Password
    {
        get => _password;
        set
        {
            this.RaiseAndSetIfChanged(ref _password, value);
            ClearError();
        }
    }

    public string ErrorMessage
    {
        get => _errorMessage;
        set => this.RaiseAndSetIfChanged(ref _errorMessage, value);
    }

    public bool HasError
    {
        get => _hasError;
        set
        {
            this.RaiseAndSetIfChanged(ref _hasError, value);
            ShowError = value;
        }
    }
    
    public bool IsLoading
    {
        get => _isLoading;
        set => this.RaiseAndSetIfChanged(ref _isLoading, value);
    }

    public bool IsNotLoading => !IsLoading;

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> LoginCommand { get; }

    #endregion

    #region 私有方法

    private void ClearError()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            HasError = false;
            ShowError = false;
            ErrorMessage = string.Empty;
        }).ConfigureAwait(false);
    }

    private void SetError(string message)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            ErrorMessage = message;
            HasError = true;
            ShowError = true;
        }).ConfigureAwait(false);
    }

    #endregion

    #region 登录逻辑

    private async Task LoginAsync()
    {
        try
        {
            // 获取输入值
            var username = Username?.Trim() ?? string.Empty;
            var password = Password ?? string.Empty;

            // 验证输入
            if (string.IsNullOrWhiteSpace(username))
            {
                await SetErrorAsync("请输入用户名");
                return;
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                await SetErrorAsync("请输入密码");
                return;
            }

            // 设置加载状态
            await SetLoadingStateAsync(true);
            await ClearErrorAsync();

            // 执行登录验证（在后台线程）
            var result = await Task.Run(async () => 
                await _userService.ValidateUserDetailedAsync(username, password));

            // 处理结果（在UI线程）
            await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                switch (result)
                {
                    case LoginResult.Success:
                        await HandleSuccessfulLoginAsync(username);
                        break;
                    
                    case LoginResult.UserNotFound:
                        SetError("用户不存在，请检查用户名");
                        break;
                    
                    case LoginResult.InvalidPassword:
                        Password = string.Empty;
                        SetError("密码错误，请重新输入");
                        break;
                    
                    case LoginResult.Error:
                        SetError("登录失败，请稍后重试");
                        break;
                }
            });
        }
        catch (Exception ex)
        {
            await SetErrorAsync($"登录失败：{ex.Message}");
        }
        finally
        {
            await SetLoadingStateAsync(false);
        }
    }

    private async Task HandleSuccessfulLoginAsync(string username)
    {
        try
        {
            // 确保在UI线程
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                var mainWindow = new MainWindow(username);
                mainWindow.Show();
                _parentWindow.Close();
            });
        }
        catch (Exception ex)
        {
            await SetErrorAsync($"打开主窗口失败：{ex.Message}");
        }
    }

    #endregion

    #region 线程安全的UI更新方法

    private async Task SetLoadingStateAsync(bool isLoading)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            IsLoading = isLoading;
        });
    }

    private async Task ClearErrorAsync()
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            ClearError();
        });
    }

    private async Task SetErrorAsync(string message)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            SetError(message);
        });
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                LoginCommand?.Dispose();
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