using ReactiveUI;
using System;
using System.Reactive;
using System.Reactive.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ExcelToSQLite.Services;
using Avalonia.Controls;
using ExcelToSQLite.Helpers;
using Avalonia.Media;

namespace ExcelToSQLite.ViewModels;

public class ChangePasswordViewModel : ReactiveObject, IDisposable
{
    private readonly UserService _userService;
    private readonly string _username;
    private Window? _parentWindow;
    private bool _isDisposed;
    private IDisposable? _validationSubscription;

    // 密码字段
    private string _currentPassword = string.Empty;
    private string _newPassword = string.Empty;
    private string _confirmPassword = string.Empty;

    // UI状态
    private string _messageText = string.Empty;
    private string _messageIcon = string.Empty;
    private IBrush _messageColor = new SolidColorBrush(Colors.Red);
    private IBrush _messageBackground = new SolidColorBrush(Color.Parse("#FFF8E1"));
    private bool _hasMessage;
    private bool _isLoading;
    private bool _canSave;
    private bool _hasError;

    // 密码强度指示器
    private IBrush _strength1Color = new SolidColorBrush(Color.Parse("#E0E0E0"));
    private IBrush _strength2Color = new SolidColorBrush(Color.Parse("#E0E0E0"));
    private IBrush _strength3Color = new SolidColorBrush(Color.Parse("#E0E0E0"));
    private IBrush _strength4Color = new SolidColorBrush(Color.Parse("#E0E0E0"));
    private string _strengthText = "密码强度：";
    private int _passwordStrength;

    // 密码可见性
    private bool _showPassword;

    public ChangePasswordViewModel(string username) : this(username, new UserService())
    {
    }

    public ChangePasswordViewModel(string username, UserService userService)
    {
        _username = username ?? throw new ArgumentNullException(nameof(username));
        _userService = userService ?? throw new ArgumentNullException(nameof(userService));

        // 初始化命令
        SaveCommand = ReactiveCommand.CreateFromTask(ExecuteSaveAsync);
        CancelCommand = ReactiveCommand.Create(Cancel);

        // 设置验证订阅 - 监听所有相关属性变化
        SetupValidation();

        // 初始化时清除消息
        ClearMessage();
        
        // 初始更新按钮状态
        UpdateCanSave();
    }

    #region 属性

    public string CurrentPassword
    {
        get => _currentPassword;
        set 
        { 
            this.RaiseAndSetIfChanged(ref _currentPassword, value);
            UpdateCanSave();
            ClearMessageIfNoError();
        }
    }

    public string NewPassword
    {
        get => _newPassword;
        set 
        { 
            this.RaiseAndSetIfChanged(ref _newPassword, value);
            UpdatePasswordStrength();
            UpdateCanSave();
            ClearMessageIfNoError();
        }
    }

    public string ConfirmPassword
    {
        get => _confirmPassword;
        set 
        { 
            this.RaiseAndSetIfChanged(ref _confirmPassword, value);
            UpdateCanSave();
            ValidatePasswords();
            ClearMessageIfNoError();
        }
    }

    public string MessageText
    {
        get => _messageText;
        set => this.RaiseAndSetIfChanged(ref _messageText, value);
    }

    public string MessageIcon
    {
        get => _messageIcon;
        set => this.RaiseAndSetIfChanged(ref _messageIcon, value);
    }

    public IBrush MessageColor
    {
        get => _messageColor;
        set => this.RaiseAndSetIfChanged(ref _messageColor, value);
    }

    public IBrush MessageBackground
    {
        get => _messageBackground;
        set => this.RaiseAndSetIfChanged(ref _messageBackground, value);
    }

    public bool HasMessage
    {
        get => _hasMessage;
        set => this.RaiseAndSetIfChanged(ref _hasMessage, value);
    }

    public bool HasError
    {
        get => _hasError;
        set => this.RaiseAndSetIfChanged(ref _hasError, value);
    }

    public bool IsLoading
    {
        get => _isLoading;
        set 
        { 
            this.RaiseAndSetIfChanged(ref _isLoading, value);
            UpdateCanSave();
        }
    }

    public bool CanSave
    {
        get => _canSave;
        set => this.RaiseAndSetIfChanged(ref _canSave, value);
    }

    public bool ShowPassword
    {
        get => _showPassword;
        set => this.RaiseAndSetIfChanged(ref _showPassword, value);
    }

    public IBrush Strength1Color
    {
        get => _strength1Color;
        set => this.RaiseAndSetIfChanged(ref _strength1Color, value);
    }

    public IBrush Strength2Color
    {
        get => _strength2Color;
        set => this.RaiseAndSetIfChanged(ref _strength2Color, value);
    }

    public IBrush Strength3Color
    {
        get => _strength3Color;
        set => this.RaiseAndSetIfChanged(ref _strength3Color, value);
    }

    public IBrush Strength4Color
    {
        get => _strength4Color;
        set => this.RaiseAndSetIfChanged(ref _strength4Color, value);
    }

    public string StrengthText
    {
        get => _strengthText;
        set => this.RaiseAndSetIfChanged(ref _strengthText, value);
    }

    public int PasswordStrength
    {
        get => _passwordStrength;
        set => this.RaiseAndSetIfChanged(ref _passwordStrength, value);
    }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> SaveCommand { get; }
    public ReactiveCommand<Unit, Unit> CancelCommand { get; }

    #endregion

    #region 事件

    public event Action<bool>? CloseDialog;

    #endregion

    #region 公共方法

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    #endregion

    #region 私有方法

    private void SetupValidation()
    {
        _validationSubscription = this.WhenAnyValue(
            x => x.CurrentPassword,
            x => x.NewPassword,
            x => x.ConfirmPassword,
            x => x.IsLoading
        ).Subscribe(_ => UpdateCanSave());
    }

    private void ValidatePasswords()
    {
        if (!string.IsNullOrEmpty(NewPassword) && !string.IsNullOrEmpty(ConfirmPassword))
        {
            if (NewPassword != ConfirmPassword)
            {
                ShowMessage("⚠️", "两次输入的密码不一致", Colors.Red, Color.Parse("#FFF8E1"));
            }
            else if (HasMessage && MessageText == "两次输入的密码不一致")
            {
                ClearMessage();
            }
        }
    }

    private void ClearMessageIfNoError()
    {
        if (HasMessage && !HasError)
        {
            ClearMessage();
        }
    }

    private void UpdateCanSave()
    {
        CanSave = !IsLoading &&
                  !string.IsNullOrWhiteSpace(CurrentPassword) &&
                  !string.IsNullOrWhiteSpace(NewPassword) &&
                  NewPassword.Length >= 6 &&
                  !string.IsNullOrWhiteSpace(ConfirmPassword) &&
                  NewPassword == ConfirmPassword;
    }

    private void ClearMessage()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            HasMessage = false;
            HasError = false;
            MessageText = string.Empty;
            MessageIcon = string.Empty;
        }).ConfigureAwait(false);
    }

    private void ShowMessage(string icon, string text, IBrush color, IBrush background)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            MessageIcon = icon;
            MessageText = text;
            MessageColor = color;
            MessageBackground = background;
            HasMessage = true;
            HasError = color == new SolidColorBrush(Colors.Red);
        }).ConfigureAwait(false);
    }

    private void ShowMessage(string icon, string text, Color color, Color background)
    {
        ShowMessage(icon, text, new SolidColorBrush(color), new SolidColorBrush(background));
    }

    private void ShowSuccessMessage(string icon, string text)
    {
        ShowMessage(icon, text, Colors.Green, Color.Parse("#E8F5E9"));
    }

    private void ShowErrorMessage(string icon, string text)
    {
        ShowMessage(icon, text, Colors.Red, Color.Parse("#FFF8E1"));
    }

    private async Task ExecuteSaveAsync()
    {
        try
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = true;
                ClearMessage();
            });

            // 验证密码
            if (string.IsNullOrWhiteSpace(CurrentPassword))
            {
                ShowErrorMessage("⚠️", "请输入当前密码");
                return;
            }

            if (string.IsNullOrWhiteSpace(NewPassword))
            {
                ShowErrorMessage("⚠️", "请输入新密码");
                return;
            }

            if (NewPassword.Length < 6)
            {
                ShowErrorMessage("⚠️", "新密码长度至少为6位");
                return;
            }

            if (NewPassword != ConfirmPassword)
            {
                ShowErrorMessage("⚠️", "两次输入的密码不一致");
                return;
            }

            // 执行密码修改（在后台线程执行）
            var result = await Task.Run(() => 
                _userService.ChangePasswordAsync(_username, CurrentPassword, NewPassword));
            
            if (result)
            {
                ShowSuccessMessage("✅", "密码修改成功！");
                await Task.Delay(800);
                
                // 关闭对话框（确保在UI线程）
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    CloseDialog?.Invoke(true);
                });
            }
            else
            {
                ShowErrorMessage("❌", "当前密码错误，请重新输入");
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    CurrentPassword = string.Empty;
                    UpdateCanSave();
                });
            }
        }
        catch (Exception ex)
        {
            ShowErrorMessage("❌", $"修改密码失败：{ex.Message}");
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsLoading = false;
            });
        }
    }

    private void Cancel()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            CloseDialog?.Invoke(false);
        }).ConfigureAwait(false);
    }

    private void UpdatePasswordStrength()
    {
        var password = NewPassword;
        var strength = CalculatePasswordStrength(password);
        PasswordStrength = strength;
        
        // 更新强度指示器颜色
        var colors = new[]
        {
            new SolidColorBrush(Color.Parse("#E0E0E0")),
            new SolidColorBrush(Color.Parse("#E0E0E0")),
            new SolidColorBrush(Color.Parse("#E0E0E0")),
            new SolidColorBrush(Color.Parse("#E0E0E0"))
        };
        
        var colorMap = new[]
        {
            new SolidColorBrush(Color.Parse("#EF5350")), // 红色 - 弱
            new SolidColorBrush(Color.Parse("#FFA726")), // 橙色 - 一般
            new SolidColorBrush(Color.Parse("#66BB6A")), // 绿色 - 强
            new SolidColorBrush(Color.Parse("#26A69A"))  // 青色 - 非常强
        };
        
        var strengthLevels = new[] { "密码强度：", "弱", "一般", "强", "非常强" };

        if (password.Length >= 6)
        {
            var strengthIndex = Math.Max(0, Math.Min(strength - 1, 3));
            for (int i = 0; i <= strengthIndex; i++)
            {
                colors[i] = colorMap[i];
            }
            
            StrengthText = $"密码强度：{strengthLevels[Math.Min(strength, 4)]}";
        }
        else if (password.Length > 0)
        {
            StrengthText = "密码至少需要6位";
            colors[0] = new SolidColorBrush(Color.Parse("#EF5350"));
        }
        else
        {
            StrengthText = "密码强度：";
            colors = new[]
            {
                new SolidColorBrush(Color.Parse("#E0E0E0")),
                new SolidColorBrush(Color.Parse("#E0E0E0")),
                new SolidColorBrush(Color.Parse("#E0E0E0")),
                new SolidColorBrush(Color.Parse("#E0E0E0"))
            };
        }

        Strength1Color = colors[0];
        Strength2Color = colors[1];
        Strength3Color = colors[2];
        Strength4Color = colors[3];
    }

    private int CalculatePasswordStrength(string password)
    {
        if (string.IsNullOrEmpty(password))
            return 0;

        int score = 0;

        // 长度评分
        if (password.Length >= 6) score++;
        if (password.Length >= 10) score++;
        if (password.Length >= 14) score++;

        // 复杂度评分
        bool hasDigit = Regex.IsMatch(password, @"\d");
        bool hasLower = Regex.IsMatch(password, @"[a-z]");
        bool hasUpper = Regex.IsMatch(password, @"[A-Z]");
        bool hasSpecial = Regex.IsMatch(password, @"[!@#$%^&*(),.?""':{}|<>]");

        if (hasDigit) score++;
        if (hasLower) score++;
        if (hasUpper) score++;
        if (hasSpecial) score++;

        // 计算最终强度等级 (1-4)
        if (score <= 2) return 1;      // 弱
        if (score <= 4) return 2;      // 一般
        if (score <= 6) return 3;      // 强
        return 4;                      // 非常强
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        if (!_isDisposed)
        {
            try
            {
                _validationSubscription?.Dispose();
                SaveCommand?.Dispose();
                CancelCommand?.Dispose();
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