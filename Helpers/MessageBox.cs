using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls.Primitives;
using Avalonia.Media;
using Avalonia.Layout;
using Avalonia.Input;

namespace ExcelToSQLite.Helpers;

public enum MessageBoxResult
{
    Yes,
    No,
    Cancel,
    OK
}

public enum MessageBoxButtons
{
    OK,
    OKCancel,
    YesNo,
    YesNoCancel
}

public enum MessageBoxIcon
{
    None,
    Information,
    Success,
    Warning,
    Error,
    Question
}

public class MessageBox : Window
{
    private MessageBoxResult _result = MessageBoxResult.Cancel;
    private string? _messageBoxContent;
    private MessageBoxIcon _messageBoxIcon = MessageBoxIcon.None;

    public MessageBox()
    {
        Width = 460;
        MinHeight = 130;
        MaxHeight = 500;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        CanResize = false;
        Background = new SolidColorBrush(Color.Parse("#FAFAFA"));
        SizeToContent = SizeToContent.Height;
        
        // 设置窗口图标为应用程序图标（与 IconHelper 保持一致）
        try
        {
            base.Icon = IconHelper.GetAppIcon();
        }
        catch
        {
            // 如果获取图标失败，忽略
        }
    }

    public string? MessageBoxContent
    {
        get => _messageBoxContent;
        set => _messageBoxContent = value;
    }

    public MessageBoxButtons Buttons { get; set; } = MessageBoxButtons.OK;
    
    // 重命名属性以避免与 Window.Icon 冲突
    public MessageBoxIcon MessageBoxIconType
    {
        get => _messageBoxIcon;
        set => _messageBoxIcon = value;
    }

    // 静态方法：自动获取主窗口
    public static async Task<MessageBoxResult> ShowAsync(
        string content, 
        string title = "提示", 
        MessageBoxButtons buttons = MessageBoxButtons.OK,
        MessageBoxIcon icon = MessageBoxIcon.None)
    {
        return await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            var owner = GetMainWindow();
            if (owner == null || !owner.IsVisible)
            {
                owner = new Window
                {
                    Width = 1,
                    Height = 1,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    Opacity = 0,
                    ShowInTaskbar = false,
                    CanResize = false
                };
                owner.Show();
            }

            var dialog = new MessageBox
            {
                Title = title,
                MessageBoxContent = content,
                Buttons = buttons,
                MessageBoxIconType = icon
            };

            return await dialog.ShowDialogAsync(owner);
        });
    }

    // 静态方法：指定所有者窗口
    public static async Task<MessageBoxResult> ShowAsync(
        Window owner,
        string content, 
        string title = "提示", 
        MessageBoxButtons buttons = MessageBoxButtons.OK,
        MessageBoxIcon icon = MessageBoxIcon.None)
    {
        if (owner == null || !owner.IsVisible)
        {
            return await ShowAsync(content, title, buttons, icon);
        }

        return await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            var dialog = new MessageBox
            {
                Title = title,
                MessageBoxContent = content,
                Buttons = buttons,
                MessageBoxIconType = icon
            };

            return await dialog.ShowDialogAsync(owner);
        });
    }

    // 实例方法
    public async Task<MessageBoxResult> ShowDialogAsync(Window owner)
    {
        return await ThreadingHelper.RunOnUIThreadAsync(async () =>
        {
            // 检查 owner 是否有效
            if (owner == null || !owner.IsVisible)
            {
                var newOwner = GetMainWindow();
                if (newOwner != null && newOwner.IsVisible)
                {
                    owner = newOwner;
                }
                else
                {
                    owner = new Window
                    {
                        Width = 1,
                        Height = 1,
                        WindowStartupLocation = WindowStartupLocation.CenterScreen,
                        Opacity = 0,
                        ShowInTaskbar = false,
                        CanResize = false
                    };
                    owner.Show();
                }
            }

            // 构建主布局
            var mainPanel = new StackPanel
            {
                Margin = new Thickness(20, 16, 20, 16),
                Spacing = 8
            };

            // 1. 头部区域（包含图标）
            var headerPanel = CreateHeaderPanel();
            mainPanel.Children.Add(headerPanel);

            // 2. 内容区域（带滚动）
            var contentPanel = CreateContentPanel();
            mainPanel.Children.Add(contentPanel);

            // 3. 按钮区域
            var buttonPanel = CreateButtonPanel();
            mainPanel.Children.Add(buttonPanel);

            this.Content = mainPanel;

            try
            {
                await base.ShowDialog(owner);
            }
            catch
            {
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                await base.ShowDialog(null!);
            }

            return _result;
        });
    }

    private Border CreateHeaderPanel()
    {
        // 获取图标信息（包括 emoji 和颜色）
        var (iconText, color) = GetIconInfo();

        var headerPanel = new Border
        {
            Padding = new Thickness(0, 0, 0, 8),
            BorderThickness = new Thickness(0, 0, 0, 1),
            BorderBrush = new SolidColorBrush(Color.Parse("#E8E8E8"))
        };

        var headerStack = new StackPanel
        {
            Orientation = Orientation.Vertical,
            HorizontalAlignment = HorizontalAlignment.Center,
            Spacing = 4
        };

        // 如果有图标，显示图标
        if (!string.IsNullOrEmpty(iconText))
        {
            // 使用与 IconHelper 一致的图标显示
            var iconBorder = new Border
            {
                Width = 48,
                Height = 48,
                CornerRadius = new CornerRadius(24),
                Background = new SolidColorBrush(color)
            };

            var iconTextBlock = new TextBlock
            {
                Text = iconText,
                FontSize = 24,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0)
            };

            iconBorder.Child = iconTextBlock;
            headerStack.Children.Add(iconBorder);
        }

        // 标题
        headerStack.Children.Add(new TextBlock
        {
            Text = this.Title,
            FontSize = 16,
            FontWeight = FontWeight.SemiBold,
            Foreground = new SolidColorBrush(Color.Parse("#1A1A1A")),
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, iconText != null ? 4 : 0, 0, 0)
        });

        headerPanel.Child = headerStack;
        return headerPanel;
    }

    private Border CreateContentPanel()
    {
        // 内容文本
        var contentTextBlock = new TextBlock
        {
            Text = MessageBoxContent ?? string.Empty,
            TextWrapping = TextWrapping.Wrap,
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.Parse("#333333")),
            LineHeight = 20,
            TextAlignment = TextAlignment.Left,
            Padding = new Thickness(8, 4)
        };

        // 创建滚动容器 - 让 SizeToContent 自动处理高度
        var scrollViewer = new ScrollViewer
        {
            Content = contentTextBlock,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            MaxHeight = 300,
            Padding = new Thickness(0),
            Background = new SolidColorBrush(Color.Parse("#F8F9FA")),
            CornerRadius = new CornerRadius(4)
        };

        return new Border
        {
            Child = scrollViewer,
            Padding = new Thickness(0, 4, 0, 4)
        };
    }

    private StackPanel CreateButtonPanel()
    {
        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            Spacing = 10,
            Margin = new Thickness(0, 4, 0, 0)
        };

        if (Buttons == MessageBoxButtons.OK)
        {
            buttonPanel.Children.Add(CreateButton("确定", MessageBoxResult.OK, "#4CAF50"));
        }
        else if (Buttons == MessageBoxButtons.OKCancel)
        {
            buttonPanel.Children.Add(CreateButton("确定", MessageBoxResult.OK, "#4CAF50"));
            buttonPanel.Children.Add(CreateButton("取消", MessageBoxResult.Cancel, "#78909C"));
        }
        else if (Buttons == MessageBoxButtons.YesNo)
        {
            buttonPanel.Children.Add(CreateButton("是", MessageBoxResult.Yes, "#4CAF50"));
            buttonPanel.Children.Add(CreateButton("否", MessageBoxResult.No, "#EF5350"));
        }
        else if (Buttons == MessageBoxButtons.YesNoCancel)
        {
            buttonPanel.Children.Add(CreateButton("是", MessageBoxResult.Yes, "#4CAF50"));
            buttonPanel.Children.Add(CreateButton("否", MessageBoxResult.No, "#EF5350"));
            buttonPanel.Children.Add(CreateButton("取消", MessageBoxResult.Cancel, "#78909C"));
        }

        return buttonPanel;
    }

    private Button CreateButton(string text, MessageBoxResult result, string colorHex)
    {
        var baseColor = Color.Parse(colorHex);
        var hoverColor = baseColor.Lighten(0.15);
        var pressedColor = baseColor.Lighten(0.3);

        var button = new Button
        {
            Content = text,
            Width = 90,
            Height = 32,
            Background = new SolidColorBrush(baseColor),
            Foreground = Brushes.White,
            FontWeight = FontWeight.Medium,
            FontSize = 13,
            HorizontalAlignment = HorizontalAlignment.Center,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            Cursor = new Cursor(StandardCursorType.Hand),
            CornerRadius = new CornerRadius(4)
        };

        button.PointerEntered += (s, e) => 
        {
            button.Background = new SolidColorBrush(hoverColor);
        };
        
        button.PointerExited += (s, e) => 
        {
            button.Background = new SolidColorBrush(baseColor);
        };

        button.PointerPressed += (s, e) =>
        {
            button.Background = new SolidColorBrush(pressedColor);
        };

        button.PointerReleased += (s, e) =>
        {
            button.Background = new SolidColorBrush(hoverColor);
        };

        button.Click += (s, e) =>
        {
            _result = result;
            this.Close();
        };

        return button;
    }

    private (string Icon, Color Color) GetIconInfo()
    {
        return MessageBoxIconType switch
        {
            MessageBoxIcon.Success => ("✅", Color.Parse("#2E7D32")),
            MessageBoxIcon.Information => ("ℹ️", Color.Parse("#1565C0")),
            MessageBoxIcon.Warning => ("⚠️", Color.Parse("#E65100")),
            MessageBoxIcon.Error => ("❌", Color.Parse("#C62828")),
            MessageBoxIcon.Question => ("❓", Color.Parse("#1565C0")),
            _ => ("", Colors.Transparent)
        };
    }

    private static Window? GetMainWindow()
    {
        var desktop = App.Current?.ApplicationLifetime as IClassicDesktopStyleApplicationLifetime;
        if (desktop?.MainWindow != null && desktop.MainWindow.IsVisible)
        {
            return desktop.MainWindow;
        }

        if (desktop?.Windows != null)
        {
            foreach (var window in desktop.Windows)
            {
                if (window.IsVisible && window != desktop.MainWindow)
                {
                    return window;
                }
            }
        }

        return null;
    }
}

// 颜色扩展方法
public static class ColorExtensions
{
    public static Color Lighten(this Color color, double factor)
    {
        var r = (byte)Math.Min(255, color.R + (255 - color.R) * factor);
        var g = (byte)Math.Min(255, color.G + (255 - color.G) * factor);
        var b = (byte)Math.Min(255, color.B + (255 - color.B) * factor);
        return Color.FromRgb(r, g, b);
    }
}