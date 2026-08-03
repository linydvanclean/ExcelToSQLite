using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using System;
using System.IO;

namespace ExcelToSQLite.Helpers;

public static class IconHelper
{
    private static WindowIcon? _cachedIcon;

    public static WindowIcon GetAppIcon()
    {
        if (_cachedIcon != null)
            return _cachedIcon;

        try
        {
            // 创建一个简单的柱状图图标
            var bitmap = new RenderTargetBitmap(
                new PixelSize(32, 32),
                new Vector(96, 96));

            using (var context = bitmap.CreateDrawingContext())
            {
                // 背景透明
                context.FillRectangle(
                    new SolidColorBrush(Colors.Transparent),
                    new Rect(0, 0, 32, 32));
                
                // 绘制四个柱子
                var data = new[] { 
                    new { Height = 16, Color = Color.FromRgb(255, 50, 50), X = 4 },
                    new { Height = 22, Color = Color.FromRgb(255, 165, 0), X = 10 },
                    new { Height = 12, Color = Color.FromRgb(50, 205, 50), X = 16 },
                    new { Height = 18, Color = Color.FromRgb(66, 133, 244), X = 22 }
                };
                
                foreach (var item in data)
                {
                    context.FillRectangle(
                        new SolidColorBrush(item.Color),
                        new Rect(item.X, 32 - item.Height, 4, item.Height));
                }
            }

            using var ms = new MemoryStream();
            // 使用 PngBitmapEncoderOptions 消除警告
            bitmap.Save(ms, new PngBitmapEncoderOptions());
            ms.Position = 0;
            
            _cachedIcon = new WindowIcon(ms);
            return _cachedIcon;
        }
        catch
        {
            return CreateDefaultIcon();
        }
    }

    private static WindowIcon CreateDefaultIcon()
    {
        var bitmap = new RenderTargetBitmap(
            new PixelSize(32, 32),
            new Vector(96, 96));

        using (var context = bitmap.CreateDrawingContext())
        {
            context.FillRectangle(
                new SolidColorBrush(Colors.Blue),
                new Rect(0, 0, 32, 32));
        }

        using var ms = new MemoryStream();
        bitmap.Save(ms, new PngBitmapEncoderOptions());
        ms.Position = 0;
        
        return new WindowIcon(ms);
    }

    public static void SetAppIcon(this Window window)
    {
        if (window == null) return;
        window.Icon = GetAppIcon();
    }
}