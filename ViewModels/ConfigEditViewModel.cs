using ReactiveUI;
using System;
using System.IO;
using System.Reactive;
using System.Text.Json;
using System.Threading.Tasks;
using Avalonia.Media;
using ExcelToSQLite.Services;
using ExcelToSQLite.Helpers;
using Avalonia.Controls;
using Avalonia.Threading;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.ViewModels
{
    public class ConfigEditViewModel : ReactiveObject, IDisposable
    {
        private readonly AppConfigService _configService;
        private bool _isDisposed;
        
        private string _configJson = string.Empty;
        private string _statusMessage = string.Empty;
        private IBrush _statusColor = Brushes.Transparent;
        private bool _isLoading = false;
        private bool _isSaving = false;
        private bool _hasChanges = false;
        private string _configPath = string.Empty;
        private string _originalJson = string.Empty;
        private bool _showStatus = false;

        public ConfigEditViewModel()
        {
            _configService = new AppConfigService();
            
            LoadCommand = ReactiveCommand.CreateFromTask(LoadConfigAsync);
            SaveCommand = ReactiveCommand.CreateFromTask(SaveConfigAsync, this.WhenAnyValue(x => x.HasChanges));
            ResetCommand = ReactiveCommand.CreateFromTask(LoadConfigAsync);
            FormatCommand = ReactiveCommand.Create(FormatJson);
            CancelCommand = ReactiveCommand.Create(Cancel);
            
            // 异步加载配置
            _ = LoadConfigSafeAsync();
        }

        #region 属性

        public string ConfigJson
        {
            get => _configJson;
            set 
            {
                this.RaiseAndSetIfChanged(ref _configJson, value);
                CheckHasChanges();
            }
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

        public bool ShowStatus
        {
            get => _showStatus;
            set => this.RaiseAndSetIfChanged(ref _showStatus, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            set => this.RaiseAndSetIfChanged(ref _isLoading, value);
        }

        public bool IsSaving
        {
            get => _isSaving;
            set => this.RaiseAndSetIfChanged(ref _isSaving, value);
        }

        public bool IsBusy => IsLoading || IsSaving;

        public bool HasChanges
        {
            get => _hasChanges;
            set => this.RaiseAndSetIfChanged(ref _hasChanges, value);
        }

        public string ConfigPath
        {
            get => _configPath;
            set => this.RaiseAndSetIfChanged(ref _configPath, value);
        }

        #endregion

        #region 命令

        public ReactiveCommand<Unit, Unit> LoadCommand { get; }
        public ReactiveCommand<Unit, Unit> SaveCommand { get; }
        public ReactiveCommand<Unit, Unit> ResetCommand { get; }
        public ReactiveCommand<Unit, Unit> FormatCommand { get; }
        public ReactiveCommand<Unit, Unit> CancelCommand { get; }

        #endregion

        #region 私有方法

        private async Task LoadConfigSafeAsync()
        {
            try 
            { 
                await LoadConfigAsync(); 
            }
            catch (Exception ex) 
            { 
                
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    StatusMessage = $"❌ 加载失败: {ex.Message}";
                    StatusColor = Brushes.Red;
                    ShowStatus = true;
                });
            }
        }

        private void CheckHasChanges()
        {
            // ✅ 确保在UI线程执行比较
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.Post(CheckHasChanges);
                return;
            }

            bool hasChanges = _configJson != _originalJson;
            if (HasChanges != hasChanges)
            {
                HasChanges = hasChanges;
            }
        }

        private async Task LoadConfigAsync()
        {
            try
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    IsLoading = true;
                    ShowStatus = true;
                    StatusMessage = "正在加载配置...";
                    StatusColor = Brushes.Orange;
                });

                // 在后台线程加载配置
                var config = await Task.Run(() => _configService.GetConfigAsync());
                
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var json = JsonSerializer.Serialize(config, options);

                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    _originalJson = json;
                    ConfigJson = json;
                    ConfigPath = Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory,
                        "settings",
                        "appconfig.json"
                    );
                    HasChanges = false;
                    ShowStatus = true;
                    StatusMessage = "✅ 配置加载成功";
                    StatusColor = Brushes.Green;
                    IsLoading = false;
                });
            }
            catch (Exception ex)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    ShowStatus = true;
                    StatusMessage = $"❌ 加载失败: {ex.Message}";
                    StatusColor = Brushes.Red;
                    IsLoading = false;
                });
                
            }
        }

        private async Task SaveConfigAsync()
        {
            try
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    IsSaving = true;
                    ShowStatus = true;
                    StatusMessage = "正在保存配置...";
                    StatusColor = Brushes.Orange;
                });

                // 验证 JSON 格式
                try
                {
                    JsonDocument.Parse(ConfigJson);
                }
                catch (JsonException ex)
                {
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        ShowStatus = true;
                        StatusMessage = $"❌ JSON 格式错误: {ex.Message}";
                        StatusColor = Brushes.Red;
                        IsSaving = false;
                    });
                    return;
                }

                // 在后台线程反序列化和保存
                var config = await Task.Run(() => JsonSerializer.Deserialize<AppConfig>(ConfigJson));
                
                if (config == null)
                {
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        ShowStatus = true;
                        StatusMessage = "❌ 配置内容无效";
                        StatusColor = Brushes.Red;
                        IsSaving = false;
                    });
                    return;
                }

                await Task.Run(() => _configService.SaveConfigAsync(config));
                
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    _originalJson = ConfigJson;
                    HasChanges = false;
                    ShowStatus = true;
                    StatusMessage = "✅ 配置保存成功！";
                    StatusColor = Brushes.Green;
                    IsSaving = false;
                });
            }
            catch (Exception ex)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    ShowStatus = true;
                    StatusMessage = $"❌ 保存失败: {ex.Message}";
                    StatusColor = Brushes.Red;
                    IsSaving = false;
                });
                
            }
        }

        private void FormatJson()
        {
            // ✅ 确保在UI线程执行
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(ConfigJson))
                    {
                        ShowStatus = true;
                        StatusMessage = "⚠️ 配置内容为空，无法格式化";
                        StatusColor = Brushes.Orange;
                        return;
                    }

                    using var doc = JsonDocument.Parse(ConfigJson);
                    var options = new JsonSerializerOptions 
                    { 
                        WriteIndented = true,
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    };
                    var formattedJson = JsonSerializer.Serialize(doc.RootElement, options);
                    ConfigJson = formattedJson;
                    ShowStatus = true;
                    StatusMessage = "✅ JSON 格式化完成";
                    StatusColor = Brushes.Green;
                }
                catch (JsonException ex)
                {
                    ShowStatus = true;
                    StatusMessage = $"❌ JSON 格式错误: {ex.Message}";
                    StatusColor = Brushes.Red;
                }
                catch (Exception ex)
                {
                    ShowStatus = true;
                    StatusMessage = $"❌ 格式化失败: {ex.Message}";
                    StatusColor = Brushes.Red;
                    
                }
            }).ConfigureAwait(false);
        }

        private void Cancel()
        {
            // ✅ 确保在UI线程执行
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                try
                {
                    // 重置状态
                    ShowStatus = true;
                    StatusMessage = "已取消操作";
                    StatusColor = Brushes.Gray;
                    
                    // 如果有未保存的更改，恢复到原始状态
                    if (HasChanges)
                    {
                        ConfigJson = _originalJson;
                        HasChanges = false;
                    }
                }
                catch
                {
                }
            }).ConfigureAwait(false);
        }

        #endregion

        #region 资源清理

        public void Dispose()
        {
            if (!_isDisposed)
            {
                try
                {
                    LoadCommand?.Dispose();
                    SaveCommand?.Dispose();
                    ResetCommand?.Dispose();
                    FormatCommand?.Dispose();
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
}