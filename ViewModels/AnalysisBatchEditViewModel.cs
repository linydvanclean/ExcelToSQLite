using ReactiveUI;
using System;
using System.Linq;
using System.Reactive;
using System.Threading.Tasks;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ExcelToSQLite.Helpers;
using Avalonia.Media;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;

namespace ExcelToSQLite.ViewModels;

public class AnalysisBatchEditViewModel : ReactiveObject
{
    private int _id;
    private string _name = string.Empty;
    private DateTimeOffset _periodStart = DateTimeOffset.Now;
    private DateTimeOffset _periodEnd = DateTimeOffset.Now;
    private string _tablePrefix = string.Empty;
    private string _dialogTitle = "创建新批次";
    private string _dialogSubtitle = "填写批次信息，创建新的分析批次";
    private bool _isEditing = false;
    private bool _isDefaultBatch = false;
    private bool _isSaving = false;
    private string _statusMessage = string.Empty;
    private IBrush _statusColor = new SolidColorBrush(Colors.Green);
    private bool _showStatus = false;
    private Window? _parentWindow;

    public AnalysisBatchEditViewModel()
    {
        SaveCommand = ReactiveCommand.CreateFromTask(SaveAsync);
        CancelCommand = ReactiveCommand.Create(Cancel);
    }

    #region 属性

    public int Id
    {
        get => _id;
        set => this.RaiseAndSetIfChanged(ref _id, value);
    }

    public string Name
    {
        get => _name;
        set => this.RaiseAndSetIfChanged(ref _name, value);
    }

    public DateTimeOffset PeriodStart
    {
        get => _periodStart;
        set => this.RaiseAndSetIfChanged(ref _periodStart, value);
    }

    public DateTimeOffset PeriodEnd
    {
        get => _periodEnd;
        set => this.RaiseAndSetIfChanged(ref _periodEnd, value);
    }

    public string TablePrefix
    {
        get => _tablePrefix;
        set
        {
            var cleaned = string.Concat(value?.Where(c => char.IsLetterOrDigit(c) || c == '_') ?? "");
            this.RaiseAndSetIfChanged(ref _tablePrefix, cleaned);
        }
    }

    public string DialogTitle
    {
        get => _dialogTitle;
        set => this.RaiseAndSetIfChanged(ref _dialogTitle, value);
    }

    public string DialogSubtitle
    {
        get => _dialogSubtitle;
        set => this.RaiseAndSetIfChanged(ref _dialogSubtitle, value);
    }

    public bool IsEditing
    {
        get => _isEditing;
        set
        {
            this.RaiseAndSetIfChanged(ref _isEditing, value);
            UpdateDialogTitle();
        }
    }

    public bool IsDefaultBatch
    {
        get => _isDefaultBatch;
        set => this.RaiseAndSetIfChanged(ref _isDefaultBatch, value);
    }

    public bool IsSaving
    {
        get => _isSaving;
        set => this.RaiseAndSetIfChanged(ref _isSaving, value);
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

    #endregion

    #region 公共属性

    public AnalysisBatch? EditingBatch { get; private set; }
    public Func<Task<bool>>? OnSaveBatch { get; set; }
    public Action? OnCancel { get; set; }

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> SaveCommand { get; }
    public ReactiveCommand<Unit, Unit> CancelCommand { get; }

    #endregion

    #region 公共方法

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    public void LoadBatch(AnalysisBatch batch)
    {
        if (batch == null)
            throw new ArgumentNullException(nameof(batch));

        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            EditingBatch = batch;
            Id = batch.Id;
            Name = batch.Name;
            PeriodStart = new DateTimeOffset(batch.PeriodStart);
            PeriodEnd = new DateTimeOffset(batch.PeriodEnd);
            TablePrefix = batch.TablePrefix ?? string.Empty;
            IsEditing = true;
            IsDefaultBatch = AnalysisBatchService.IsDefaultBatch(batch.Id);
            StatusMessage = "已加载批次数据";
            StatusColor = new SolidColorBrush(Colors.Green);
            ShowStatus = true;
        }).ConfigureAwait(false);
    }

    public void Reset()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            Id = 0;
            Name = string.Empty;
            PeriodStart = DateTimeOffset.Now;
            PeriodEnd = DateTimeOffset.Now;
            TablePrefix = string.Empty;
            IsEditing = false;
            IsDefaultBatch = false;
            EditingBatch = null;
            IsSaving = false;
            StatusMessage = "已重置表单";
            StatusColor = new SolidColorBrush(Colors.Gray);
            ShowStatus = true;
        }).ConfigureAwait(false);
    }

    #endregion

    #region 私有方法

    private void UpdateDialogTitle()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            if (_isEditing)
            {
                DialogTitle = "编辑批次";
                DialogSubtitle = "修改批次信息并保存";
            }
            else
            {
                DialogTitle = "创建新批次";
                DialogSubtitle = "填写批次信息，创建新的分析批次";
            }
        }).ConfigureAwait(false);
    }

    private Window? GetWindow()
    {
        if (_parentWindow != null)
            return _parentWindow;

        if (App.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            return desktop.MainWindow;
        }
        return null;
    }

    private async Task SaveAsync()
    {
        if (IsSaving) return;

        try
        {
            // ✅ UI 状态更新 - 切换到 UI 线程
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsSaving = true;
                ShowStatus = true;
                StatusMessage = "正在保存...";
                StatusColor = new SolidColorBrush(Colors.Orange);
            });

            // 验证数据
            if (string.IsNullOrWhiteSpace(Name))
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    StatusMessage = "❌ 请输入批次名称";
                    StatusColor = new SolidColorBrush(Colors.Red);
                    IsSaving = false;
                });
                return;
            }

            if (PeriodStart > PeriodEnd)
            {
                await ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    StatusMessage = "❌ 开始日期不能晚于结束日期";
                    StatusColor = new SolidColorBrush(Colors.Red);
                    IsSaving = false;
                });
                return;
            }

            if (OnSaveBatch != null)
            {
                try
                {
                    // ✅ 执行保存操作
                    var result = await OnSaveBatch.Invoke();
                    
                    // ✅ 根据结果更新 UI
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        if (result)
                        {
                            StatusMessage = "✅ 保存成功！";
                            StatusColor = new SolidColorBrush(Colors.Green);
                            IsSaving = false;
                            
                            // ✅ 延迟关闭窗口
                            Task.Delay(300).ContinueWith(_ =>
                            {
                                ThreadingHelper.RunOnUIThreadAsync(() =>
                                {
                                    // ✅ 通过 OnCancel 或直接关闭
                                    OnCancel?.Invoke();
                                    var window = GetWindow();
                                    window?.Close();
                                }).ConfigureAwait(false);
                            });
                        }
                        else
                        {
                            StatusMessage = "❌ 保存失败，请重试";
                            StatusColor = new SolidColorBrush(Colors.Red);
                            IsSaving = false;
                        }
                    });
                }
                catch (Exception ex)
                {
                    // ✅ 捕获 OnSaveBatch 中的异常
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        StatusMessage = $"❌ 保存失败: {ex.Message}";
                        StatusColor = new SolidColorBrush(Colors.Red);
                        IsSaving = false;
                        ShowStatus = true;
                    });
                }
            }
        }
        catch (Exception ex)
        {
            // ✅ 捕获所有异常
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                StatusMessage = $"❌ 系统错误: {ex.Message}";
                StatusColor = new SolidColorBrush(Colors.Red);
                IsSaving = false;
                ShowStatus = true;
            });
        }
    }

    private void Cancel()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            try
            {
                OnCancel?.Invoke();
                var window = GetWindow();
                window?.Close();
            }
            catch
            {
            }
        }).ConfigureAwait(false);
    }

    #endregion
}