using Avalonia.Controls;
using Avalonia.Media;
using ExcelToSQLite.Models;
using ExcelToSQLite.Services;
using ReactiveUI;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Threading.Tasks;
using ExcelToSQLite.Views;
using ExcelToSQLite.Helpers;
using System.IO;
using System.Text.Json;
using System.Collections.Generic;

namespace ExcelToSQLite.ViewModels;

public class DataDictionaryViewModel : BaseEditableViewModel<DataDictionary>
{
    private readonly DataDictionaryService _dictionaryService = null!;
    private ObservableCollection<DataDictionary> _dictionaries = new();
    private bool _isImporting = false;
    private bool _isExporting = false;
    private Window? _parentWindow;
    private bool _isInitialized = false;
    private bool _isCleaned = false;

    public DataDictionaryViewModel()
    {
        try
        {
            _dictionaryService = new DataDictionaryService();

            CreateCommand = ReactiveCommand.CreateFromTask(CreateDictionaryAsync);
            EditCommand = ReactiveCommand.CreateFromTask<DataDictionary>(EditDictionaryAsync);
            DeleteCommand = ReactiveCommand.CreateFromTask<DataDictionary>(DeleteDictionaryAsync);
            RefreshCommand = ReactiveCommand.CreateFromTask(RefreshAsync);
            ImportCommand = ReactiveCommand.CreateFromTask(ImportAsync);
            ExportAllCommand = ReactiveCommand.CreateFromTask(ExportAllAsync);
            ExportSelectedCommand = ReactiveCommand.CreateFromTask(ExportSelectedAsync);
            SelectAllCommand = ReactiveCommand.Create(SelectAll);
            ClearSelectionCommand = ReactiveCommand.Create(ClearSelection);

            // 异步加载数据
            _ = LoadDataSafeAsync();

        }
        catch (Exception ex)
        {
            SetStatusSafe($"初始化失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #region 公共方法

    public void SetParentWindow(Window window)
    {
        _parentWindow = window;
    }

    #endregion

    #region 属性

    public ObservableCollection<DataDictionary> Dictionaries
    {
        get => _dictionaries;
        set => this.RaiseAndSetIfChanged(ref _dictionaries, value);
    }

    public bool IsImporting
    {
        get => _isImporting;
        set => this.RaiseAndSetIfChanged(ref _isImporting, value);
    }

    public bool IsExporting
    {
        get => _isExporting;
        set => this.RaiseAndSetIfChanged(ref _isExporting, value);
    }

    // ← 添加 new 关键字
    public new bool IsBusy => IsLoading || IsSaving || IsImporting || IsExporting;

    public bool HasSelectedDictionaries => Dictionaries.Any(d => d.IsSelected);

    public bool IsAllSelected
    {
        get => Dictionaries.All(d => d.IsSelected) && Dictionaries.Count > 0;
        set
        {
            if (value)
                SelectAll();
            else
                ClearSelection();
        }
    }

    public ObservableCollection<DataDictionary> SelectedDictionaries => 
        new ObservableCollection<DataDictionary>(Dictionaries.Where(d => d.IsSelected));

    #endregion

    #region 命令

    public ReactiveCommand<Unit, Unit> CreateCommand { get; } = null!;
    public ReactiveCommand<DataDictionary, Unit> EditCommand { get; } = null!;
    public ReactiveCommand<DataDictionary, Unit> DeleteCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> ImportCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> ExportAllCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> ExportSelectedCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> SelectAllCommand { get; } = null!;
    public ReactiveCommand<Unit, Unit> ClearSelectionCommand { get; } = null!;

    #endregion

    #region 数据加载

    private async Task LoadDataSafeAsync()
    {
        try
        {
            await LoadDataAsync();
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"加载数据失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    protected override async Task LoadDataAsync()
    {
        try
        {
            await SetLoadingAsync(true, "正在加载数据字典...");

            // 只在第一次加载时初始化
            if (!_isInitialized)
            {
                await _dictionaryService.InitializeAsync();
                _isInitialized = true;
            }
            else
            {
            }

            var list = await _dictionaryService.GetAllAsync();

            if (list == null)
            {
                list = new List<DataDictionary>();
            }

            var sortedList = list.OrderByDescending(d => d.CreatedAt).ToList();

            for (int i = 0; i < sortedList.Count; i++)
            {
                sortedList[i].Index = i + 1;
                sortedList[i].IsSelected = false;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Dictionaries = new ObservableCollection<DataDictionary>(sortedList);
            });
            

            await SetStatusSafeAsync($"✅ 加载完成，共 {Dictionaries.Count} 个数据字典", new SolidColorBrush(Colors.Green));

        }
        catch (Exception ex)
        {
            if (ex.InnerException != null)
            {
            }
            await SetStatusSafeAsync($"❌ 加载失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await SetLoadingAsync(false);
        }
    }

    #endregion

    #region CRUD操作

    private async Task CreateDictionaryAsync()
    {
        try
        {
            if (_parentWindow == null)
            {
                await SetStatusSafeAsync("无法打开对话框: 父窗口未设置", new SolidColorBrush(Colors.Red));
                return;
            }

            var editViewModel = new DataDictionaryEditViewModel();
            editViewModel.SetDialogWindow(_parentWindow);
            
            var dialog = new DataDictionaryEditDialog();
            dialog.Icon = IconHelper.GetAppIcon();
            dialog.DataContext = editViewModel;
            
            editViewModel.OnSave = async (dictionary) =>
            {
                try
                {
                    await _dictionaryService.AddAsync(dictionary);
                    
                    // ✅ 刷新数据
                    await LoadDataAsync();
                    await SetStatusSafeAsync($"✅ 数据字典 '{dictionary.Name}' 创建成功", new SolidColorBrush(Colors.Green));
                    
                    // ✅ 保存成功后关闭窗口（在 UI 线程）
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        editViewModel.OnClose?.Invoke(true);
                    });
                    
                    return true;
                }
                catch (Exception ex)
                {
                    await SetStatusSafeAsync($"❌ 创建失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                    return false;
                }
            };
            
            // ✅ OnClose 使用 ThreadingHelper 确保在 UI 线程关闭窗口
            editViewModel.OnClose = (isSaved) =>
            {
                
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    try
                    {
                        dialog.Close();
                    }
                    catch
                    {
                    }
                }).ConfigureAwait(false);
                
                // ✅ 只有取消操作时才显示"已取消操作"
                if (!isSaved)
                {
                    SetStatusSafeAsync("已取消操作", new SolidColorBrush(Color.Parse("#78909C"))).ConfigureAwait(false);
                }
            };

            await dialog.ShowDialog(_parentWindow);
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 打开对话框失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private async Task EditDictionaryAsync(DataDictionary dictionary)
    {
        try
        {
            if (dictionary == null)
            {
                await SetStatusSafeAsync("请选择要编辑的数据字典", new SolidColorBrush(Colors.Orange));
                return;
            }

            if (_parentWindow == null)
            {
                await SetStatusSafeAsync("无法打开对话框: 父窗口未设置", new SolidColorBrush(Colors.Red));
                return;
            }

            
            var editViewModel = new DataDictionaryEditViewModel();
            editViewModel.SetDialogWindow(_parentWindow);
            editViewModel.LoadDictionary(dictionary);
            
            var dialog = new DataDictionaryEditDialog();
            dialog.Icon = IconHelper.GetAppIcon();
            dialog.DataContext = editViewModel;
            
            editViewModel.OnSave = async (updatedDict) =>
            {
                try
                {
                    await _dictionaryService.UpdateAsync(updatedDict);
                    
                    await LoadDataAsync();
                    await SetStatusSafeAsync($"✅ 数据字典 '{updatedDict.Name}' 更新成功", new SolidColorBrush(Colors.Green));
                    
                    // ✅ 保存成功后关闭窗口（在 UI 线程）
                    await ThreadingHelper.RunOnUIThreadAsync(() =>
                    {
                        editViewModel.OnClose?.Invoke(true);
                    });
                    
                    return true;
                }
                catch (Exception ex)
                {
                    await SetStatusSafeAsync($"❌ 更新失败: {ex.Message}", new SolidColorBrush(Colors.Red));
                    return false;
                }
            };
            
            // ✅ OnClose 使用 ThreadingHelper 确保在 UI 线程关闭窗口
            editViewModel.OnClose = (isSaved) =>
            {
                
                ThreadingHelper.RunOnUIThreadAsync(() =>
                {
                    try
                    {
                        dialog.Close();
                    }
                    catch
                    {
                    }
                }).ConfigureAwait(false);
                
                if (!isSaved)
                {
                    SetStatusSafeAsync("已取消操作", new SolidColorBrush(Color.Parse("#78909C"))).ConfigureAwait(false);
                }
            };

            await dialog.ShowDialog(_parentWindow);
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 编辑失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    private async Task DeleteDictionaryAsync(DataDictionary dictionary)
    {
        try
        {
            if (dictionary == null)
            {
                await SetStatusSafeAsync("请选择要删除的数据字典", new SolidColorBrush(Colors.Orange));
                return;
            }


            var confirmResult = await ShowConfirmDialogAsync(
                $"确认删除\n\n数据字典: {dictionary.Name}\n表名: {dictionary.TableName ?? "(无)"}\n\n删除后无法恢复，确认删除？",
                "确认删除"
            );

            if (!confirmResult) return;

            await _dictionaryService.DeleteAsync(dictionary.Id);
            await LoadDataAsync();
            await SetStatusSafeAsync($"✅ 数据字典 '{dictionary.Name}' 删除成功", new SolidColorBrush(Colors.Green));
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 删除失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
    }

    #endregion

    #region 导入导出

    private async Task ImportAsync()
    {
        try
        {
            if (_parentWindow == null)
            {
                await SetStatusSafeAsync("无法导入: 父窗口未设置", new SolidColorBrush(Colors.Red));
                return;
            }

            await SetStatusSafeAsync("正在导入数据字典...", new SolidColorBrush(Colors.Orange));
            
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsImporting = true;
            });

            var filePath = await FileDialogHelper.OpenJsonFileAsync(_parentWindow, "选择数据字典 JSON 文件");
            
            if (string.IsNullOrEmpty(filePath))
            {
                await SetStatusSafeAsync("已取消导入", new SolidColorBrush(Color.Parse("#78909C")));
                return;
            }

            
            var json = await File.ReadAllTextAsync(filePath);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
            
            var dictionaries = JsonSerializer.Deserialize<List<DataDictionary>>(json, options);
            
            if (dictionaries == null || dictionaries.Count == 0)
            {
                await SetStatusSafeAsync("❌ 导入失败: 没有有效的数据", new SolidColorBrush(Colors.Red));
                return;
            }

            int importCount = 0;
            foreach (var dict in dictionaries)
            {
                try
                {
                    var existing = Dictionaries.FirstOrDefault(d => d.Name == dict.Name);
                    if (existing == null)
                    {
                        dict.Id = Guid.NewGuid().ToString();
                        dict.CreatedAt = DateTime.Now;
                        dict.UpdatedAt = DateTime.Now;
                        dict.IsActive = true;
                        dict.CreatedBy = dict.CreatedBy ?? "admin";
                        
                        await _dictionaryService.AddAsync(dict);
                        importCount++;
                    }
                    else
                    {
                        existing.TableName = dict.TableName;
                        existing.Description = dict.Description;
                        existing.UpdatedAt = DateTime.Now;
                        await _dictionaryService.UpdateAsync(existing);
                        importCount++;
                    }
                }
                catch
                {
                }
            }
            
            if (importCount > 0)
            {
                await LoadDataAsync();
                await SetStatusSafeAsync($"✅ 成功导入 {importCount} 个数据字典", new SolidColorBrush(Colors.Green));
            }
            else
            {
                await SetStatusSafeAsync("❌ 导入失败: 没有导入任何数据", new SolidColorBrush(Colors.Red));
            }
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 导入失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsImporting = false;
            });
        }
    }

    private async Task ExportAllAsync()
    {
        try
        {
            if (_parentWindow == null)
            {
                await SetStatusSafeAsync("无法导出: 父窗口未设置", new SolidColorBrush(Colors.Red));
                return;
            }

            if (Dictionaries.Count == 0)
            {
                await SetStatusSafeAsync("没有数据字典可导出", new SolidColorBrush(Colors.Orange));
                return;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = true;
            });

            await SetStatusSafeAsync("正在导出数据字典...", new SolidColorBrush(Colors.Orange));

            var result = await ExportDictionariesAsync(Dictionaries.ToList(), "所有数据字典");
            
            if (result)
            {
                await SetStatusSafeAsync($"✅ 成功导出 {Dictionaries.Count} 个数据字典", new SolidColorBrush(Colors.Green));
            }
            else
            {
                await SetStatusSafeAsync("❌ 导出失败", new SolidColorBrush(Colors.Red));
            }
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 导出失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = false;
            });
        }
    }

    private async Task ExportSelectedAsync()
    {
        try
        {
            if (_parentWindow == null)
            {
                await SetStatusSafeAsync("无法导出: 父窗口未设置", new SolidColorBrush(Colors.Red));
                return;
            }

            var selected = Dictionaries.Where(d => d.IsSelected).ToList();
            if (selected.Count == 0)
            {
                await SetStatusSafeAsync("请先选择要导出的数据字典", new SolidColorBrush(Colors.Orange));
                return;
            }

            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = true;
            });

            await SetStatusSafeAsync($"正在导出 {selected.Count} 个数据字典...", new SolidColorBrush(Colors.Orange));

            var result = await ExportDictionariesAsync(selected, $"选中的数据字典_{selected.Count}个");
            
            if (result)
            {
                await SetStatusSafeAsync($"✅ 成功导出 {selected.Count} 个数据字典", new SolidColorBrush(Colors.Green));
            }
            else
            {
                await SetStatusSafeAsync("❌ 导出失败", new SolidColorBrush(Colors.Red));
            }
        }
        catch (Exception ex)
        {
            await SetStatusSafeAsync($"❌ 导出失败: {ex.Message}", new SolidColorBrush(Colors.Red));
        }
        finally
        {
            await ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                IsExporting = false;
            });
        }
    }

    private async Task<bool> ExportDictionariesAsync(List<DataDictionary> dictionaries, string fileNamePrefix)
    {
        try
        {
            if (dictionaries == null || dictionaries.Count == 0)
                return false;

            var exportData = new List<object>();
            foreach (var dict in dictionaries)
            {
                exportData.Add(new
                {
                    dict.Id,
                    dict.Name,
                    dict.TableName,
                    dict.Description,
                    dict.CreatedAt,
                    dict.UpdatedAt,
                    dict.CreatedBy,
                    dict.IsActive
                });
            }

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            var json = JsonSerializer.Serialize(exportData, options);

            var fileName = $"{fileNamePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
            
            var savePath = await FileDialogHelper.SaveFileAsync(
                _parentWindow!,
                json,
                "导出数据字典",
                "json",
                "JSON 文件"
            );

            return !string.IsNullOrEmpty(savePath);
        }
        catch
        {
            throw;
        }
    }

    #endregion

    #region 选择操作

    private void SelectAll()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            try
            {
                foreach (var item in Dictionaries)
                {
                    item.IsSelected = true;
                }
                this.RaisePropertyChanged(nameof(HasSelectedDictionaries));
                this.RaisePropertyChanged(nameof(IsAllSelected));
                SetStatusSafeAsync($"已选择 {Dictionaries.Count} 个数据字典", new SolidColorBrush(Color.Parse("#1565C0")))
                    .ConfigureAwait(false);
            }
            catch
            {
            }
        }).ConfigureAwait(false);
    }

    private void ClearSelection()
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            try
            {
                foreach (var item in Dictionaries)
                {
                    item.IsSelected = false;
                }
                this.RaisePropertyChanged(nameof(HasSelectedDictionaries));
                this.RaisePropertyChanged(nameof(IsAllSelected));
                SetStatusSafeAsync("已清空选择", new SolidColorBrush(Color.Parse("#78909C")))
                    .ConfigureAwait(false);
            }
            catch
            {
            }
        }).ConfigureAwait(false);
    }

    #endregion

    #region UI辅助方法

    private void SetStatusSafe(string message, IBrush color)
    {
        ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = color;
            ShowStatus = true;
        }).ConfigureAwait(false);
    }

    private async Task SetStatusSafeAsync(string message, IBrush color)
    {
        await ThreadingHelper.RunOnUIThreadAsync(() =>
        {
            StatusMessage = message;
            StatusColor = color;
            ShowStatus = true;
        });
    }

    private async Task<bool> ShowConfirmDialogAsync(string message, string title = "确认")
    {
        try
        {
            if (_parentWindow == null) return false;

            return await ThreadingHelper.RunOnUIThreadAsync(async () =>
            {
                var result = await MessageBox.ShowAsync(
                    _parentWindow,
                    message,
                    title,
                    MessageBoxButtons.YesNo
                );

                return result == MessageBoxResult.Yes;
            });
        }
        catch
        {
            return false;
        }
    }

    #endregion

    #region 清理

    public new void Cleanup()
    {
        if (_isCleaned) return;

        try
        {
            
            ThreadingHelper.RunOnUIThreadAsync(() =>
            {
                Dictionaries.Clear();
            }).ConfigureAwait(false);
            
            _isCleaned = true;
        }
        catch
        {
        }
    }

    #endregion
}