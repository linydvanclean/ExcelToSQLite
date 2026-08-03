using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Media;
using ExcelToSQLite.Services;
using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reactive;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 动态数据行 - 支持 DataGrid 绑定
/// 线程安全的动态数据行实现
/// </summary>
public class DataGridRow : ReactiveObject, IDisposable
{
    private readonly ConcurrentDictionary<string, string> _values = new ConcurrentDictionary<string, string>();
    private int _index;
    private bool _isDisposed;
    private readonly object _lock = new object();

    public DataGridRow()
    {
        // 初始化
    }

    /// <summary>
    /// 行索引
    /// </summary>
    public int Index
    {
        get => _index;
        set => this.RaiseAndSetIfChanged(ref _index, value);
    }

    /// <summary>
    /// 设置值（线程安全）
    /// </summary>
    public void SetValue(string key, string value)
    {
        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        if (_isDisposed)
            throw new ObjectDisposedException(nameof(DataGridRow));

        lock (_lock)
        {
            if (_values.TryGetValue(key, out var existingValue))
            {
                if (existingValue != value)
                {
                    _values[key] = value ?? string.Empty;
                    this.RaisePropertyChanged(key);
                }
            }
            else
            {
                _values.TryAdd(key, value ?? string.Empty);
                this.RaisePropertyChanged(key);
            }
        }
    }

    /// <summary>
    /// 获取值（线程安全）
    /// </summary>
    public string GetValue(string key)
    {
        if (string.IsNullOrEmpty(key))
            return string.Empty;

        if (_isDisposed)
            return string.Empty;

        return _values.TryGetValue(key, out var value) ? value : string.Empty;
    }

    /// <summary>
    /// 检查是否包含指定列（线程安全）
    /// </summary>
    public bool HasColumn(string key)
    {
        if (string.IsNullOrEmpty(key))
            return false;

        return _values.ContainsKey(key);
    }

    /// <summary>
    /// 获取所有列名（线程安全）
    /// </summary>
    public IEnumerable<string> GetColumnNames()
    {
        return _values.Keys.ToList();
    }

    /// <summary>
    /// 索引器 - 用于 DataGrid 绑定（线程安全）
    /// </summary>
    public string this[string columnName]
    {
        get => GetValue(columnName);
        set => SetValue(columnName, value);
    }

    /// <summary>
    /// 获取所有值（线程安全）- 返回副本
    /// </summary>
    public IReadOnlyDictionary<string, string> GetAllValues()
    {
        return new ReadOnlyDictionary<string, string>(
            _values.ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
        );
    }

    /// <summary>
    /// 获取所有值的快照（性能更好）
    /// </summary>
    public Dictionary<string, string> GetValuesSnapshot()
    {
        lock (_lock)
        {
            return _values.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }
    }

    /// <summary>
    /// 批量设置值（线程安全）
    /// </summary>
    public void SetValues(Dictionary<string, string> values)
    {
        if (values == null)
            throw new ArgumentNullException(nameof(values));

        if (_isDisposed)
            throw new ObjectDisposedException(nameof(DataGridRow));

        lock (_lock)
        {
            foreach (var kvp in values)
            {
                if (_values.TryGetValue(kvp.Key, out var existingValue))
                {
                    if (existingValue != kvp.Value)
                    {
                        _values[kvp.Key] = kvp.Value ?? string.Empty;
                        this.RaisePropertyChanged(kvp.Key);
                    }
                }
                else
                {
                    _values.TryAdd(kvp.Key, kvp.Value ?? string.Empty);
                    this.RaisePropertyChanged(kvp.Key);
                }
            }
        }
    }

    /// <summary>
    /// 批量设置值（使用 IEnumerable）
    /// </summary>
    public void SetValues(IEnumerable<KeyValuePair<string, string>> values)
    {
        if (values == null)
            throw new ArgumentNullException(nameof(values));

        if (_isDisposed)
            throw new ObjectDisposedException(nameof(DataGridRow));

        lock (_lock)
        {
            foreach (var kvp in values)
            {
                if (_values.TryGetValue(kvp.Key, out var existingValue))
                {
                    if (existingValue != kvp.Value)
                    {
                        _values[kvp.Key] = kvp.Value ?? string.Empty;
                        this.RaisePropertyChanged(kvp.Key);
                    }
                }
                else
                {
                    _values.TryAdd(kvp.Key, kvp.Value ?? string.Empty);
                    this.RaisePropertyChanged(kvp.Key);
                }
            }
        }
    }

    /// <summary>
    /// 重置所有值（线程安全）
    /// </summary>
    public void Reset()
    {
        if (_isDisposed)
            return;

        lock (_lock)
        {
            var keys = _values.Keys.ToList();
            foreach (var key in keys)
            {
                _values[key] = string.Empty;
                this.RaisePropertyChanged(key);
            }
            Index = 0;
        }
    }

    /// <summary>
    /// 重置为指定值（线程安全）
    /// </summary>
    public void Reset(Dictionary<string, string> initialValues)
    {
        if (_isDisposed)
            return;

        lock (_lock)
        {
            _values.Clear();
            if (initialValues != null)
            {
                foreach (var kvp in initialValues)
                {
                    _values.TryAdd(kvp.Key, kvp.Value ?? string.Empty);
                    this.RaisePropertyChanged(kvp.Key);
                }
            }
            Index = 0;
        }
    }

    /// <summary>
    /// 清空所有值（线程安全）
    /// </summary>
    public void Clear()
    {
        if (_isDisposed)
            return;

        lock (_lock)
        {
            var keys = _values.Keys.ToList();
            _values.Clear();
            foreach (var key in keys)
            {
                this.RaisePropertyChanged(key);
            }
        }
    }

    /// <summary>
    /// 获取值的数量（线程安全）
    /// </summary>
    public int Count => _values.Count;

    /// <summary>
    /// 检查是否包含任何值（线程安全）
    /// </summary>
    public bool HasValues => !_values.IsEmpty;

    /// <summary>
    /// 比较两个行是否相等（基于值）
    /// </summary>
    public bool Equals(DataGridRow? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        lock (_lock)
        {
            var otherValues = other.GetValuesSnapshot();
            if (_values.Count != otherValues.Count)
                return false;

            foreach (var kvp in _values)
            {
                if (!otherValues.TryGetValue(kvp.Key, out var otherValue))
                    return false;

                if (kvp.Value != otherValue)
                    return false;
            }

            return true;
        }
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as DataGridRow);
    }

    public override int GetHashCode()
    {
        lock (_lock)
        {
            return HashCode.Combine(_values.Count, Index);
        }
    }

    #region IDisposable 实现

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_isDisposed)
            return;

        if (disposing)
        {
            lock (_lock)
            {
                _values.Clear();
            }
        }

        _isDisposed = true;
    }

    #endregion
}