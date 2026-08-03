using ReactiveUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;

namespace ExcelToSQLite.ViewModels;

/// <summary>
/// 动态数据行项 - 支持 DataGrid 绑定
/// 线程安全的实现
/// </summary>
public class DataRowItem : ReactiveObject, IDisposable
{
    private readonly ConcurrentDictionary<string, string> _values = new ConcurrentDictionary<string, string>();
    private int _index;
    private bool _isDisposed;
    private readonly object _lock = new object();

    public DataRowItem()
    {
        // 初始化空字典
    }

    /// <summary>
    /// 使用初始值创建数据行
    /// </summary>
    public DataRowItem(Dictionary<string, string> initialValues)
    {
        if (initialValues != null)
        {
            foreach (var kvp in initialValues)
            {
                _values.TryAdd(kvp.Key, kvp.Value ?? string.Empty);
            }
        }
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
    /// 获取所有值的只读副本
    /// </summary>
    public IReadOnlyDictionary<string, string> Values => 
        new ReadOnlyDictionary<string, string>(
            _values.ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
        );

    /// <summary>
    /// 获取值的数量
    /// </summary>
    public int Count => _values.Count;

    /// <summary>
    /// 是否有任何值
    /// </summary>
    public bool HasValues => !_values.IsEmpty;

    /// <summary>
    /// 索引器 - 用于 DataGrid 绑定（线程安全）
    /// </summary>
    public string this[string key]
    {
        get => GetValue(key);
        set => SetValue(key, value);
    }

    /// <summary>
    /// 设置值（线程安全）
    /// </summary>
    public void SetValue(string key, string value)
    {
        if (string.IsNullOrEmpty(key))
            throw new ArgumentNullException(nameof(key));

        if (_isDisposed)
            throw new ObjectDisposedException(nameof(DataRowItem));

        lock (_lock)
        {
            var newValue = value ?? string.Empty;
            if (_values.TryGetValue(key, out var existingValue))
            {
                if (existingValue != newValue)
                {
                    _values[key] = newValue;
                    this.RaisePropertyChanged($"Item[{key}]");
                    this.RaisePropertyChanged($"Values");
                }
            }
            else
            {
                _values.TryAdd(key, newValue);
                this.RaisePropertyChanged($"Item[{key}]");
                this.RaisePropertyChanged($"Values");
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
    /// 检查是否包含指定键（线程安全）
    /// </summary>
    public bool HasKey(string key)
    {
        if (string.IsNullOrEmpty(key))
            return false;

        return _values.ContainsKey(key);
    }

    /// <summary>
    /// 获取所有键（线程安全）
    /// </summary>
    public IEnumerable<string> GetKeys()
    {
        return _values.Keys.ToList();
    }

    /// <summary>
    /// 获取所有键值对的快照（线程安全）
    /// </summary>
    public Dictionary<string, string> GetSnapshot()
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
            throw new ObjectDisposedException(nameof(DataRowItem));

        lock (_lock)
        {
            foreach (var kvp in values)
            {
                var newValue = kvp.Value ?? string.Empty;
                if (_values.TryGetValue(kvp.Key, out var existingValue))
                {
                    if (existingValue != newValue)
                    {
                        _values[kvp.Key] = newValue;
                        this.RaisePropertyChanged($"Item[{kvp.Key}]");
                    }
                }
                else
                {
                    _values.TryAdd(kvp.Key, newValue);
                    this.RaisePropertyChanged($"Item[{kvp.Key}]");
                }
            }
            this.RaisePropertyChanged($"Values");
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
            throw new ObjectDisposedException(nameof(DataRowItem));

        lock (_lock)
        {
            foreach (var kvp in values)
            {
                var newValue = kvp.Value ?? string.Empty;
                if (_values.TryGetValue(kvp.Key, out var existingValue))
                {
                    if (existingValue != newValue)
                    {
                        _values[kvp.Key] = newValue;
                        this.RaisePropertyChanged($"Item[{kvp.Key}]");
                    }
                }
                else
                {
                    _values.TryAdd(kvp.Key, newValue);
                    this.RaisePropertyChanged($"Item[{kvp.Key}]");
                }
            }
            this.RaisePropertyChanged($"Values");
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
                this.RaisePropertyChanged($"Item[{key}]");
            }
            this.RaisePropertyChanged($"Values");
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
                    this.RaisePropertyChanged($"Item[{kvp.Key}]");
                }
            }
            this.RaisePropertyChanged($"Values");
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
                this.RaisePropertyChanged($"Item[{key}]");
            }
            this.RaisePropertyChanged($"Values");
        }
    }

    /// <summary>
    /// 移除指定键（线程安全）
    /// </summary>
    public bool RemoveKey(string key)
    {
        if (string.IsNullOrEmpty(key))
            return false;

        if (_isDisposed)
            return false;

        lock (_lock)
        {
            if (_values.TryRemove(key, out _))
            {
                this.RaisePropertyChanged($"Item[{key}]");
                this.RaisePropertyChanged($"Values");
                return true;
            }
            return false;
        }
    }

    /// <summary>
    /// 比较两个行是否相等（基于值）
    /// </summary>
    public bool Equals(DataRowItem? other)  // ← 修复：参数标记为可空
    {
        if (other is null)  // ← 修复：使用 is null 检查
            return false;

        if (ReferenceEquals(this, other))
            return true;

        lock (_lock)
        {
            var otherValues = other.GetSnapshot();
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
        return Equals(obj as DataRowItem);
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