using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using ExcelToSQLite.Services.TableDefinitions;

namespace ExcelToSQLite.Services;

/// <summary>
/// 表初始化服务 - 统一管理所有表的创建和初始化
/// </summary>
public class TableInitializerService : DisposableBase
{
    private readonly DatabaseService _databaseService;
    private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);
    private bool _isInitialized = false;
    private readonly Dictionary<string, bool> _tableStatus = new Dictionary<string, bool>();
    private readonly List<ITableDefinition> _tableDefinitions = new List<ITableDefinition>();

    public TableInitializerService()
    {
        _databaseService = DatabaseService.Instance;
        RegisterDisposable(_initLock);
        
        // 注册所有表定义
        RegisterTableDefinitions();
    }

    protected override void DisposeManagedResources()
    {
    }

    /// <summary>
    /// 注册所有表定义
    /// </summary>
    private void RegisterTableDefinitions()
    {
        // 注册各个表
        _tableDefinitions.Add(new UserTableDefinition());
        _tableDefinitions.Add(new AnalysisBatchsTableDefinition());
        _tableDefinitions.Add(new IndicatorTableDefinition());
        _tableDefinitions.Add(new DataDictionaryTableDefinition());
        _tableDefinitions.Add(new ScanResultsTableDefinition());
        // 添加更多表...
        

    }

    /// <summary>
    /// 初始化所有表（应用启动时调用一次）
    /// </summary>
    public async Task InitializeAllTablesAsync()
    {
        if (_isInitialized) return;

        await _initLock.WaitAsync();
        try
        {
            if (_isInitialized) return;


            
            var totalTables = _tableDefinitions.Count;
            var createdCount = 0;
            var errorCount = 0;

            foreach (var tableDef in _tableDefinitions)
            {
                try
                {
                    await InitializeTableAsync(tableDef);
                    createdCount++;
                }
                catch
                {
                    errorCount++;
                }
            }

            _isInitialized = true;
            

        }
        finally
        {
            _initLock.Release();
        }
    }

    /// <summary>
    /// 初始化单个表
    /// </summary>
    private async Task InitializeTableAsync(ITableDefinition tableDef)
    {


        // 1. 创建表
        await _databaseService.ExecuteNonQueryAsync(tableDef.CreateTableSql, new List<object>());
        
        // 2. 创建索引
        if (!string.IsNullOrEmpty(tableDef.CreateIndexSql))
        {
            await _databaseService.ExecuteNonQueryAsync(tableDef.CreateIndexSql, new List<object>());
        }
        
        // 3. 初始化默认数据
        await tableDef.InitializeDefaultDataAsync(_databaseService);
        
        // 4. 标记为已初始化
        _tableStatus[tableDef.TableName] = true;
        

    }

    /// <summary>
    /// 检查表是否存在
    /// </summary>
    public async Task<bool> TableExistsAsync(string tableName)
    {
        var sql = @"SELECT name FROM sqlite_master WHERE type='table' AND name=@p0";
        var parameters = new List<object> { tableName };
        var result = await _databaseService.ExecuteQueryAsync(sql, parameters);
        return result != null && result.Count > 1;
    }

    /// <summary>
    /// 获取所有表名
    /// </summary>
    public async Task<List<string>> GetAllTableNamesAsync()
    {
        var sql = "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name";
        var result = await _databaseService.ExecuteQueryAsync(sql, new List<object>());
        var tables = new List<string>();
        
        if (result != null && result.Count > 1)
        {
            for (int i = 1; i < result.Count; i++)
            {
                if (result[i].Count > 0)
                {
                    tables.Add(result[i][0]?.ToString() ?? string.Empty);
                }
            }
        }
        
        return tables;
    }
}