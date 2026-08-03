using System.Threading.Tasks;

namespace ExcelToSQLite.Services;

/// <summary>
/// 表定义接口
/// </summary>
public interface ITableDefinition
{
    /// <summary>
    /// 表名
    /// </summary>
    string TableName { get; }
    
    /// <summary>
    /// 创建表的 SQL
    /// </summary>
    string CreateTableSql { get; }
    
    /// <summary>
    /// 创建索引的 SQL（可选）
    /// </summary>
    string? CreateIndexSql { get; }
    
    /// <summary>
    /// 初始化默认数据（可选）
    /// </summary>
    Task InitializeDefaultDataAsync(DatabaseService dbService);
}