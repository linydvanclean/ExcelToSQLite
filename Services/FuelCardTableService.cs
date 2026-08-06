using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services
{
    /// <summary>
    /// 加油卡表专用服务 - 负责创建和管理加油卡相关的数据库表
    /// </summary>
    public class FuelCardTableService
    {
        private readonly DatabaseService _databaseService;

        public FuelCardTableService()
        {
            _databaseService = DatabaseService.Instance;
        }

        /// <summary>
        /// 创建加油卡表（使用中文列名），如果表名已存在且结构不匹配，自动生成新表名
        /// </summary>
        /// <returns>实际使用的表名</returns>
        public async Task<string> CreateTableAsync(string tableName, List<string> columns)
        {
            // 检查表是否存在
            var tableExists = await _databaseService.TableExistsAsync(tableName);
            
            if (tableExists)
            {
                // 验证表结构
                var isValid = await ValidateTableStructureAsync(tableName);
                if (isValid)
                {
                    return tableName; // 表已存在且结构正确，直接使用
                }
                
                // 表结构不匹配，生成新表名
                var newTableName = await GenerateAvailableTableNameAsync(tableName);
                
                // 创建新表
                await CreateNewTableAsync(newTableName, columns);
                return newTableName;
            }

            // 表不存在，直接创建
            await CreateNewTableAsync(tableName, columns);
            return tableName;
        }

        /// <summary>
        /// 生成可用的表名（如果表名已存在且结构不匹配，添加序号后缀）
        /// </summary>
        public async Task<string> GenerateAvailableTableNameAsync(string baseTableName)
        {
            int counter = 1;
            string newTableName;
            
            do
            {
                newTableName = $"{baseTableName}_{counter}";
                counter++;
                
                // 检查表是否存在
                var exists = await _databaseService.TableExistsAsync(newTableName);
                if (!exists)
                {
                    return newTableName;
                }
                
                // 如果表存在，验证结构
                var isValid = await ValidateTableStructureAsync(newTableName);
                if (isValid)
                {
                    // 结构正确，可以使用这个表名
                    return newTableName;
                }
                // 否则继续尝试下一个序号
                
            } while (counter <= 100); // 最多尝试100次，防止死循环
            
            // 如果尝试了100次都失败，使用时间戳
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            return $"{baseTableName}_{timestamp}";
        }

        /// <summary>
        /// 创建新的加油卡表
        /// </summary>
        private async Task CreateNewTableAsync(string tableName, List<string> columns)
        {
            // 定义标准列（使用中文列名）
            var standardColumns = new List<string>
            {
                "Id INTEGER PRIMARY KEY AUTOINCREMENT",
                "卡号 TEXT",
                "交易时间 TEXT",
                "业务类型 TEXT",
                "油品类型 TEXT",
                "数量 REAL",
                "单价 REAL",
                "金额 REAL",
                "奖励分值 REAL",
                "优惠价 REAL",
                "余额 REAL",
                "地点 TEXT",
                "操作员 TEXT",
                "备注 TEXT",
                "客户名称 TEXT",
                "网点名称 TEXT",
                "创建时间 TEXT"
            };

            // 如果有额外的列，添加进去
            foreach (var col in columns)
            {
                var colName = SanitizeIdentifier(col);
                if (!standardColumns.Any(c => c.StartsWith(colName + " ")))
                {
                    standardColumns.Add($"\"{colName}\" TEXT");
                }
            }

            // 构建建表SQL
            string createTableSql = $"CREATE TABLE \"{tableName}\" ({string.Join(", ", standardColumns)})";

            // 执行建表
            await _databaseService.ExecuteNonQueryAsync(createTableSql, new List<SqliteParameter>());

            // 创建索引
            await CreateIndexesAsync(tableName);
        }

        /// <summary>
        /// 创建索引
        /// </summary>
        private async Task CreateIndexesAsync(string tableName)
        {
            var indexCommands = new[]
            {
                $"CREATE INDEX IF NOT EXISTS idx_{tableName}_卡号 ON \"{tableName}\" (\"卡号\")",
                $"CREATE INDEX IF NOT EXISTS idx_{tableName}_交易时间 ON \"{tableName}\" (\"交易时间\")",
                $"CREATE INDEX IF NOT EXISTS idx_{tableName}_业务类型 ON \"{tableName}\" (\"业务类型\")",
                $"CREATE INDEX IF NOT EXISTS idx_{tableName}_地点 ON \"{tableName}\" (\"地点\")",
                $"CREATE INDEX IF NOT EXISTS idx_{tableName}_油品类型 ON \"{tableName}\" (\"油品类型\")"
            };

            foreach (var indexCmd in indexCommands)
            {
                try
                {
                    await _databaseService.ExecuteNonQueryAsync(indexCmd, new List<SqliteParameter>());
                }
                catch
                {
                    // 索引可能已存在，忽略错误
                }
            }
        }

        /// <summary>
        /// 验证表结构是否正确
        /// </summary>
        public async Task<bool> ValidateTableStructureAsync(string tableName)
        {
            try
            {
                // 获取表结构
                var sql = $"PRAGMA table_info(\"{tableName}\")";
                var result = await _databaseService.ExecuteQueryAsync(sql, new List<SqliteParameter>());

                if (result == null || result.Count < 2)
                    return false;

                // 跳过表头行，从第1行开始
                var existingColumns = new List<string>();
                for (int i = 1; i < result.Count; i++)
                {
                    var row = result[i];
                    if (row != null && row.Count > 1)
                    {
                        var columnName = row[1]?.ToString() ?? string.Empty;
                        if (!string.IsNullOrEmpty(columnName))
                        {
                            existingColumns.Add(columnName);
                        }
                    }
                }

                // 必需的列（使用中文列名）
                var requiredColumns = new[]
                {
                    "Id", "卡号", "交易时间", "业务类型",
                    "数量", "单价", "金额", "余额", "创建时间"
                };

                var missingColumns = new List<string>();
                foreach (var requiredCol in requiredColumns)
                {
                    if (!existingColumns.Contains(requiredCol, StringComparer.OrdinalIgnoreCase))
                    {
                        missingColumns.Add(requiredCol);
                    }
                }

                return missingColumns.Count == 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 清理标识符（移除危险字符）
        /// </summary>
        private string SanitizeIdentifier(string identifier)
        {
            if (string.IsNullOrEmpty(identifier))
                return "Column";

            var sanitized = identifier
                .Replace("'", "")
                .Replace(";", "")
                .Replace("--", "")
                .Replace("/", "_")
                .Replace("\\", "_")
                .Replace("\n", "")
                .Replace("\r", "")
                .Replace("\t", "")
                .Trim();

            if (string.IsNullOrEmpty(sanitized))
                sanitized = "Column";

            if (char.IsDigit(sanitized[0]))
                sanitized = "C" + sanitized;

            sanitized = sanitized.Replace(" ", "_");

            return sanitized;
        }

        /// <summary>
        /// 批量插入加油卡数据
        /// </summary>
        public async Task<int> InsertRecordsAsync(string tableName, List<FuelCardRecord> records)
        {
            if (records == null || records.Count == 0)
                return 0;

            var columns = new List<string>
            {
                "卡号",
                "交易时间",
                "业务类型",
                "油品类型",
                "数量",
                "单价",
                "金额",
                "奖励分值",
                "优惠价",
                "余额",
                "地点",
                "操作员",
                "备注",
                "客户名称",
                "网点名称",
                "创建时间"
            };

            var rows = new List<List<object>>();
            var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            foreach (var record in records)
            {
                var row = new List<object>
                {
                    record.CardNumber ?? string.Empty,
                    record.TransactionTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    record.BusinessType ?? string.Empty,
                    record.FuelType ?? string.Empty,
                    record.Quantity.ToString("F2"),
                    record.UnitPrice.ToString("F2"),
                    record.Amount.ToString("F2"),
                    record.BonusPoints.ToString("F2"),
                    record.DiscountPrice.ToString("F2"),
                    record.Balance.ToString("F2"),
                    record.Location ?? string.Empty,
                    record.Operator ?? string.Empty,
                    record.Remarks ?? string.Empty,
                    record.CustomerName ?? string.Empty,
                    record.NetworkName ?? string.Empty,
                    now
                };
                rows.Add(row);
            }

            await _databaseService.InsertDataAsync(tableName, columns, rows);
            return records.Count;
        }

        /// <summary>
        /// 检查表是否存在
        /// </summary>
        public async Task<bool> TableExistsAsync(string tableName)
        {
            return await _databaseService.TableExistsAsync(tableName);
        }

        /// <summary>
        /// 删除表
        /// </summary>
        public async Task DropTableAsync(string tableName)
        {
            await _databaseService.DropTableAsync(tableName);
        }

        /// <summary>
        /// 获取表中所有记录
        /// </summary>
        public async Task<List<FuelCardRecord>> GetAllRecordsAsync(string tableName)
        {
            var records = new List<FuelCardRecord>();

            var sql = $"SELECT * FROM \"{tableName}\" ORDER BY \"交易时间\" DESC";
            var result = await _databaseService.ExecuteQueryAsync(sql, new List<SqliteParameter>());

            if (result == null || result.Count < 2)
                return records;

            // 获取列索引映射
            var header = result[0];
            var columnIndex = new Dictionary<string, int>();
            for (int i = 0; i < header.Count; i++)
            {
                var colName = header[i]?.ToString() ?? string.Empty;
                columnIndex[colName] = i;
            }

            // 解析数据行（从第1行开始）
            for (int i = 1; i < result.Count; i++)
            {
                var row = result[i];
                try
                {
                    var record = new FuelCardRecord();

                    if (columnIndex.TryGetValue("卡号", out int idx) && idx < row.Count)
                        record.CardNumber = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("交易时间", out idx) && idx < row.Count)
                    {
                        if (DateTime.TryParse(row[idx]?.ToString(), out DateTime time))
                            record.TransactionTime = time;
                    }

                    if (columnIndex.TryGetValue("业务类型", out idx) && idx < row.Count)
                        record.BusinessType = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("油品类型", out idx) && idx < row.Count)
                        record.FuelType = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("数量", out idx) && idx < row.Count)
                    {
                        if (decimal.TryParse(row[idx]?.ToString(), out decimal qty))
                            record.Quantity = qty;
                    }

                    if (columnIndex.TryGetValue("单价", out idx) && idx < row.Count)
                    {
                        if (decimal.TryParse(row[idx]?.ToString(), out decimal price))
                            record.UnitPrice = price;
                    }

                    if (columnIndex.TryGetValue("金额", out idx) && idx < row.Count)
                    {
                        if (decimal.TryParse(row[idx]?.ToString(), out decimal amount))
                            record.Amount = amount;
                    }

                    if (columnIndex.TryGetValue("奖励分值", out idx) && idx < row.Count)
                    {
                        if (decimal.TryParse(row[idx]?.ToString(), out decimal bonus))
                            record.BonusPoints = bonus;
                    }

                    if (columnIndex.TryGetValue("优惠价", out idx) && idx < row.Count)
                    {
                        if (decimal.TryParse(row[idx]?.ToString(), out decimal discount))
                            record.DiscountPrice = discount;
                    }

                    if (columnIndex.TryGetValue("余额", out idx) && idx < row.Count)
                    {
                        if (decimal.TryParse(row[idx]?.ToString(), out decimal balance))
                            record.Balance = balance;
                    }

                    if (columnIndex.TryGetValue("地点", out idx) && idx < row.Count)
                        record.Location = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("操作员", out idx) && idx < row.Count)
                        record.Operator = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("备注", out idx) && idx < row.Count)
                        record.Remarks = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("客户名称", out idx) && idx < row.Count)
                        record.CustomerName = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("网点名称", out idx) && idx < row.Count)
                        record.NetworkName = row[idx]?.ToString();

                    if (columnIndex.TryGetValue("创建时间", out idx) && idx < row.Count)
                    {
                        if (DateTime.TryParse(row[idx]?.ToString(), out DateTime created))
                            record.CreatedAt = created;
                    }

                    records.Add(record);
                }
                catch
                {
                    // 跳过解析失败的行
                }
            }

            return records;
        }

        /// <summary>
        /// 获取统计信息
        /// </summary>
        public async Task<(int TotalRecords, int CardCount, decimal TotalAmount, decimal TotalFuelVolume)> 
            GetStatisticsAsync(string tableName)
        {
            try
            {
                // 总记录数
                var countSql = $"SELECT COUNT(*) FROM \"{tableName}\"";
                var countResult = await _databaseService.ExecuteQueryAsync(countSql, new List<SqliteParameter>());
                int totalRecords = 0;
                if (countResult != null && countResult.Count > 1 && countResult[1].Count > 0)
                {
                    int.TryParse(countResult[1][0]?.ToString(), out totalRecords);
                }

                // 卡数
                var cardSql = $"SELECT COUNT(DISTINCT \"卡号\") FROM \"{tableName}\"";
                var cardResult = await _databaseService.ExecuteQueryAsync(cardSql, new List<SqliteParameter>());
                int cardCount = 0;
                if (cardResult != null && cardResult.Count > 1 && cardResult[1].Count > 0)
                {
                    int.TryParse(cardResult[1][0]?.ToString(), out cardCount);
                }

                // 总金额
                var amountSql = $"SELECT SUM(\"金额\") FROM \"{tableName}\"";
                var amountResult = await _databaseService.ExecuteQueryAsync(amountSql, new List<SqliteParameter>());
                decimal totalAmount = 0;
                if (amountResult != null && amountResult.Count > 1 && amountResult[1].Count > 0)
                {
                    decimal.TryParse(amountResult[1][0]?.ToString(), out totalAmount);
                }

                // 总升数
                var volumeSql = $"SELECT SUM(\"数量\") FROM \"{tableName}\"";
                var volumeResult = await _databaseService.ExecuteQueryAsync(volumeSql, new List<SqliteParameter>());
                decimal totalFuelVolume = 0;
                if (volumeResult != null && volumeResult.Count > 1 && volumeResult[1].Count > 0)
                {
                    decimal.TryParse(volumeResult[1][0]?.ToString(), out totalFuelVolume);
                }

                return (totalRecords, cardCount, totalAmount, totalFuelVolume);
            }
            catch
            {
                return (0, 0, 0, 0);
            }
        }
    }
}