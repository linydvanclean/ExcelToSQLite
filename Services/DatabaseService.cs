using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToSQLite.Services;

public class DatabaseService : IDisposable
{
    private static DatabaseService? _instance;
    private static readonly object _instanceLock = new object();
    private bool _isInitialized = false;
    private readonly object _initLock = new object();
    
    private bool _disposed;
    private readonly string _databasePath;
    private readonly string _connectionString;
    private readonly string _instanceId;

    // 操作队列（串行化所有数据库写操作，SQLite 仅支持单写者）
    private readonly SemaphoreSlim _operationSemaphore = new SemaphoreSlim(1, 1);
    
    // 超时时间
    private static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(30);
    
    // ✅ 数据库验证状态 - 只验证一次
    private bool _databaseValidated = false;
    private readonly object _validationLock = new object();
    
    /// <summary>
    /// 获取单例实例
    /// </summary>
    public static DatabaseService Instance
    {
        get
        {
            if (_instance == null)
            {
                lock (_instanceLock)
                {
                    if (_instance == null)
                    {
                        _instance = new DatabaseService();
                    }
                }
            }
            return _instance;
        }
    }
    
    private DatabaseService(string databaseFile = "Data.db")
    {
        _instanceId = Guid.NewGuid().ToString("N").Substring(0, 8);
        var baseDirectory = AppContext.BaseDirectory;
        
        // ========== 检查路径是否包含中文字符 ==========
        if (HasChineseCharacters(baseDirectory))
        {
        }

        string dataDirectory;
        
        // 检查是否有写入权限，且路径不含中文字符
        if (HasWritePermission(baseDirectory) && !HasChineseCharacters(baseDirectory))
        {
            dataDirectory = Path.Combine(baseDirectory, "data");
        }
        else
        {
            // 切换到用户目录（不含中文字符）
            var userHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var appName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name ?? "ExcelToSQLite";
            
            // 确保 appName 不含中文字符
            appName = RemoveChineseCharacters(appName);
            if (string.IsNullOrEmpty(appName))
            {
                appName = "ExcelToSQLite";
            }
            
            dataDirectory = Path.Combine(userHome, ".local", "share", appName, "data");
        }
        
        // ========== 检查 dataDirectory 是否包含中文字符 ==========
        if (HasChineseCharacters(dataDirectory))
        {
            dataDirectory = Path.Combine(Path.GetTempPath(), "ExcelToSQLite", "data");
        }
        
        // 创建目录
        if (!Directory.Exists(dataDirectory))
        {
            try
            {
                Directory.CreateDirectory(dataDirectory);
            }
            catch
            {
                dataDirectory = Path.Combine(Path.GetTempPath(), "ExcelToSQLite", "data");
                if (!Directory.Exists(dataDirectory))
                {
                    Directory.CreateDirectory(dataDirectory);
                }
            }
        }
        
        // ========== 构建数据库路径 ==========
        // 确保文件名不含中文字符
        var sanitizedDbFile = RemoveChineseCharacters(databaseFile);
        if (string.IsNullOrEmpty(sanitizedDbFile) || !sanitizedDbFile.EndsWith(".db"))
        {
            sanitizedDbFile = "Data.db";
        }
        
        _databasePath = Path.Combine(dataDirectory, sanitizedDbFile);
        
        // ========== 创建数据库文件 ==========
        try
        {
            if (!File.Exists(_databasePath))
            {
                using (File.Create(_databasePath)) { }
            }
            
            // ========== 测试文件写入权限 ==========
            TestFileWritePermission(_databasePath);
        }
        catch
        {

            var tempDbPath = Path.Combine(Path.GetTempPath(), "ExcelToSQLite", "Data.db");
            var tempDir = Path.GetDirectoryName(tempDbPath);
            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir!);
            }
            
            if (!File.Exists(tempDbPath))
            {
                using (File.Create(tempDbPath)) { }
            }
            
            _databasePath = tempDbPath;
        }
        
        _connectionString = $"Data Source={_databasePath};Mode=ReadWriteCreate;Cache=Shared;Pooling=True;Default Timeout=30;";
        
        // ========== 验证数据库文件（只在构造函数中执行一次） ==========
        EnsureDatabaseValid();
    }

    // ========== 数据库文件验证方法 ==========

    /// <summary>
    /// 验证是否为有效的 SQLite 数据库文件（更健壮的实现）
    /// </summary>
    private bool IsValidSqliteDatabase(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return false;

            var fileInfo = new FileInfo(filePath);
            if (fileInfo.Length < 100) // 太小的文件可能是空的或损坏的
                return false;
            
            var header = new byte[16];
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (fs.Length < 16)
                    return false;
                var readCount = fs.Read(header, 0, 16);
                if (readCount < 16)
                    return false;
            }

            var headerString = System.Text.Encoding.ASCII.GetString(header);
            return headerString.StartsWith("SQLite format 3");
        }
        catch (IOException)
        {
            // 文件被锁定，视为有效（可能是正在使用）
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 备份并删除损坏的数据库文件
    /// </summary>
    private void BackupAndDeleteDatabase()
    {
        try
        {
            if (!File.Exists(_databasePath))
                return;

            var backupPath = _databasePath + ".backup." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            File.Copy(_databasePath, backupPath);
            File.Delete(_databasePath);
            using (File.Create(_databasePath)) { }
        }
        catch
        {
        }
    }

    /// <summary>
    /// 确保数据库文件有效，如果损坏则自动修复（只在第一次调用时执行）
    /// </summary>
    private void EnsureDatabaseValid()
    {
        // ✅ 如果已经验证过，直接返回
        if (_databaseValidated) return;
        
        lock (_validationLock)
        {
            if (_databaseValidated) return;
            try
            {
                // 如果文件不存在，直接返回（第一次运行，会创建）
                if (!File.Exists(_databasePath))
                {
                    _databaseValidated = true;
                    return;
                }

                // 验证数据库文件是否有效
                if (!IsValidSqliteDatabase(_databasePath))
                {
                    BackupAndDeleteDatabase();
                    _databaseValidated = true;
                    return;
                }

                // 尝试打开连接测试
                try
                {
                    using var testConnection = new SqliteConnection(_connectionString);
                    testConnection.Open();
                    testConnection.Close();
                }
                catch
                {
                    BackupAndDeleteDatabase();
                }
            }
            catch
            {
                if (!IsValidSqliteDatabase(_databasePath))
                {
                    BackupAndDeleteDatabase();
                }
            }
            finally
            {
                _databaseValidated = true;
            }
        }
    }

    // ========== 辅助方法 ==========

    /// <summary>
    /// 检查字符串是否包含中文字符
    /// </summary>
    private bool HasChineseCharacters(string text)
    {
        if (string.IsNullOrEmpty(text)) return false;
        
        foreach (char c in text)
        {
            if ((c >= 0x4E00 && c <= 0x9FFF) ||
                (c >= 0x3400 && c <= 0x4DBF) ||
                (c >= 0xF900 && c <= 0xFAFF))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// 移除字符串中的中文字符
    /// </summary>
    private string RemoveChineseCharacters(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        
        var result = new System.Text.StringBuilder();
        foreach (char c in text)
        {
            if (!((c >= 0x4E00 && c <= 0x9FFF) ||
                  (c >= 0x3400 && c <= 0x4DBF) ||
                  (c >= 0xF900 && c <= 0xFAFF)))
            {
                result.Append(c);
            }
        }
        return result.ToString();
    }

    /// <summary>
    /// 测试文件写入权限
    /// </summary>
    private void TestFileWritePermission(string filePath)
    {
        try
        {
            var testContent = $"Test write permission at {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
            File.AppendAllText(filePath, testContent + Environment.NewLine);
        }
        catch
        {
            throw;
        }
    }
    
    private async Task<T> ExecuteOperationAsync<T>(Func<Task<T>> operation, string operationName, CancellationToken cancellationToken = default)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(DefaultTimeout);
        
        try
        {
            if (!await _operationSemaphore.WaitAsync(DefaultTimeout, cts.Token))
            {
                throw new TimeoutException($"操作 {operationName} 获取信号量超时");
            }
            
            try
            {
                var result = await operation();
                return result;
            }
            finally
            {
                _operationSemaphore.Release();
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            throw;
        }
    }

    private async Task ExecuteOperationAsync(Func<Task> operation, string operationName, CancellationToken cancellationToken = default)
    {
        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(DefaultTimeout);
        
        try
        {
            if (!await _operationSemaphore.WaitAsync(DefaultTimeout, cts.Token))
            {
                throw new TimeoutException($"操作 {operationName} 获取信号量超时");
            }
            
            try
            {
                await operation();
            }
            finally
            {
                _operationSemaphore.Release();
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            throw;
        }
    }

    /// <summary>
    /// 确保 SQLite 已初始化（使用内置库）
    /// </summary>
    private void EnsureSQLiteInitialized()
    {
        if (_isInitialized) return;

        lock (_initLock)
        {
            if (_isInitialized) return;

            try
            {
                using var testConnection = new SqliteConnection("Data Source=:memory:");
                testConnection.Open();
                _isInitialized = true;
            }
            catch
            {
                try
                {
                    using var testConnection = new SqliteConnection("Data Source=:memory:");
                    testConnection.Open();


                    _isInitialized = true;
                }
                catch
                {
                    throw;
                }
            }
        }
    }
    
    private bool HasWritePermission(string directory)
    {
        try
        {
            var testFile = Path.Combine(directory, "test_write_permission.tmp");
            using (File.Create(testFile)) { }
            File.Delete(testFile);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private async Task<SqliteConnection> GetConnectionAsync()
    {
        EnsureDatabaseValid();
        EnsureSQLiteInitialized();
        
        var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();

        using var pragmaCmd = connection.CreateCommand();
        pragmaCmd.CommandText = "PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA cache_size=10000;";
        await pragmaCmd.ExecuteNonQueryAsync();

        return connection;
    }
    
    #region 公开方法

    public async Task CreateTableAsync(string tableName, List<string> columns)
    {
        await ExecuteOperationAsync(async () =>
            {
                // 清理表名但保留中文
                var cleanTableName = SanitizeIdentifier(tableName);
        
                // 使用双引号转义表名（SQLite 标准）
                var escapedTableName = $"\"{cleanTableName}\"";

                var uniqueColumns = new List<string>();
                var columnNameCount = new Dictionary<string, int>();

                foreach (var col in columns)
                {
                    var colName = SanitizeIdentifier(col);

                    if (string.IsNullOrEmpty(colName))
                    {
                        colName = "Column";
                    }

                    // 处理 Id 列冲突
                    if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase))
                    {
                        colName = "Id_Data";
                    }

                    // 处理重复列名
                    if (columnNameCount.ContainsKey(colName))
                    {
                        columnNameCount[colName]++;
                        colName = $"{colName}_{columnNameCount[colName]}";
                    }
                    else
                    {
                        columnNameCount[colName] = 1;
                    }

                    uniqueColumns.Add(colName);
                }

                // 构建列定义 - 使用双引号转义每个列名
                var columnsDefinition = new List<string>
                {
                    "Id INTEGER PRIMARY KEY AUTOINCREMENT"
                };

                foreach (var col in uniqueColumns)
                {
                    columnsDefinition.Add($"\"{col}\" TEXT");
                }

                if (columnsDefinition.Count == 1)
                {
                    columnsDefinition.Add("\"Data\" TEXT");
                }

                // 使用转义后的表名
                string createTableSql = $"CREATE TABLE IF NOT EXISTS {escapedTableName} ({string.Join(", ", columnsDefinition)})";

                using var connection = await GetConnectionAsync();
                using var command = new SqliteCommand(createTableSql, connection);
                await command.ExecuteNonQueryAsync();
                
            }, $"CreateTable_{tableName}");
    }

    public async Task InsertDataAsync(string tableName, List<string> columns, List<List<object>> rows)
{
    await ExecuteOperationAsync(async () =>
    {
        var cleanTableName = SanitizeIdentifier(tableName);
        var escapedTableName = $"\"{cleanTableName}\"";

        var sanitizedColumns = new List<string>();
        var columnNameCount = new Dictionary<string, int>();

        foreach (var col in columns)
        {
            var colName = SanitizeIdentifier(col);

            if (string.IsNullOrEmpty(colName))
            {
                colName = "Column";
            }

            if (colName.Equals("Id", StringComparison.OrdinalIgnoreCase))
            {
                colName = "Id_Data";
            }

            if (columnNameCount.ContainsKey(colName))
            {
                columnNameCount[colName]++;
                colName = $"{colName}_{columnNameCount[colName]}";
            }
            else
            {
                columnNameCount[colName] = 1;
            }

            sanitizedColumns.Add(colName);
        }

        using var connection = await GetConnectionAsync();
        using var transaction = await connection.BeginTransactionAsync();

        try
        {
            // 使用双引号转义列名
            var columnNames = sanitizedColumns.Select(c => $"\"{c}\"").ToList();
            var placeholders = sanitizedColumns.Select((_, i) => $"@p{i}").ToList();

            string insertSql =
                $"INSERT INTO {escapedTableName} ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", placeholders)})";

            foreach (var row in rows)
            {
                using var command = new SqliteCommand(insertSql, connection);
                command.Transaction = transaction as SqliteTransaction;

                for (int i = 0; i < sanitizedColumns.Count && i < row.Count; i++)
                {
                    var value = row[i]?.ToString() ?? string.Empty;
                    command.Parameters.AddWithValue($"@p{i}", value);
                }

                await command.ExecuteNonQueryAsync();
            }

            await transaction.CommitAsync();
        }
        catch
        {
            await transaction.RollbackAsync();
            throw;
        }
    }, $"InsertData_{tableName}");
}
    
    public async Task DeleteAllDataAsync(string tableName)
    {
        await ExecuteOperationAsync(async () =>
            {
                var cleanTableName = SanitizeIdentifier(tableName);
                var escapedTableName = $"\"{cleanTableName}\"";

                using var connection = await GetConnectionAsync();
                var sql = $"DELETE FROM {escapedTableName}";

                using var command = new SqliteCommand(sql, connection);
                await command.ExecuteNonQueryAsync();
        
            }, $"DeleteAllData_{tableName}");
    }
    
    public async Task<int> ExecuteNonQueryAsync(string sql, List<object> parameters)
    {
        return await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();
            using var command = connection.CreateCommand();
            command.CommandText = sql;

            for (int i = 0; i < parameters.Count; i++)
            {
                command.Parameters.AddWithValue($"@p{i}", parameters[i] ?? DBNull.Value);
            }

            var result = await command.ExecuteNonQueryAsync();
            return result;
        }, "ExecuteNonQuery");
    }
    
    public async Task<int> ExecuteNonQueryAsync(string sql, List<SqliteParameter> parameters)
    {
        return await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();
            using var command = new SqliteCommand(sql, connection);
            if (parameters != null && parameters.Count > 0)
            {
                command.Parameters.AddRange(parameters.ToArray());
            }
            var result = await command.ExecuteNonQueryAsync();
            return result;
        }, "ExecuteNonQuery");
    }

    public async Task<List<List<object>>> ExecuteQueryAsync(string sql, List<object> parameters)
    {
        return await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();
            using var command = new SqliteCommand(sql, connection);

            for (int i = 0; i < parameters.Count; i++)
            {
                var paramName = $"@p{i}";
                var value = parameters[i]?.ToString() ?? string.Empty;
                command.Parameters.AddWithValue(paramName, value);
            }

            using var reader = await command.ExecuteReaderAsync();

            var result = new List<List<object>>();

            var headerRow = new List<object>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                headerRow.Add(reader.GetName(i));
            }

            var dataRows = new List<List<object>>();
            while (await reader.ReadAsync())
            {
                var row = new List<object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row.Add(reader.GetValue(i));
                }
                dataRows.Add(row);
            }

            if (dataRows.Count > 0)
            {
                result.Add(headerRow);
                result.AddRange(dataRows);
            }
            return result;
        }, "ExecuteQuery");
    }

    public async Task<List<List<object>>> ExecuteQueryAsync(string sql, List<SqliteParameter> parameters)
    {
        return await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();
            using var command = new SqliteCommand(sql, connection);

            if (parameters != null && parameters.Count > 0)
            {
                command.Parameters.AddRange(parameters.ToArray());
            }

            using var reader = await command.ExecuteReaderAsync();

            var result = new List<List<object>>();

            var headerRow = new List<object>();
            for (int i = 0; i < reader.FieldCount; i++)
            {
                headerRow.Add(reader.GetName(i));
            }

            var dataRows = new List<List<object>>();
            while (await reader.ReadAsync())
            {
                var row = new List<object>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row.Add(reader.GetValue(i));
                }
                dataRows.Add(row);
            }

            if (dataRows.Count > 0)
            {
                result.Add(headerRow);
                result.AddRange(dataRows);
            }
            return result;
        }, "ExecuteQuery");
    }
    public async Task<bool> TableExistsAsync(string tableName)
    {
        return await ExecuteOperationAsync(async () =>
            {
                var cleanTableName = SanitizeIdentifier(tableName);

                using var connection = await GetConnectionAsync();
                string sql = "SELECT name FROM sqlite_master WHERE type='table' AND name=@tableName";

                using var command = new SqliteCommand(sql, connection);
                command.Parameters.AddWithValue("@tableName", cleanTableName);

                var result = await command.ExecuteScalarAsync();
                return result != null;
            }, $"TableExists_{tableName}");
    }

    public async Task<List<string>> GetAllTableNamesAsync()
    {
        return await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();
            string sql = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'";

            using var command = new SqliteCommand(sql, connection);
            using var reader = await command.ExecuteReaderAsync();

            var tables = new List<string>();
            while (await reader.ReadAsync())
            {
                tables.Add(reader.GetString(0));
            }

            return tables;
        }, "GetAllTableNames");
    }

    public async Task DropTableAsync(string tableName)
    {
        await ExecuteOperationAsync(async () =>
            {
                var cleanTableName = SanitizeIdentifier(tableName);
                var escapedTableName = $"\"{cleanTableName}\"";

                using var connection = await GetConnectionAsync();
                string sql = $"DROP TABLE IF EXISTS {escapedTableName}";

                using var command = new SqliteCommand(sql, connection);
                await command.ExecuteNonQueryAsync();
                
            }, $"DropTable_{tableName}");
    }

    #endregion

    #region 私有辅助方法

    private string SanitizeIdentifier(string identifier)
    {
        if (string.IsNullOrEmpty(identifier))
            return "Column";

        // 移除危险字符，但保留中文字符
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

        // 如果清理后为空，返回默认值
        if (string.IsNullOrEmpty(sanitized))
        {
            sanitized = "Column";
        }

        // 如果以数字开头，添加前缀
        if (char.IsDigit(sanitized[0]))
        {
            sanitized = "C" + sanitized;
        }

        // 如果包含空格，替换为下划线
        sanitized = sanitized.Replace(" ", "_");

        return sanitized;
    }

    #endregion

    #region 考勤和加油卡表创建

    public async Task CreateAttendanceTableAsync(string tableName, List<string> columns)
    {
        await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();

            var checkCommand = connection.CreateCommand();
            checkCommand.CommandText = $"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName}'";
            var existing = await checkCommand.ExecuteScalarAsync();

            if (existing == null)
            {
                await CreateNewAttendanceTableAsync(connection, tableName, columns);
                return;
            }

            var isValid = await ValidateAttendanceTableStructureAsync(connection, tableName);
            if (isValid)
            {
                return;
            }
            else
            {
                throw new Exception($"表 {tableName} 已存在但结构不正确。\n" +
                                    $"请使用不同的表名，或删除现有表后重试。");
            }
        }, $"CreateAttendanceTable_{tableName}");
    }

    private async Task CreateNewAttendanceTableAsync(SqliteConnection connection, string tableName,
        List<string> columns)
    {
        var columnDefs = new List<string>
        {
            "Id INTEGER PRIMARY KEY AUTOINCREMENT",
            "EmployeeId TEXT",
            "EmployeeName TEXT",
            "Department TEXT",
            "CheckTime TEXT",
            "DayOfMonth INTEGER",
            "CreatedAt TEXT"
        };

        foreach (var col in columns)
        {
            var colName = SanitizeIdentifier(col);
            if (!columnDefs.Any(c => c.StartsWith(colName + " ")))
            {
                columnDefs.Add($"\"{colName}\" TEXT");
            }
        }

        var createCommand = connection.CreateCommand();
        createCommand.CommandText = $"CREATE TABLE \"{tableName}\" ({string.Join(", ", columnDefs)})";
        await createCommand.ExecuteNonQueryAsync();

        var indexCommands = new[]
        {
            $"CREATE INDEX idx_{tableName}_EmployeeId ON \"{tableName}\" (EmployeeId)",
            $"CREATE INDEX idx_{tableName}_CheckTime ON \"{tableName}\" (CheckTime)",
            $"CREATE INDEX idx_{tableName}_Department ON \"{tableName}\" (Department)"
        };

        foreach (var indexCmd in indexCommands)
        {
            try
            {
                var indexCommand = connection.CreateCommand();
                indexCommand.CommandText = indexCmd;
                await indexCommand.ExecuteNonQueryAsync();
            }
            catch
            {
                /* 索引可能已存在 */
            }
        }
        
    }

    private async Task<bool> ValidateAttendanceTableStructureAsync(SqliteConnection connection, string tableName)
    {
        try
        {
            var getColumns = connection.CreateCommand();
            getColumns.CommandText = $"PRAGMA table_info(\"{tableName}\")";

            var existingColumns = new List<string>();
            using var reader = await getColumns.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var columnName = reader.GetString(1);
                existingColumns.Add(columnName);
            }

            var requiredColumns = new[]
            {
                "Id", "EmployeeId", "EmployeeName", "Department",
                "CheckTime", "DayOfMonth", "CreatedAt"
            };

            var missingColumns = new List<string>();
            foreach (var requiredCol in requiredColumns)
            {
                if (!existingColumns.Contains(requiredCol, StringComparer.OrdinalIgnoreCase))
                {
                    missingColumns.Add(requiredCol);
                }
            }

            if (missingColumns.Count > 0)
            {
                return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    public async Task CreateFuelCardTableAsync(string tableName, List<string> columns)
    {
        await ExecuteOperationAsync(async () =>
        {
            using var connection = await GetConnectionAsync();

            var checkCommand = connection.CreateCommand();
            checkCommand.CommandText = $"SELECT name FROM sqlite_master WHERE type='table' AND name='{tableName}'";
            var existing = await checkCommand.ExecuteScalarAsync();

            if (existing == null)
            {
                await CreateNewFuelCardTableAsync(connection, tableName, columns);
                return;
            }

            var isValid = await ValidateFuelCardTableStructureAsync(connection, tableName);
            if (isValid)
            {
                return;
            }
            else
            {
                throw new Exception($"表 {tableName} 已存在但结构不正确。\n" +
                                    $"请使用不同的表名，或删除现有表后重试。");
            }
        }, $"CreateFuelCardTable_{tableName}");
    }

    private async Task CreateNewFuelCardTableAsync(SqliteConnection connection, string tableName, List<string> columns)
    {
        var columnDefs = new List<string>
        {
            "Id INTEGER PRIMARY KEY AUTOINCREMENT",
            "CardNumber TEXT",
            "TransactionTime TEXT",
            "BusinessType TEXT",
            "FuelType TEXT",
            "Quantity REAL",
            "UnitPrice REAL",
            "Amount REAL",
            "BonusPoints REAL",
            "DiscountPrice REAL",
            "Balance REAL",
            "Location TEXT",
            "Operator TEXT",
            "Remarks TEXT",
            "CustomerName TEXT",
            "NetworkName TEXT",
            "CreatedAt TEXT"
        };

        foreach (var col in columns)
        {
            var colName = SanitizeIdentifier(col);
            if (!columnDefs.Any(c => c.StartsWith(colName + " ")))
            {
                columnDefs.Add($"\"{colName}\" TEXT");
            }
        }

        var createCommand = connection.CreateCommand();
        createCommand.CommandText = $"CREATE TABLE \"{tableName}\" ({string.Join(", ", columnDefs)})";
        await createCommand.ExecuteNonQueryAsync();

        var indexCommands = new[]
        {
            $"CREATE INDEX idx_{tableName}_CardNumber ON \"{tableName}\" (CardNumber)",
            $"CREATE INDEX idx_{tableName}_TransactionTime ON \"{tableName}\" (TransactionTime)",
            $"CREATE INDEX idx_{tableName}_Location ON \"{tableName}\" (Location)",
            $"CREATE INDEX idx_{tableName}_FuelType ON \"{tableName}\" (FuelType)"
        };

        foreach (var indexCmd in indexCommands)
        {
            try
            {
                var indexCommand = connection.CreateCommand();
                indexCommand.CommandText = indexCmd;
                await indexCommand.ExecuteNonQueryAsync();
            }
            catch
            {
                /* 索引可能已存在 */
            }
        }
        
    }

    private async Task<bool> ValidateFuelCardTableStructureAsync(SqliteConnection connection, string tableName)
    {
        try
        {
            var getColumns = connection.CreateCommand();
            getColumns.CommandText = $"PRAGMA table_info(\"{tableName}\")";

            var existingColumns = new List<string>();
            using var reader = await getColumns.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var columnName = reader.GetString(1);
                existingColumns.Add(columnName);
            }

            var requiredColumns = new[]
            {
                "Id", "CardNumber", "TransactionTime", "BusinessType",
                "FuelType", "Quantity", "UnitPrice", "Amount",
                "Balance", "Location", "Operator", "CreatedAt"
            };

            var missingColumns = new List<string>();
            foreach (var requiredCol in requiredColumns)
            {
                if (!existingColumns.Contains(requiredCol, StringComparer.OrdinalIgnoreCase))
                {
                    missingColumns.Add(requiredCol);
                }
            }

            if (missingColumns.Count > 0)
            {
                return false;
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    #endregion

    #region IDisposable 实现

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                _operationSemaphore?.Dispose();
            }

            _disposed = true;
        }
    }

    #endregion
}