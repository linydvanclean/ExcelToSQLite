using ExcelToSQLite.Models;
using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading;
using ExcelToSQLite.Services.TableDefinitions;

namespace ExcelToSQLite.Services;

public enum LoginResult
{
    Success,
    UserNotFound,
    InvalidPassword,
    Error
}

public class UserService : IDisposable
{
    private readonly DatabaseService _databaseService;
    private readonly string _tableName = TableNames.Users;
    private bool _disposed;
    private readonly SemaphoreSlim _initSemaphore = new SemaphoreSlim(1, 1);
    // 密码策略配置
    private const int MinPasswordLength = 6;

    public UserService()
    {
        _databaseService = DatabaseService.Instance;
    }
    
    /// <summary>
    /// 用户验证
    /// </summary>
    /// <param name="username"></param>
    /// <param name="password"></param>
    /// <returns></returns>
    public async Task<LoginResult> ValidateUserDetailedAsync(string username, string password)
    {
        if (string.IsNullOrWhiteSpace(username))
            return LoginResult.UserNotFound;

        if (string.IsNullOrWhiteSpace(password))
            return LoginResult.InvalidPassword;

        try
        {
            // 检查用户是否存在
            var user = await GetUserByUsernameAsync(username);
            if (user == null)
                return LoginResult.UserNotFound;

            // 验证密码
            string inputHash = HashPassword(password);
            if (user.PasswordHash != inputHash)
                return LoginResult.InvalidPassword;

            // 更新最后登录时间
            await UpdateLastLoginAsync(username);

            return LoginResult.Success;
        }
        catch
        {
            return LoginResult.Error;
        }
    }

    #region 用户管理
    /// <summary>
    /// 判断用户是否存在
    /// </summary>
    /// <param name="username"></param>
    /// <returns></returns>

    public async Task<bool> UserExistsAsync(string username)
    {
        if (string.IsNullOrWhiteSpace(username))
            return false;

        try
        {
            var user = await GetUserByUsernameAsync(username);
            return user != null;
        }
        catch
        {
            return false;
        }
    }
    /// <summary>
    /// 添加用户
    /// </summary>
    /// <param name="username"></param>
    /// <param name="password"></param>
    /// <param name="role"></param>
    /// <returns></returns>

    public async Task<bool> CreateUserAsync(string username, string password, string role = "User")
    {
        if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            return false;

        if (password.Length < MinPasswordLength)
            return false;

        try
        {
            // 检查用户是否已存在
            if (await UserExistsAsync(username))
                return false;

            string sql = $@"
                INSERT INTO {_tableName} (Username, PasswordHash, Role, CreatedAt)
                VALUES (@p0, @p1, @p2, @p3)";

            var parameters = new List<object>
            {
                username,
                HashPassword(password),
                role,
                GetCurrentTimestamp()
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(sql, parameters);
            
            if (rowsAffected > 0)
            {
                return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// 修改密码
    /// </summary>
    /// <param name="username"></param>
    /// <param name="oldPassword"></param>
    /// <param name="newPassword"></param>
    /// <returns></returns>
    public async Task<bool> ChangePasswordAsync(string username, string oldPassword, string newPassword)
    {
        if (string.IsNullOrWhiteSpace(username) || 
            string.IsNullOrWhiteSpace(oldPassword) || 
            string.IsNullOrWhiteSpace(newPassword))
        {
            return false;
        }

        if (newPassword.Length < MinPasswordLength)
            return false;

        try
        {
            // 验证旧密码
            var result = await ValidateUserDetailedAsync(username, oldPassword);
            if (result != LoginResult.Success)
                return false;

            // 更新密码
            string updateSql = $"UPDATE {_tableName} SET PasswordHash = @p0 WHERE Username = @p1";
            var parameters = new List<object>
            {
                HashPassword(newPassword),
                username
            };

            var rowsAffected = await _databaseService.ExecuteNonQueryAsync(updateSql, parameters);
            
            if (rowsAffected > 0)
            {
                return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }
    /// <summary>
    /// 根据用户名查询用户信息
    /// </summary>
    /// <param name="username"></param>
    /// <returns></returns>

    public async Task<User?> GetUserByUsernameAsync(string username)
    {
        if (string.IsNullOrWhiteSpace(username))
            return null;

        try
        {
            string sql = $"SELECT Id, Username, PasswordHash, Role, CreatedAt, LastLoginAt FROM {_tableName} WHERE Username = @p0";
            var parameters = new List<object> { username };
            var data = await _databaseService.ExecuteQueryAsync(sql, parameters);

            if (data == null || data.Count <= 1)
                return null;

            return MapRowToUser(data[1]);
        }
        catch
        {
            return null;
        }
    }
    #endregion
    
    #region 辅助方法
    /// <summary>
    /// 更新最近一次登录时间
    /// </summary>
    /// <param name="username"></param>
    private async Task UpdateLastLoginAsync(string username)
    {
        try
        {
            string sql = $"UPDATE {_tableName} SET LastLoginAt = @p0 WHERE Username = @p1";
            var parameters = new List<object>
            {
                GetCurrentTimestamp(),
                username
            };

            await _databaseService.ExecuteNonQueryAsync(sql, parameters);
        }
        catch
        {
        }
    }

    private string HashPassword(string password) => PublicEvent.HashString(password);

    private string GetCurrentTimestamp() => PublicEvent.GetCurrentTimestamp();

    private User? MapRowToUser(List<object> row)
    {
        try
        {
            int idx = 0;
            var createdAt = DateTime.Now;
            DateTime? lastLoginAt = null;

            if (row.Count > 4 && row[4] != null)
            {
                DateTime.TryParse(row[4].ToString(), out createdAt);
            }

            if (row.Count > 5 && row[5] != null && !string.IsNullOrEmpty(row[5].ToString()))
            {
                DateTime.TryParse(row[5].ToString(), out var lastLogin);
                lastLoginAt = lastLogin;
            }

            return new User
            {
                Id = Convert.ToInt32(row[idx++]),
                Username = row[idx++]?.ToString() ?? string.Empty,
                PasswordHash = row[idx++]?.ToString() ?? string.Empty,
                Role = row[idx++]?.ToString() ?? "User",
                CreatedAt = createdAt,
                LastLoginAt = lastLoginAt
            };
        }
        catch
        {
            return null;
        }
    }

    #endregion

    #region 密码策略验证

    public bool ValidatePasswordStrength(string password, out string message)
    {
        message = string.Empty;

        if (string.IsNullOrWhiteSpace(password))
        {
            message = "密码不能为空";
            return false;
        }

        if (password.Length < MinPasswordLength)
        {
            message = $"密码长度至少为 {MinPasswordLength} 位";
            return false;
        }

        // 检查密码复杂度
        bool hasDigit = false;
        bool hasLower = false;
        bool hasUpper = false;
        bool hasSpecial = false;

        foreach (char c in password)
        {
            if (char.IsDigit(c)) hasDigit = true;
            else if (char.IsLower(c)) hasLower = true;
            else if (char.IsUpper(c)) hasUpper = true;
            else if (!char.IsLetterOrDigit(c)) hasSpecial = true;
        }

        int strength = 0;
        if (hasDigit) strength++;
        if (hasLower) strength++;
        if (hasUpper) strength++;
        if (hasSpecial) strength++;

        if (strength < 2)
        {
            message = "密码至少包含数字和字母的组合";
            return false;
        }

        return true;
    }

    public int CalculatePasswordStrength(string password)
    {
        if (string.IsNullOrWhiteSpace(password) || password.Length < MinPasswordLength)
            return 0;

        int score = 0;

        // 长度评分
        if (password.Length >= 8) score++;
        if (password.Length >= 12) score++;

        // 复杂度评分
        if (System.Text.RegularExpressions.Regex.IsMatch(password, @"\d")) score++;
        if (System.Text.RegularExpressions.Regex.IsMatch(password, @"[a-z]")) score++;
        if (System.Text.RegularExpressions.Regex.IsMatch(password, @"[A-Z]")) score++;
        if (System.Text.RegularExpressions.Regex.IsMatch(password, @"[!@#$%^&*(),.?"":{}|<>]")) score++;

        // 返回1-4的强度等级
        return Math.Min(Math.Max(score - 1, 1), 4);
    }

    public string GetPasswordStrengthText(int strength)
    {
        return strength switch
        {
            1 => "弱",
            2 => "一般",
            3 => "强",
            4 => "非常强",
            _ => "未知"
        };
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
                _initSemaphore?.Dispose();
            }
            _disposed = true;
        }
    }

    #endregion
}