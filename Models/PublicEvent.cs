using System;
using System.Security.Cryptography;
using System.Text;
using System.Collections.Generic;

namespace ExcelToSQLite.Models;

public class PublicEvent
{
    public const string Version = "1.1";
    public const string DeepseekBaseUrl = "https://api.deepseek.com/v1/chat/completions";
    public const string DeepseekDefaultModel = "deepseek-v4-flash";//""deepseek-v4-pro";
    
    /// <summary>
    /// 格式化处理文件名
    /// </summary>
    /// <param name="tablePrefix">批次前缀</param>
    /// <param name="fileName">文件名</param>
    /// <param name="dictionaryTableName">数据字典表名（可选）</param>
    /// <returns>格式化后的表名</returns>
    public static string GetFormatFilename(string tablePrefix, string fileName, string? dictionaryTableName = null)
    {
        string generatedName;
        
        // 如果提供了字典表名且不为空，使用字典表名
        if (!string.IsNullOrEmpty(dictionaryTableName))
        {
            generatedName = $"{tablePrefix}{dictionaryTableName}";
        }
        else
        {
            // 使用文件名生成
            generatedName = $"{tablePrefix}{fileName}";
        }
        
        // 替换特殊字符
        generatedName = generatedName.Replace(" ", "_").Replace("-", "_");
        
        // 移除可能的多余下划线
        while (generatedName.Contains("__"))
        {
            generatedName = generatedName.Replace("__", "_");
        }
        
        // 确保表名以字母开头
        if (!string.IsNullOrEmpty(generatedName) && !char.IsLetter(generatedName[0]))
        {
            generatedName = "T_" + generatedName;
        }
        
        // 截断过长的表名（SQLite 限制）
        if (generatedName.Length > 60)
        {
            generatedName = generatedName.Substring(0, 60);
        }
        
        return generatedName;
    }
    
    /// <summary>
    /// Hash string
    /// </summary>
    /// <param name="parameters"></param>
    /// <returns></returns>
    public static string HashString(string parameters)
    {
        if (string.IsNullOrEmpty(parameters))
            return string.Empty;

        using var sha256 = SHA256.Create();
        var bytes = Encoding.UTF8.GetBytes(parameters);
        var hash = sha256.ComputeHash(bytes);
        return Convert.ToBase64String(hash);
    }
    
    public static string GetCurrentTimestamp()
    {
        return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    }
}