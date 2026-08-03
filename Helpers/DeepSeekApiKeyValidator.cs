using System;

namespace ExcelToSQLite.Helpers;

public static class DeepSeekApiKeyValidator
{
    private const int ExpectedKeyLength = 32;
    private const string Prefix = "sk-";
    
    public static bool IsValid(string apiKey)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
            return false;
        
        var trimmed = apiKey.Trim();
        
        // 1. 检查前缀
        if (!trimmed.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase))
            return false;
        
        // 2. 提取密钥部分并检查长度 (总长度应为 3 + 32 = 35)
        if (trimmed.Length != Prefix.Length + ExpectedKeyLength)
            return false;
        
        var keyPart = trimmed.Substring(Prefix.Length);
        
        // 3. (可选) 检查是否只包含合法字符 (十六进制)
        foreach (char c in keyPart)
        {
            if (!char.IsLetterOrDigit(c))
                return false;
        }
        
        return true;
    }
}