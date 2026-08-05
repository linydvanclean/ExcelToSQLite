using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelToSQLite.Models;

namespace ExcelToSQLite.Services
{
    /// <summary>
    /// 应用配置服务 - 管理 JSON 配置文件
    /// </summary>
    public class AppConfigService
    {
        private string _configPath;
        private string _configFile;
        private AppConfig? _config;

        public AppConfigService()
        {
            var baseDirectory = AppContext.BaseDirectory;
            string configDirectory;
            
            if (HasWritePermission(baseDirectory))
            {
                // 如果有权限，使用程序目录下的 settings 子目录
                configDirectory = Path.Combine(baseDirectory, "settings");
            }
            else
            {
                // 否则使用用户目录
                var userHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                var appName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name ?? "ExcelToSQLite";
                configDirectory = Path.Combine(userHome, ".local", "share", appName, "settings");
            }
            
            _configPath = configDirectory;
            _configFile = Path.Combine(_configPath, "appconfig.json");

            // =============== 🟢 以下是新增的代码 ===============
            try 
            {
                // 1. 创建 settings 文件夹（如果不存在）
                if (!Directory.Exists(_configPath))
                {
                    Directory.CreateDirectory(_configPath);
                }

                // 2. 如果配置文件不存在，创建一个默认的空 JSON 文件（或写入默认配置）
                if (!File.Exists(_configFile))
                {
                    // 方法A：创建一个空文件
                    using (File.Create(_configFile)) { }

                    // 方法B（推荐）：写入一个基础的空 JSON 对象，这样你读取时不会报错
                    // File.WriteAllText(_configFile, "{}");
                }
            }
            catch
            {
                // 兜底方案：如果上述目录创建都失败了（权限极其受限），最终会跑到临时目录
                
                string tempConfigDir = Path.Combine(Path.GetTempPath(), "ExcelToSQLite", "settings");
                if (!Directory.Exists(tempConfigDir))
                {
                    Directory.CreateDirectory(tempConfigDir);
                }
                _configPath = tempConfigDir;
                _configFile = Path.Combine(_configPath, "appconfig.json");
                
                // 同样在临时目录创建文件
                if (!File.Exists(_configFile))
                {
                    using (File.Create(_configFile)) { }
                }
            }
            // ====================================================
        }

        /// <summary>
        /// 检查是否有写入权限
        /// </summary>
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

        /// <summary>
        /// 获取配置（如果文件不存在则创建默认配置）
        /// </summary>
        public async Task<AppConfig> GetConfigAsync()
        {
            if (_config != null)
                return _config;

            EnsureDirectoryExists();

            if (File.Exists(_configFile))
            {
                try
                {
                    var json = await File.ReadAllTextAsync(_configFile);
                    _config = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                    return _config;
                }
                catch
                {
                    // 读取失败则创建默认配置
                    _config = CreateDefaultConfig();
                    await SaveConfigAsync(_config);
                    return _config;
                }
            }
            else
            {
                // 创建默认配置
                _config = CreateDefaultConfig();
                await SaveConfigAsync(_config);
                return _config;
            }
        }

        /// <summary>
        /// 保存配置
        /// </summary>
        public async Task SaveConfigAsync(AppConfig config)
        {
            _config = config;
            EnsureDirectoryExists();
            
            var json = JsonSerializer.Serialize(config, new JsonSerializerOptions 
            { 
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
            await File.WriteAllTextAsync(_configFile, json);
        }

        /// <summary>
        /// 获取指定键的值
        /// </summary>
        public async Task<T?> GetValueAsync<T>(string key, T? defaultValue = default)
        {
            var config = await GetConfigAsync();
            if (config.Values.TryGetValue(key, out var value))
            {
                try
                {
                    return JsonSerializer.Deserialize<T>(value.ToString() ?? string.Empty);
                }
                catch
                {
                    return defaultValue;
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// 设置指定键的值
        /// </summary>
        public async Task SetValueAsync<T>(string key, T value)
        {
            var config = await GetConfigAsync();
            config.Values[key] = JsonSerializer.Serialize(value);
            await SaveConfigAsync(config);
        }

        /// <summary>
        /// 获取分类列表
        /// </summary>
        public async Task<List<string>> GetCategoriesAsync()
        {
            var config = await GetConfigAsync();
    
            // 如果 Categories 为空或 null，返回默认分类
            if (config.Categories == null || config.Categories.Count == 0)
            {
                var defaultCategories = GetDefaultCategories();
                // 更新配置文件，确保下次有数据
                config.Categories = defaultCategories;
                await SaveConfigAsync(config);
                return defaultCategories;
            }
    
            return config.Categories;
        }

        /// <summary>
        /// 更新分类列表
        /// </summary>
        public async Task UpdateCategoriesAsync(List<string> categories)
        {
            var config = await GetConfigAsync();
            config.Categories = categories ?? GetDefaultCategories();
            config.LastUpdated = DateTime.Now;
            await SaveConfigAsync(config);
        }

        private void EnsureDirectoryExists()
        {
            if (!Directory.Exists(_configPath))
            {
                try
                {
                    Directory.CreateDirectory(_configPath);
                }
                catch
                {
                    // 如果创建失败，使用临时目录
                    var tempPath = Path.Combine(Path.GetTempPath(), "ExcelToSQLite", "settings");
                    if (!Directory.Exists(tempPath))
                    {
                        Directory.CreateDirectory(tempPath);
                    }
                    // 更新路径（不再使用 readonly）
                    _configPath = tempPath;
                    _configFile = Path.Combine(_configPath, "appconfig.json");
                }
            }
        }

        private AppConfig CreateDefaultConfig()
        {
            return new AppConfig
            {
                SystemName = "智慧监督数据汇集分析平台",
                Version = PublicEvent.Version,
                LastUpdated = DateTime.Now,
                Categories = GetDefaultCategories(),
                Values = new Dictionary<string, object>
                {
                    { "AppName", "ExcelToSQLite" },
                    { "Theme", "Light" },
                    { "Language", "zh-CN" },
                    { "AutoBackup", true },
                    { "BackupInterval", 7 }
                }
            };
        }

        public List<string> GetDefaultCategories()
        {
            return new List<string>
            {
                "税收风险",
                "执法风险",
                "廉政风险",
                "主体责任",
                "个体监督",
                "其他监督"
            };
        }
    }    
}