using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateRaveSDSByAi.utils
{
    public static class ConfigHelper
    {
        /// <summary>
        /// 读取aiConfig.json并初始化AIconfig（JSON值覆盖默认值，无JSON则使用默认值）
        /// </summary>
        /// <returns>初始化后的AIconfig实例</returns>
        public static AIconfig LoadAIConfig()
        {
            AIconfig config = new AIconfig();
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"配置\aiConfig.json");

            try
            {
                // 若配置文件存在，读取并覆盖默认值
                if (File.Exists(configPath))
                {
                    string jsonContent = File.ReadAllText(configPath);
                    // 反序列化时覆盖已有实例的属性（保留未在JSON中定义的默认值）
                    JsonConvert.PopulateObject(jsonContent, config);
                }
                // 验证核心配置（ApiKey无默认值，需确保非空）
                if (string.IsNullOrWhiteSpace(config.ApiKey))
                {
                    throw new InvalidDataException("AI配置中的ApiKey不能为空（JSON文件中未配置或配置为空）");
                }
                return config;
            }
            catch (Exception ex)
            {
                throw new Exception($"读取AI配置失败：{ex.Message}", ex);
            }
        }

        public static List<string> LoadAnalytesConfig()
        {            
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"配置\analytes.json");
            List<string> analytes = new();
            try
            {
                // 若配置文件存在，读取并覆盖默认值
                if (File.Exists(configPath))
                {
                    string jsonContent = File.ReadAllText(configPath);
                    // 反序列化时覆盖已有实例的属性（保留未在JSON中定义的默认值）
                    var fileAnalytes = JsonConvert.DeserializeObject<List<string>>(jsonContent);
                    if (fileAnalytes != null && fileAnalytes.Count > 0)
                    {
                        analytes = fileAnalytes; 
                    }
                }
                
                return analytes;
            }
            catch (Exception ex)
            {
                throw new Exception($"读取分析物配置失败：{ex.Message}", ex);
            }
        }
    }
}
