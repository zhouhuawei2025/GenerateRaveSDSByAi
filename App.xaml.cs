using GenerateRaveSDSByAi.utils;
using System.Configuration;
using System.Data;
using System.Windows;

namespace GenerateRaveSDSByAi;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
    /// <summary>
    /// 全局AI配置（internal类仅程序集内可访问，符合你的定义）
    /// </summary>
    internal static AIconfig GlobalAIConfig { get; private set; }
    internal static List<string> Analytes {  get; private set; }

    /// <summary>
    /// 程序启动事件（主窗口显示前执行）
    /// </summary>
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        try
        {
            string? keyFromEnv = Environment.GetEnvironmentVariable("OPENAIAPIKEY");
            //string? keyFromEnv = Environment.GetEnvironmentVariable("QWENAPIKEY");
            // 读取配置（JSON覆盖默认值，无JSON用默认值）
            GlobalAIConfig = ConfigHelper.LoadAIConfig();

            if (keyFromEnv != null)
            {
                GlobalAIConfig.ApiKey = keyFromEnv;
            }

            Analytes = ConfigHelper.LoadAnalytesConfig();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"程序启动失败：{ex.Message}",
                "配置错误",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );
            Shutdown();
        }
    }
}
