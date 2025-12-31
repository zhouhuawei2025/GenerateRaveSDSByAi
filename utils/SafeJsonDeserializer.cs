using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;


namespace GenerateRaveSDSByAi.utils;

public class SafeJsonDeserializer
{
    /// <summary>
    /// 从AI返回的原始文本中提取JSON并反序列化为List<T>（容错AI返回的无关字符）
    /// </summary>
    /// <param name="rawText">AI返回的原始文本（可能含无关字符+JSON）</param>
    /// <param name="result">反序列化结果（失败时为null）</param>
    /// <returns>是否成功</returns>
    public static bool TryDeserializeFromAiText<T>(string rawText, out List<T> result, string path)
    {
        File.AppendAllText(path, "=============开始进行json反序列化 ......" + DateTime.Now.ToString() + "\r\n");
        result = null;
        if (string.IsNullOrWhiteSpace(rawText))
        {
            Debug.WriteLine("错误：输入的json文本为空");
            File.AppendAllText(path, " 错误：输入的json文本为空" + "\r\n");
            return false;
        }

        try
        {
            // 步骤1：预处理 - 去除首尾空白、特殊字符（换行、制表符、全角空格等）
            string cleanedText = rawText.Trim('\r', '\n', '\t', ' ', '　', '"', '\'', '`', ':', '-', '=');

            // 步骤2：正则提取JSON片段（匹配以[开头    ]结尾的JSON数组，兼容嵌套）
            // 正则说明：匹配最外层的 [ ... ]，支持嵌套的{}和[]
            var jsonRegex = new Regex(@"\[(?:[^\[\]]|(?<open>\[)|(?<-open>\]))*(?(open)(?!))\]", RegexOptions.Singleline);
            Match jsonMatch = jsonRegex.Match(cleanedText);

            if (!jsonMatch.Success)
            {
                // 兜底：尝试匹配以{开头、}结尾的JSON对象（若AI返回单对象而非数组）
                jsonRegex = new Regex(@"\{(?:[^\{\}]|(?<open>\{)|(?<-open>\}))*(?(open)(?!))\}", RegexOptions.Singleline);
                jsonMatch = jsonRegex.Match(cleanedText);
                if (!jsonMatch.Success)
                {
                    Debug.WriteLine("错误：未从文本中提取到有效JSON片段");
                    Debug.WriteLine($"原始文本（清洗后）：{cleanedText}");
                    File.AppendAllText(path, " 错误：未从文本中提取到有效JSON片段" + "\r\n");
                    File.AppendAllText(path, " 原始文本（清洗后）：{cleanedText}" + "\r\n");
                    return false;
                }
                // 若提取到单对象，包装为数组（适配List<DataDictionary>）
                cleanedText = $"[{jsonMatch.Value}]";
            }
            else
            {
                cleanedText = jsonMatch.Value;
            }

            Debug.WriteLine("提取到的纯JSON：");
            File.AppendAllText(path, "提取到的纯JSON：" + "\r\n");

            Debug.WriteLine(cleanedText);
            File.AppendAllText(path, cleanedText + "\r\n");

            Debug.WriteLine("------------------------");
            File.AppendAllText(path, "------------------------" + "\r\n");

            // 步骤3：容错反序列化
            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true, // 忽略大小写
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping, // 支持中文
                AllowTrailingCommas = true, // 容忍JSON末尾多余的逗号（AI常见错误）
                ReadCommentHandling = JsonCommentHandling.Skip // 跳过JSON中的注释（若有）
            };

            result = JsonSerializer.Deserialize<List<T>>(cleanedText, jsonOptions);
            return result != null;
        }
        catch (JsonException ex)
        {
            Debug.WriteLine($"反序列化失败：{ex.Message}");
            Debug.WriteLine($"错误位置：Line {ex.LineNumber}, Column {ex.BytePositionInLine}");
            File.AppendAllText(path, $"反序列化失败：{ex.Message}" + "\r\n");
            File.AppendAllText(path, $"错误位置：Line {ex.LineNumber}, Column {ex.BytePositionInLine}" + "\r\n");
            return false;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"处理失败：{ex.Message}");
            File.AppendAllText(path, $"处理失败：{ex.Message}" + "\r\n");

            return false;
        }
    }
}
