using System.IO;

namespace FirmFormatter.OpenXml.Contracts;

public sealed class RequestContract
{
    public string InputPath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public bool? HasCoverOverride { get; set; }
    public string? ScenarioName { get; set; }
    public string? RunSource { get; set; }
    public bool RequireScenarioVerification { get; set; }
    [System.Text.Json.Serialization.JsonIgnore]
    public CancellationToken CancellationToken { get; set; }

    /// <summary>
    /// 基本入参校验。返回 null 表示通过，否则返回可直接展示给用户的错误信息。
    /// </summary>
    public string? 校验()
    {
        if (string.IsNullOrWhiteSpace(InputPath))
        {
            return "请求参数不完整：InputPath 不能为空";
        }
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            return "请求参数不完整：OutputPath 不能为空";
        }
        if (!File.Exists(InputPath))
        {
            return $"输入文件不存在：{InputPath}";
        }
        if (!string.Equals(Path.GetExtension(InputPath), ".docx", StringComparison.OrdinalIgnoreCase))
        {
            return "输入文件必须是 .docx 格式";
        }
        if (!string.Equals(Path.GetExtension(OutputPath), ".docx", StringComparison.OrdinalIgnoreCase))
        {
            return "输出文件必须是 .docx 格式";
        }
        string? outputDirectory;
        try
        {
            outputDirectory = Path.GetDirectoryName(Path.GetFullPath(OutputPath));
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return $"输出路径无效：{ex.Message}";
        }
        if (string.IsNullOrWhiteSpace(outputDirectory) || !Directory.Exists(outputDirectory))
        {
            return $"输出目录不存在：{outputDirectory}";
        }
        if (File.Exists(OutputPath))
        {
            return $"输出文件已存在，拒绝覆盖：{OutputPath}";
        }
        return null;
    }
}
