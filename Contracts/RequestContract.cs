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
        return null;
    }
}
