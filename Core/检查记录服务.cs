using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed record 批次检查项(string 原稿, ResponseContract 结果);

public static class 检查记录服务
{
    private static readonly JsonSerializerOptions 格式 = new() { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    public static string 保存单文件(RequestContract request, ResponseContract result, IReadOnlyList<LogEventContract> events)
    {
        var basis = string.IsNullOrWhiteSpace(result.OutputPath) ? request.InputPath : result.OutputPath;
        return 保存新文件(basis + ".audit.json", new
        {
            时间 = DateTimeOffset.Now,
            原稿 = request.InputPath,
            处理方式 = request.IsPureCoverOverride switch { true => "纯封面", false => "正文", _ => "自动" },
            Word实测已执行 = false,
            结果 = result,
            事件 = events
        });
    }

    public static string 保存批次(string directory, IReadOnlyList<批次检查项> items, bool cancelled)
        => 保存新文件(Path.Combine(directory, $"批次检查_{DateTime.Now:yyyyMMdd_HHmmss}.json"), new
        {
            时间 = DateTimeOffset.Now,
            已取消 = cancelled,
            成功数 = items.Count(x => x.结果.Success),
            失败数 = items.Count(x => !x.结果.Success && x.结果.ErrorCode != "cancelled"),
            取消数 = items.Count(x => x.结果.ErrorCode == "cancelled"),
            文件 = items
        });

    private static string 保存新文件<T>(string path, T value)
    {
        // 不覆盖记录；同一目标的并发竞争也只改用新名称。
        while (true)
        {
            FileStream stream;
            try { stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write); }
            catch (IOException) when (File.Exists(path))
            {
                path = Path.Combine(Path.GetDirectoryName(path)!, Path.GetFileNameWithoutExtension(path) + "_" + Guid.NewGuid().ToString("N") + ".json");
                continue;
            }
            using (stream) JsonSerializer.Serialize(stream, value, 格式);
            return path;
        }
    }
}
