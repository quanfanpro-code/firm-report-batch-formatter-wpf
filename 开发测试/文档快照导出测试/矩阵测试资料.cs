using DocxValidationTool;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;

namespace 文档快照导出测试;

internal static class 矩阵测试资料
{
    private static readonly object 同步锁 = new();
    private static readonly Lazy<string> 根目录 = new(() =>
    {
        var path = Path.Combine(Path.GetTempPath(), $"事务所排版矩阵_{Guid.NewGuid():N}");
        MatrixSampleBuilder.GenerateAll(path);
        return path;
    });

    public static string 获取输出(string fileName)
    {
        lock (同步锁)
        {
            var outputPath = Path.Combine(根目录.Value, "Output", fileName);
            if (File.Exists(outputPath)) return outputPath;

            var inputName = fileName.Replace("_已排版.docx", ".docx", StringComparison.Ordinal);
            var inputPath = Path.Combine(根目录.Value, "Input", inputName);
            var result = new DocumentPipeline(_ => { }).Run(new RequestContract
            {
                InputPath = inputPath,
                OutputPath = outputPath,
                ScenarioName = Path.GetFileNameWithoutExtension(inputPath),
                RunSource = "自动测试"
            });

            return result.Success
                ? outputPath
                : throw new InvalidOperationException($"生成矩阵输出失败：{fileName}；{result.Message}");
        }
    }
}
