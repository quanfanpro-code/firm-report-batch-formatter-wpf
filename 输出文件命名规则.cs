using System.IO;
using System.Text.RegularExpressions;

namespace 事务所出报告批量排版WPF版;

public static class 输出文件命名规则
{
    private static readonly Regex 已排版文件名模式 = new(
        @"_已排版(?:\(\d+\)|_[0-9a-fA-F]{32})?\.docx$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static string 生成输出路径(string inputPath)
    {
        var dir = Path.GetDirectoryName(inputPath)!;
        var name = Path.GetFileNameWithoutExtension(inputPath);
        var path = Path.Combine(dir, $"{name}_已排版.docx");
        if (!File.Exists(path)) return path;
        for (var i = 2; i < 100; i++)
        {
            var candidate = Path.Combine(dir, $"{name}_已排版({i}).docx");
            if (!File.Exists(candidate)) return candidate;
        }
        return Path.Combine(dir, $"{name}_已排版_{Guid.NewGuid():N}.docx");
    }

    public static bool 是已排版文件(string path)
    {
        var fileName = Path.GetFileName(path);
        return 已排版文件名模式.IsMatch(fileName);
    }
}
