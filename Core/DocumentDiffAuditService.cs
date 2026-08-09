using System.Text.Json;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class DocumentDiffAuditService
{
    private static readonly HashSet<string> 忽略路径 = new(StringComparer.Ordinal)
    {
        "$.文档路径"
    };

    public SnapshotDiffReportContract CompareDocuments(string leftDocxPath, string rightDocxPath)
    {
        var snapshotService = new DocumentSnapshotService();
        var leftJson = snapshotService.ExportAsJson(leftDocxPath);
        var rightJson = snapshotService.ExportAsJson(rightDocxPath);
        return CompareJson(leftJson, rightJson);
    }

    public SnapshotDiffReportContract CompareJson(string leftJson, string rightJson)
    {
        using var left = JsonDocument.Parse(leftJson);
        using var right = JsonDocument.Parse(rightJson);

        var report = new SnapshotDiffReportContract();
        CompareElement("$", left.RootElement, right.RootElement, report);
        BuildAreaSummary(report);
        return report;
    }

    private static void CompareElement(string path, JsonElement left, JsonElement right, SnapshotDiffReportContract report)
    {
        if (忽略路径.Contains(path))
        {
            return;
        }

        if (left.ValueKind != right.ValueKind)
        {
            report.差异列表.Add(new SnapshotDifferenceContract(path, left.ValueKind.ToString(), right.ValueKind.ToString()));
            return;
        }

        switch (left.ValueKind)
        {
            case JsonValueKind.Object:
                CompareObject(path, left, right, report);
                return;
            case JsonValueKind.Array:
                CompareArray(path, left, right, report);
                return;
            default:
                var leftText = left.ToString();
                var rightText = right.ToString();
                if (!string.Equals(leftText, rightText, StringComparison.Ordinal))
                {
                    report.差异列表.Add(new SnapshotDifferenceContract(path, leftText, rightText));
                }
                return;
        }
    }

    private static void CompareObject(string path, JsonElement left, JsonElement right, SnapshotDiffReportContract report)
    {
        // 手动建字典，重复属性名以后值覆盖先值，避免 ToDictionary 抛 ArgumentException
        var leftProperties = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var property in left.EnumerateObject())
        {
            leftProperties[property.Name] = property.Value;
        }

        var rightProperties = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var property in right.EnumerateObject())
        {
            rightProperties[property.Name] = property.Value;
        }
        var names = leftProperties.Keys.Union(rightProperties.Keys, StringComparer.Ordinal).OrderBy(name => name, StringComparer.Ordinal);

        foreach (var name in names)
        {
            var childPath = $"{path}.{name}";
            var hasLeft = leftProperties.TryGetValue(name, out var leftValue);
            var hasRight = rightProperties.TryGetValue(name, out var rightValue);

            if (!hasLeft || !hasRight)
            {
                report.差异列表.Add(new SnapshotDifferenceContract(
                    childPath,
                    hasLeft ? leftValue.ToString() : "<缺失>",
                    hasRight ? rightValue.ToString() : "<缺失>"));
                continue;
            }

            CompareElement(childPath, leftValue, rightValue, report);
        }
    }

    private static void CompareArray(string path, JsonElement left, JsonElement right, SnapshotDiffReportContract report)
    {
        var leftItems = left.EnumerateArray().ToList();
        var rightItems = right.EnumerateArray().ToList();
        var count = Math.Max(leftItems.Count, rightItems.Count);

        for (var i = 0; i < count; i++)
        {
            var childPath = $"{path}[{i}]";
            if (i >= leftItems.Count || i >= rightItems.Count)
            {
                report.差异列表.Add(new SnapshotDifferenceContract(
                    childPath,
                    i < leftItems.Count ? leftItems[i].ToString() : "<缺失>",
                    i < rightItems.Count ? rightItems[i].ToString() : "<缺失>"));
                continue;
            }

            CompareElement(childPath, leftItems[i], rightItems[i], report);
        }
    }

    private static void BuildAreaSummary(SnapshotDiffReportContract report)
    {
        report.区域差异数.Clear();

        foreach (var item in report.差异列表)
        {
            var area = ResolveArea(item.路径);
            report.区域差异数.TryGetValue(area, out var count);
            report.区域差异数[area] = count + 1;
        }
    }

    private static string ResolveArea(string path)
    {
        if (path.Contains(".复杂结构", StringComparison.Ordinal))
        {
            return "复杂结构";
        }

        if (path.Contains(".表格", StringComparison.Ordinal))
        {
            return "表格";
        }

        if (path.Contains(".分节", StringComparison.Ordinal))
        {
            return "分节";
        }

        if (path.Contains(".标题", StringComparison.Ordinal))
        {
            return "标题";
        }

        if (path.Contains(".落款", StringComparison.Ordinal))
        {
            return "落款";
        }

        if (path.Contains(".正文", StringComparison.Ordinal) || path.Contains(".段落", StringComparison.Ordinal))
        {
            return "正文";
        }

        return "整体";
    }
}
