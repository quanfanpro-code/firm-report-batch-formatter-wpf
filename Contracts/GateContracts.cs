namespace FirmFormatter.OpenXml.Contracts;

public sealed record SnapshotDifferenceContract(string 路径, string 左值, string 右值);

public sealed class SnapshotDiffReportContract
{
    public List<SnapshotDifferenceContract> 差异列表 { get; } = [];

    public Dictionary<string, int> 区域差异数 { get; } = new(StringComparer.Ordinal);
}

public sealed class ScenarioVerificationReportContract
{
    public string Scenario { get; init; } = string.Empty;

    public List<string> Issues { get; } = [];

    public Dictionary<string, string> Facts { get; } = new(StringComparer.OrdinalIgnoreCase);

    public bool Success => Issues.Count == 0;
}

public sealed class GateCheckResultContract
{
    public ValidationReportContract? ValidationReport { get; set; }

    public ScenarioVerificationReportContract? ScenarioReport { get; set; }

    public List<ValidationIssueContract> BlockingIssues { get; } = [];

    public Dictionary<string, string> Facts { get; } = new(StringComparer.OrdinalIgnoreCase);

    public bool Success => BlockingIssues.Count == 0;
}
