namespace FirmFormatter.OpenXml.Contracts;

public sealed record ValidationIssueContract(
    string Layer,
    string Code,
    string Message
);

public sealed class ValidationReportContract
{
    /// <summary>阻断性问题（如正文为空、引用损坏等）</summary>
    public List<ValidationIssueContract> Issues { get; } = [];

    /// <summary>非阻断性警告（如 OpenXmlValidator schema 排序差异）</summary>
    public List<ValidationIssueContract> Warnings { get; } = [];

    public Dictionary<string, string> Facts { get; } = new(StringComparer.OrdinalIgnoreCase);

    public bool Success => Issues.Count == 0;

    public void AddIssue(string layer, string code, string message)
    {
        Issues.Add(new ValidationIssueContract(layer, code, message));
    }

    public void AddWarning(string layer, string code, string message)
    {
        Warnings.Add(new ValidationIssueContract(layer, code, message));
    }

    public void SetFact(string key, string value)
    {
        Facts[key] = value;
    }
}
