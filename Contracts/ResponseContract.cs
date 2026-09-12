namespace FirmFormatter.OpenXml.Contracts;

public sealed record ResponseContract(
    bool Success,
    string OutputPath,
    string? ErrorCode,
    string? Message,
    GateCheckResultContract? GateCheck = null
)
{
    public string? AuditPath { get; init; }
    public string? AuditWarning { get; init; }
}
