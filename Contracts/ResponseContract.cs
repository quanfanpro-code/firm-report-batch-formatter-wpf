namespace FirmFormatter.OpenXml.Contracts;

public sealed record ResponseContract(
    bool Success,
    string OutputPath,
    string? ErrorCode,
    string? Message,
    GateCheckResultContract? GateCheck = null
);
