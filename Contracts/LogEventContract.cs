namespace FirmFormatter.OpenXml.Contracts;

public sealed record LogEventContract(
    string Type,
    string Stage,
    string Code,
    string Message
);
