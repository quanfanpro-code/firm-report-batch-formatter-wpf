using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class FirmDocumentContext
{
    public required RequestContract Request { get; init; }

    public required int VisibleTextLength { get; init; }

    public required bool IsPureCoverDocument { get; init; }

    public required bool HasCover { get; init; }

    public required 复杂结构观察结果 复杂结构 { get; init; }

    public required string ScenarioName { get; init; }

    public required string RunSource { get; init; }
}
