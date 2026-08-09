namespace DocxValidationTool;

public static class CoverModeArgumentParser
{
    public static bool? ParseOverride(string? text)
    {
        var normalized = text?.Trim().ToLowerInvariant();
        return normalized switch
        {
            null => null,
            "" => null,
            "auto" => null,
            "true" => true,
            "yes" => true,
            "cover" => true,
            "false" => false,
            "no" => false,
            "nocover" => false,
            _ => throw new ArgumentException($"不支持的封面模式：{text}")
        };
    }
}
