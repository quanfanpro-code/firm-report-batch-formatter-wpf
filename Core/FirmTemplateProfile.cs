namespace FirmFormatter.OpenXml.Core;

public sealed class FirmTemplateProfile
{
    public static FirmTemplateProfile Default { get; } = new();

    public IReadOnlyList<string> 封面识别关键词 { get; init; } = ["会计师事务所", "地址", "电话", "传真"];

    public int 封面扫描窗口元素数 { get; init; } = 60;

    public string 事务所落款头部 { get; init; } = "四川华信(集团)会计师事务所";

    public string 落款城市行 { get; init; } = "中国·成都";
}
