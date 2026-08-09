namespace FirmFormatter.OpenXml.Core;

public sealed class FirmRuleProfile
{
    public static FirmRuleProfile Default { get; } = new();

    public string 中文字体 { get; init; } = "宋体";

    public string 西文字体 { get; init; } = "Times New Roman";

    public string 正文字号HalfPoint { get; init; } = "24";

    public string 页眉页脚字号HalfPoint { get; init; } = "21";

    public string 正文行距Twips { get; init; } = "420";

    public string 落款行距Twips { get; init; } = "840";

    public int 纯封面可见字数阈值 { get; init; } = 200;

    public int 页边距上Twips { get; init; } = 1417;

    public int 页边距左Twips { get; init; } = 1417;

    public int 页边距右Twips { get; init; } = 850;

    public int 普通文档下边距Twips { get; init; } = 1701;

    public int 纯封面下边距Twips { get; init; } = 1247;

    public uint 页眉距离Twips { get; init; } = 850U;

    public uint 页脚距离Twips { get; init; } = 567U;

    public int 注册会计师制表位Twips { get; init; } = 7200;
}
