using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirmFormatter.OpenXml.Core;

public sealed class CoverService
{
    private readonly FirmTemplateProfile _templateProfile = FirmTemplateProfile.Default;

    public bool DetectCover(WordprocessingDocument word)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        if (body is null) return false;

        var scanned = 0;
        foreach (var child in body.Elements())
        {
            scanned++;
            if (child is Table tbl)
            {
                var t = OpenXmlHelper.NormalizeText(OpenXmlHelper.提取可见文本(tbl));
                // 关键词列表为空时 All() 恒真，会把第一个表格误判为封面，必须先确认配置了关键词
                if (_templateProfile.封面识别关键词.Count > 0 &&
                    _templateProfile.封面识别关键词.All(keyword =>
                        t.Contains(keyword, StringComparison.Ordinal)))
                {
                    return true;
                }
            }
            if (scanned >= _templateProfile.封面扫描窗口元素数) break;
        }
        return false;
    }
}
