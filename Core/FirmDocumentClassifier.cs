using DocumentFormat.OpenXml.Packaging;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class FirmDocumentClassifier
{
    private readonly FirmRuleProfile _ruleProfile = FirmRuleProfile.Default;
    private readonly FirmTemplateProfile _templateProfile = FirmTemplateProfile.Default;

    public FirmDocumentContext BuildContext(WordprocessingDocument word, RequestContract request)
    {
        var body = word.MainDocumentPart?.Document?.Body;
        var visibleTextLength = OpenXmlHelper.统计正文可见有效文本字数(body);
        var isPureCoverDocument = request.IsPureCoverOverride ?? (visibleTextLength < _ruleProfile.纯封面可见字数阈值);

        var hasCover = false;
        if (!isPureCoverDocument)
        {
            var coverService = new CoverService();
            hasCover = CoverModeResolver.ResolveHasCover(request, () => coverService.DetectCover(word));
        }

        var complex = 复杂结构观察服务.观察(word, hasCover);

        return new FirmDocumentContext
        {
            Request = request,
            VisibleTextLength = visibleTextLength,
            IsPureCoverDocument = isPureCoverDocument,
            HasCover = hasCover,
            复杂结构 = complex,
            ScenarioName = ScenarioNameResolver.Resolve(request, "常规"),
            RunSource = string.IsNullOrWhiteSpace(request.RunSource) ? "未知" : request.RunSource.Trim()
        };
    }
}
