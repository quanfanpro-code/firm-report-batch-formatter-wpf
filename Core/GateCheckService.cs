using DocumentFormat.OpenXml.Packaging;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class GateCheckService
{
    public GateCheckResultContract Run(WordprocessingDocument word, RequestContract request, FirmDocumentContext context)
    {
        var result = new GateCheckResultContract();

        var validationService = new ValidationService();
        var validationReport = validationService.Validate(
            word,
            context.HasCover,
            context.IsPureCoverDocument,
            context.ScenarioName,
            throwOnFailure: false);
        result.ValidationReport = validationReport;
        foreach (var issue in validationReport.Issues)
        {
            result.BlockingIssues.Add(issue);
        }

        result.Facts["场景"] = context.ScenarioName;
        result.Facts["运行来源"] = context.RunSource;
        result.Facts["纯封面判定"] = context.IsPureCoverDocument ? "1" : "0";
        result.Facts["封面判定"] = context.HasCover ? "1" : "0";
        写入复杂结构事实(result, context.复杂结构);

        if (request.RequireScenarioVerification)
        {
            var scenarioReport = new ScenarioVerificationService().Verify(word, context);
            result.ScenarioReport = scenarioReport;
            foreach (var issue in scenarioReport.Issues)
            {
                result.BlockingIssues.Add(new ValidationIssueContract("场景", "scenario_verification", issue));
            }
        }

        return result;
    }

    private static void 写入复杂结构事实(GateCheckResultContract result, 复杂结构观察结果 complex)
    {
        result.Facts["复杂结构.文本框数"] = complex.文本框数.ToString();
        result.Facts["复杂结构.脚注数"] = complex.脚注数.ToString();
        result.Facts["复杂结构.尾注数"] = complex.尾注数.ToString();
        result.Facts["复杂结构.编号重启数"] = complex.编号重启数.ToString();
        result.Facts["复杂结构.横向合并数"] = complex.横向合并数.ToString();
        result.Facts["复杂结构.纵向合并数"] = complex.纵向合并数.ToString();
        result.Facts["复杂结构.书签数"] = complex.书签数.ToString();
        result.Facts["复杂结构.超链接数"] = complex.超链接数.ToString();
        result.Facts["复杂结构.域代码数"] = complex.域代码数.ToString();
        result.Facts["复杂结构.落款复杂结构数"] = complex.落款复杂结构数.ToString();
    }
}
