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
        检查复杂结构是否丢失(result, context.复杂结构, 复杂结构观察服务.观察(word, context.HasCover));

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

    private static void 检查复杂结构是否丢失(
        GateCheckResultContract result,
        复杂结构观察结果 before,
        复杂结构观察结果 after)
    {
        var losses = new List<string>();
        检查计数("文本框", before.文本框数, after.文本框数);
        检查计数("脚注", before.脚注数, after.脚注数);
        检查计数("尾注", before.尾注数, after.尾注数);
        检查计数("编号重启", before.编号重启数, after.编号重启数);
        检查计数("横向合并", before.横向合并数, after.横向合并数);
        检查计数("纵向合并", before.纵向合并数, after.纵向合并数);
        检查计数("书签", before.书签数, after.书签数);
        检查计数("超链接", before.超链接数, after.超链接数);
        检查计数("域代码", before.域代码数, after.域代码数);

        if (before.文本框内容指纹 != after.文本框内容指纹) losses.Add("文本框内容发生变化");
        if (before.脚注尾注内容指纹 != after.脚注尾注内容指纹) losses.Add("脚注或尾注内容发生变化");

        if (losses.Count > 0)
        {
            result.BlockingIssues.Add(new ValidationIssueContract(
                "业务",
                "complex_structure_loss",
                $"排版前后复杂结构不一致：{string.Join("；", losses)}"));
        }

        void 检查计数(string name, int original, int current)
        {
            if (current < original) losses.Add($"{name}从 {original} 减少为 {current}");
        }
    }
}
