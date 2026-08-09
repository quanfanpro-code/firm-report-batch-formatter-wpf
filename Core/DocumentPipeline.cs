using System.IO;
using DocumentFormat.OpenXml.Packaging;
using FirmFormatter.OpenXml.Contracts;

namespace FirmFormatter.OpenXml.Core;

public sealed class DocumentPipeline
{
    private readonly Action<LogEventContract> _emit;

    public DocumentPipeline(Action<LogEventContract> emit)
    {
        _emit = emit;
    }

    public ResponseContract Run(RequestContract request)
    {
        // 入参基本校验：带病请求直接返回含义明确的错误码，不进入流水线
        var 校验错误 = request.校验();
        if (校验错误 is not null)
        {
            _emit(new LogEventContract("error", "pipeline", "invalid_request", 校验错误));
            return new ResponseContract(false, request.OutputPath, "invalid_request", 校验错误);
        }

        // File.Copy 自拷贝会抛异常且语义上就是覆盖源文件，必须提前拦截
        if (string.Equals(
                Path.GetFullPath(request.InputPath),
                Path.GetFullPath(request.OutputPath),
                StringComparison.OrdinalIgnoreCase))
        {
            const string 错误 = "输入路径与输出路径相同，拒绝原地覆盖源文件";
            _emit(new LogEventContract("error", "pipeline", "invalid_request", 错误));
            return new ResponseContract(false, request.OutputPath, "invalid_request", 错误);
        }

        WordprocessingDocument? word = null;
        string? 临时输出路径 = null;
        try
        {
            var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(request.OutputPath))!;
            临时输出路径 = Path.Combine(
                outputDirectory,
                $".{Path.GetFileNameWithoutExtension(request.OutputPath)}.正在排版_{Guid.NewGuid():N}.docx");

            // 先在同目录临时文件中完成全部处理和验证，成功后再原子移动成正式输出名。
            File.Copy(request.InputPath, 临时输出路径, false);
            // Copy 会连同源文件的只读属性一起复制，不重置则随后以可写方式 Open 抛 UnauthorizedAccessException
            File.SetAttributes(临时输出路径, FileAttributes.Normal);
            _emit(new LogEventContract("info", "io", "copy_done", "已复制输入文档"));

            word = WordprocessingDocument.Open(临时输出路径, true);

            // 先修复旧版兼容格式遗留问题（如 evenAndOddHeaders 误入 SectionProperties），再排版
            OpenXmlHelper.CleanupLegacySectionProperties(word);

            var classifier = new FirmDocumentClassifier();
            var context = classifier.BuildContext(word, request);

            if (context.IsPureCoverDocument)
            {
                _emit(new LogEventContract("info", "cover", "pure_cover_detected", $"纯封面文档 (<{FirmRuleProfile.Default.纯封面可见字数阈值}字): {context.VisibleTextLength}字"));

                var headerFooter = new HeaderFooterService();
                headerFooter.ApplyPureCoverLayout(word);
                _emit(new LogEventContract("info", "header_footer", "header_footer_done", "已处理纯封面页边距"));
            }
            else
            {
                _emit(new LogEventContract("info", "cover", "cover_detected", context.HasCover ? "1" : "0"));

                var headerFooter = new HeaderFooterService();
                headerFooter.Apply(word, context.HasCover);
                _emit(new LogEventContract("info", "header_footer", "header_footer_done", "已处理页眉页脚和页码"));

                // 预扫描：精准定位落款区，以便正文排版时完美避开
                var signoffService = new SignoffService();
                var signoffParagraphs = signoffService.IdentifySignoffParagraphs(word, context.HasCover);

                // [核心流水线修复]：先铺大底色（正文），再精雕细琢（表格、落款）。绝对不能把粗活放在细活后面！
                var paragraph = new ParagraphService();
                paragraph.Apply(word, context.HasCover, signoffParagraphs);
                _emit(new LogEventContract("info", "paragraph", "paragraph_done", "已处理段落与标题"));

                var table = new TableService();
                table.Apply(word, context.HasCover);
                _emit(new LogEventContract("info", "table", "table_done", "已处理表格"));

                signoffService.Apply(word, signoffParagraphs);
                _emit(new LogEventContract("info", "signoff", "signoff_done", "已处理落款区"));
            }

            // 保存前按 schema 顺序规范化样式子元素，避免 OpenXmlValidator 误报
            OpenXmlHelper.NormalizeStyleChildOrder(word);
            word.MainDocumentPart?.Document?.Save();

            var gateCheck = new GateCheckService();
            var gateResult = gateCheck.Run(word, request, context);
            if (!gateResult.Success)
            {
                throw new InvalidDataException(BuildGateFailureMessage(gateResult));
            }

            // 输出非阻断性警告到日志（如 OpenXmlValidator schema 排序差异）
            var warnings = gateResult.ValidationReport?.Warnings;
            if (warnings is { Count: > 0 })
            {
                _emit(new LogEventContract("warning", "validation", "warnings", $"OpenXmlValidator 发现 {warnings.Count} 项 schema 差异（不阻断）"));
                foreach (var w in warnings.Take(5))
                {
                    _emit(new LogEventContract("warning", "validation", "openxml_warning", $"{w.Code}：{w.Message}"));
                }
                if (warnings.Count > 5)
                {
                    _emit(new LogEventContract("warning", "validation", "warnings_truncated", $"…其余 {warnings.Count - 5} 项已省略"));
                }
            }

            _emit(new LogEventContract("info", "validation", "validation_done", $"统一门禁通过（正文可见字数：{gateResult.ValidationReport?.Facts.GetValueOrDefault("可见正文字符数", "未知")}）"));

            // OpenXML 在 Dispose 时才真正 flush 所有 part 落盘。flush 失败必须进入外层 catch 走删半成品流程，
            // 绝不能用空 catch 吞掉，否则会静默产出损坏 docx 却返回 Success=true
            word.Dispose();
            word = null;

            // Dispose 后重新打开最终落盘文件，确保包结构确实可读；schema 兼容性差异仍按用户要求只警告。
            using (var reopened = WordprocessingDocument.Open(临时输出路径, false))
            {
                _ = new ValidationService().Validate(
                    reopened,
                    context.HasCover,
                    context.IsPureCoverDocument,
                    request.ScenarioName,
                    throwOnFailure: true);
            }

            File.Move(临时输出路径, request.OutputPath, false);
            临时输出路径 = null;

            return new ResponseContract(true, request.OutputPath, null, "ok", gateResult);
        }
        catch (Exception ex)
        {
            if (word is not null)
            {
                try { word.Dispose(); }
                catch (Exception disposeEx)
                {
                    try { _emit(new LogEventContract("warning", "pipeline", "dispose_failed", $"失败路径释放文档时出错：{disposeEx.Message}")); } catch { }
                }
            }
            var failureOutputPath = request.OutputPath;
            if (!string.IsNullOrWhiteSpace(临时输出路径) && File.Exists(临时输出路径))
            {
                try
                {
                    failureOutputPath = BuildFailureOutputPath(request.OutputPath);
                    File.Move(临时输出路径, failureOutputPath, false);
                    _emit(new LogEventContract("warning", "pipeline", "failure_copy_kept", $"失败件已保留：{failureOutputPath}"));
                }
                catch (Exception moveEx)
                {
                    failureOutputPath = 临时输出路径;
                    try { _emit(new LogEventContract("warning", "pipeline", "failure_copy_move_failed", $"失败件保留在临时路径：{临时输出路径}；改名失败：{moveEx.Message}")); } catch { }
                }
            }
            try { _emit(new LogEventContract("error", "pipeline", "fatal", ex.Message)); } catch { }
            return new ResponseContract(false, failureOutputPath, "openxml_engine_failed", ex.Message);
        }
    }

    private static string BuildGateFailureMessage(Contracts.GateCheckResultContract gateResult)
    {
        if (gateResult.BlockingIssues.Count == 0)
        {
            return "统一门禁失败";
        }

        var lines = gateResult.BlockingIssues
            .Take(8)
            .Select(issue => $"[{issue.Layer}] {issue.Code}：{issue.Message}");
        return "统一门禁失败：" + Environment.NewLine + string.Join(Environment.NewLine, lines);
    }

    private static string BuildFailureOutputPath(string requestedOutputPath)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(requestedOutputPath))!;
        var baseName = Path.GetFileNameWithoutExtension(requestedOutputPath);
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var candidate = Path.Combine(directory, $"{baseName}_排版失败_{stamp}.docx");
        for (var i = 2; File.Exists(candidate); i++)
        {
            candidate = Path.Combine(directory, $"{baseName}_排版失败_{stamp}_{i}.docx");
        }
        return candidate;
    }
}
