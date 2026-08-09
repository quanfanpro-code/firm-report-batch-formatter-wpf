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
        var 已创建输出 = false;
        try
        {
            // 上层 输出文件命名规则 已保证输出名不冲突；overwrite:false 让"检查存在"到"复制"之间的竞态快速失败，而不是覆盖他人文件
            File.Copy(request.InputPath, request.OutputPath, false);
            已创建输出 = true;
            // Copy 会连同源文件的只读属性一起复制，不重置则随后以可写方式 Open 抛 UnauthorizedAccessException
            File.SetAttributes(request.OutputPath, FileAttributes.Normal);
            _emit(new LogEventContract("info", "io", "copy_done", "已复制输入文档"));

            word = WordprocessingDocument.Open(request.OutputPath, true);

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
            // 只删除本次流水线自己创建的半成品；若 Copy 因文件名冲突失败，绝不能误删他人的既有文件
            if (已创建输出)
            {
                try { if (File.Exists(request.OutputPath)) File.Delete(request.OutputPath); }
                catch (Exception deleteEx)
                {
                    try { _emit(new LogEventContract("warning", "pipeline", "delete_failed", $"清理半成品文件失败，请手动删除 {request.OutputPath}：{deleteEx.Message}")); } catch { }
                }
            }
            try { _emit(new LogEventContract("error", "pipeline", "fatal", ex.Message)); } catch { }
            return new ResponseContract(false, request.OutputPath, "openxml_engine_failed", ex.Message);
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
}
