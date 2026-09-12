using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;

namespace 文档快照导出测试;

public sealed class 批处理互鉴测试
{
    [Theory]
    [InlineData(30, false)]
    [InlineData(260, true)]
    public void 手动选择优先于纯封面字数判断(int length, bool pure)
    {
        using var stream = new MemoryStream();
        using var word = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        word.AddMainDocumentPart().Document = new Document(new Body(new Paragraph(new Run(new Text(new string('字', length))))));
        var request = new RequestContract();
        var property = typeof(RequestContract).GetProperty("IsPureCoverOverride");
        Assert.NotNull(property);
        property.SetValue(request, pure);
        var context = new FirmDocumentClassifier().BuildContext(word, request);
        Assert.Equal(pure, context.IsPureCoverDocument);
        Assert.False(context.HasCover);
        property.SetValue(request, null);
        Assert.Equal(length < 200, new FirmDocumentClassifier().BuildContext(word, request).IsPureCoverDocument);
    }

    [Theory]
    [InlineData("成功")]
    [InlineData("失败")]
    [InlineData("取消")]
    public void 每次处理保存真实结果且不覆盖旧记录(string scenario)
    {
        var dir = Path.Combine(Path.GetTempPath(), "事务所记录_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var input = Path.Combine(dir, "原稿.docx");
        using (var doc = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            doc.AddMainDocumentPart().Document = new Document(new Body(new Paragraph(new Run(new Text("测试封面"))), new SectionProperties()));
        var original = File.ReadAllBytes(input);
        var output = Path.Combine(dir, "成稿.docx");
        if (scenario == "失败") File.WriteAllText(output, "已有文件不得覆盖");
        using var cts = new CancellationTokenSource();
        if (scenario == "取消") cts.Cancel();
        var request = new RequestContract { InputPath = input, OutputPath = output, CancellationToken = cts.Token };
        var result = new DocumentPipeline(_ => { }).Run(request);
        Assert.Equal(scenario == "成功", result.Success);
        var auditProperty = typeof(ResponseContract).GetProperty("AuditPath");
        Assert.NotNull(auditProperty);
        var audit = (string?)auditProperty.GetValue(result);
        Assert.True(File.Exists(audit));
        using var json = JsonDocument.Parse(File.ReadAllText(audit!));
        Assert.Equal(result.Success, json.RootElement.GetProperty("结果").GetProperty("Success").GetBoolean());
        Assert.False(json.RootElement.GetProperty("Word实测已执行").GetBoolean());
        var old = File.ReadAllBytes(audit!);
        var again = new DocumentPipeline(_ => { }).Run(request);
        Assert.NotEqual(audit, auditProperty.GetValue(again));
        Assert.Equal(old, File.ReadAllBytes(audit!));
        Assert.Equal(original, File.ReadAllBytes(input));
        if (scenario == "失败") Assert.Equal("已有文件不得覆盖", File.ReadAllText(output));
        if (scenario == "取消") Assert.False(File.Exists(output));
    }
}
