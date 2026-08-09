using DocxValidationTool;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Xunit;
using DocumentFormat.OpenXml.Packaging;

namespace 文档快照导出测试;

public sealed class 上下文漂移测试
{
    [Theory]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("auto", null)]
    [InlineData("true", true)]
    [InlineData("yes", true)]
    [InlineData("false", false)]
    [InlineData("no", false)]
    public void 解析封面模式参数时_应得到预期覆盖值(string? text, bool? expected)
    {
        var actual = CoverModeArgumentParser.ParseOverride(text);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void 显式禁止封面时_不应再调用自动探测()
    {
        var request = new RequestContract
        {
            HasCoverOverride = false
        };

        var detectorCalled = false;
        var actual = CoverModeResolver.ResolveHasCover(request, () =>
        {
            detectorCalled = true;
            return true;
        });

        Assert.False(actual);
        Assert.False(detectorCalled);
    }

    [Fact]
    public void 未显式覆盖时_应继续走自动探测()
    {
        var request = new RequestContract
        {
            HasCoverOverride = null
        };

        var detectorCalled = false;
        var actual = CoverModeResolver.ResolveHasCover(request, () =>
        {
            detectorCalled = true;
            return true;
        });

        Assert.True(actual);
        Assert.True(detectorCalled);
    }

    [Fact]
    public void 显式指定有封面时_不应再调用自动探测()
    {
        var request = new RequestContract
        {
            HasCoverOverride = true
        };

        var detectorCalled = false;
        var actual = CoverModeResolver.ResolveHasCover(request, () =>
        {
            detectorCalled = true;
            return false;
        });

        Assert.True(actual);
        Assert.False(detectorCalled);
    }

    [Fact]
    public void 显式场景名存在时_应优先使用显式值()
    {
        var request = new RequestContract
        {
            ScenarioName = "矩阵-封面"
        };

        var actual = ScenarioNameResolver.Resolve(request, "常规");
        Assert.Equal("矩阵-封面", actual);
    }

    [Fact]
    public void 显式场景名为空时_应回落到默认值()
    {
        var request = new RequestContract
        {
            ScenarioName = " "
        };

        var actual = ScenarioNameResolver.Resolve(request, "常规");
        Assert.Equal("常规", actual);
    }

    [Fact]
    public void 验证器上下文推导_应识别封面矩阵为有封面()
    {
        var path = GetMatrixOutputPath("封面矩阵_已排版.docx");
        using var word = WordprocessingDocument.Open(path, false);

        var context = new FirmDocumentClassifier().BuildContext(word, new RequestContract
        {
            InputPath = path,
            OutputPath = path
        });

        Assert.True(context.HasCover);
        Assert.False(context.IsPureCoverDocument);
    }

    [Fact]
    public void 验证器上下文推导_应识别纯封面矩阵为纯封面()
    {
        var path = GetMatrixOutputPath("纯封面矩阵_已排版.docx");
        using var word = WordprocessingDocument.Open(path, false);

        var context = new FirmDocumentClassifier().BuildContext(word, new RequestContract
        {
            InputPath = path,
            OutputPath = path
        });

        Assert.False(context.HasCover);
        Assert.True(context.IsPureCoverDocument);
    }

    [Fact]
    public void 验证器上下文推导_普通矩阵不应误判为封面()
    {
        var path = GetMatrixOutputPath("标题矩阵_已排版.docx");
        using var word = WordprocessingDocument.Open(path, false);

        var context = new FirmDocumentClassifier().BuildContext(word, new RequestContract
        {
            InputPath = path,
            OutputPath = path
        });

        Assert.False(context.HasCover);
        Assert.False(context.IsPureCoverDocument);
    }

    private static string GetMatrixOutputPath(string fileName)
    {
        return 矩阵测试资料.获取输出(fileName);
    }
}
