using DocumentFormat.OpenXml.Packaging;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using DocxValidationTool;
using System.Text.Json;

if (args.Length == 0)
{
    PrintUsage();
    return 2;
}

try
{
    var command = args[0].Trim().ToLowerInvariant();
    return command switch
    {
        "generate-matrix" => GenerateMatrix(args),
        "export-snapshot" => ExportSnapshot(args),
        "compare-snapshot" => CompareSnapshot(args),
        "run-pipeline" => RunPipeline(args),
        "verify-doc" => VerifyDoc(args),
        "run-matrix" => RunMatrix(args),
        "run-real-sample" => RunRealSample(args),
        _ => PrintUsageAndReturn()
    };
}
catch (Exception ex)
{
    Console.Error.WriteLine($"ERROR {ex.Message}");
    return 1;
}

static int GenerateMatrix(string[] args)
{
    if (args.Length < 2)
    {
        Console.Error.WriteLine("用法：DocxValidation generate-matrix <矩阵目录>");
        return 2;
    }

    MatrixSampleBuilder.GenerateAll(args[1]);
    Console.WriteLine($"RESULT Generated={args[1]}");
    return 0;
}

static int ExportSnapshot(string[] args)
{
    if (args.Length < 3)
    {
        Console.Error.WriteLine("用法：DocxValidation export-snapshot <输入docx> <输出json>");
        return 2;
    }

    var json = new DocumentSnapshotService().ExportAsJson(args[1]);
    File.WriteAllText(args[2], json);
    Console.WriteLine($"RESULT Snapshot={args[2]}");
    return 0;
}

static int CompareSnapshot(string[] args)
{
    if (args.Length < 4)
    {
        Console.Error.WriteLine("用法：DocxValidation compare-snapshot <左侧docx> <右侧docx> <输出json>");
        return 2;
    }

    var leftJson = new DocumentSnapshotService().ExportAsJson(args[1]);
    var rightJson = new DocumentSnapshotService().ExportAsJson(args[2]);
    var report = new DocumentDiffAuditService().CompareJson(leftJson, rightJson);
    var json = JsonSerializer.Serialize(report, new JsonSerializerOptions
    {
        WriteIndented = true
    });
    File.WriteAllText(args[3], json);
    Console.WriteLine($"RESULT DiffSnapshot={args[3]}");
    return 0;
}

static int RunPipeline(string[] args)
{
    if (args.Length < 3)
    {
        Console.Error.WriteLine("用法：DocxValidation run-pipeline <输入docx> <输出docx> [封面模式:auto|true|false] [场景名]");
        return 2;
    }

    var request = new RequestContract
    {
        InputPath = args[1],
        OutputPath = args[2],
        HasCoverOverride = args.Length >= 4 ? CoverModeArgumentParser.ParseOverride(args[3]) : null,
        ScenarioName = args.Length >= 5 ? args[4] : "命令行",
        RunSource = "命令行"
    };

    var pipeline = new DocumentPipeline(EmitLog);
    var result = pipeline.Run(request);
    Console.WriteLine($"RESULT Success={result.Success} Output={result.OutputPath} Error={result.ErrorCode} Message={result.Message}");
    return result.Success ? 0 : 1;
}

static int VerifyDoc(string[] args)
{
    if (args.Length < 3)
    {
        Console.Error.WriteLine("用法：DocxValidation verify-doc <docx路径> <场景名>");
        return 2;
    }

    using var word = WordprocessingDocument.Open(args[1], false);
    var request = new RequestContract
    {
        InputPath = args[1],
        OutputPath = args[1],
        ScenarioName = args[2],
        RunSource = "验证命令"
    };
    var context = new FirmDocumentClassifier().BuildContext(word, request);
    var report = new ScenarioVerificationService().Verify(word, context);
    PrintScenarioVerificationReport(report);
    return report.Success ? 0 : 1;
}

static int RunMatrix(string[] args)
{
    if (args.Length < 2)
    {
        Console.Error.WriteLine("用法：DocxValidation run-matrix <矩阵目录>");
        return 2;
    }

    var matrixRoot = args[1];
    MatrixSampleBuilder.GenerateAll(matrixRoot);

    var inputDir = Path.Combine(matrixRoot, "Input");
    var outputDir = Path.Combine(matrixRoot, "Output");
    Directory.CreateDirectory(outputDir);

    var inputFiles = Directory.GetFiles(inputDir, "*.docx", SearchOption.TopDirectoryOnly)
        .OrderBy(Path.GetFileNameWithoutExtension, StringComparer.OrdinalIgnoreCase)
        .ToList();

    if (inputFiles.Count == 0)
    {
        Console.Error.WriteLine("矩阵目录里没有可执行的输入样本。");
        return 1;
    }

    var failed = 0;

    foreach (var inputFile in inputFiles)
    {
        var baseName = Path.GetFileNameWithoutExtension(inputFile);
        var outputFile = Path.Combine(outputDir, $"{baseName}_已排版.docx");
        var request = new RequestContract
        {
            InputPath = inputFile,
            OutputPath = outputFile,
            HasCoverOverride = null,
            ScenarioName = baseName,
            RunSource = "矩阵",
            RequireScenarioVerification = ScenarioVerificationService.支持场景(baseName)
        };

        Console.WriteLine($"MATRIX Running={baseName}");
        var pipeline = new DocumentPipeline(EmitLog);
        var result = pipeline.Run(request);
        if (!result.Success)
        {
            failed++;
            Console.WriteLine($"MATRIX Failed={baseName} Cause={result.Message}");
            continue;
        }

        if (result.GateCheck?.ScenarioReport != null)
        {
            PrintScenarioVerificationReport(result.GateCheck.ScenarioReport);
        }

        if (result.GateCheck is null || !result.GateCheck.Success)
        {
            failed++;
        }
    }

    Console.WriteLine($"RESULT MatrixFailed={failed}");
    return failed == 0 ? 0 : 1;
}

static int RunRealSample(string[] args)
{
    if (args.Length < 3)
    {
        Console.Error.WriteLine("用法：DocxValidation run-real-sample <输入docx> <输出docx>");
        return 2;
    }

    var pipeline = new DocumentPipeline(EmitLog);
    var request = new RequestContract
    {
        InputPath = args[1],
        OutputPath = args[2],
        HasCoverOverride = null,
        ScenarioName = "真实样本",
        RunSource = "真实样本",
        RequireScenarioVerification = true
    };

    var result = pipeline.Run(request);
    Console.WriteLine($"RESULT Success={result.Success} Output={result.OutputPath} Error={result.ErrorCode} Message={result.Message}");
    if (!result.Success) return 1;

    if (result.GateCheck?.ScenarioReport != null)
    {
        PrintScenarioVerificationReport(result.GateCheck.ScenarioReport);
        return result.GateCheck.ScenarioReport.Success ? 0 : 1;
    }

    return 0;
}

static void EmitLog(LogEventContract evt)
{
    Console.WriteLine($"[{evt.Stage}] {evt.Code} {evt.Message}");
}

static void PrintScenarioVerificationReport(ScenarioVerificationReportContract report)
{
    Console.WriteLine($"VERIFY Scenario={report.Scenario} Success={report.Success}");
    foreach (var fact in report.Facts)
    {
        Console.WriteLine($"FACT {fact.Key}={fact.Value}");
    }

    foreach (var issue in report.Issues)
    {
        Console.WriteLine($"ISSUE {issue}");
    }
}

static int PrintUsageAndReturn()
{
    PrintUsage();
    return 2;
}

static void PrintUsage()
{
    Console.WriteLine("用法：");
    Console.WriteLine("  DocxValidation generate-matrix <矩阵目录>");
    Console.WriteLine("  DocxValidation export-snapshot <输入docx> <输出json>");
    Console.WriteLine("  DocxValidation compare-snapshot <左侧docx> <右侧docx> <输出json>");
    Console.WriteLine("  DocxValidation run-pipeline <输入docx> <输出docx> [封面模式:auto|true|false] [场景名]");
    Console.WriteLine("  DocxValidation verify-doc <docx路径> <场景名>");
    Console.WriteLine("  DocxValidation run-matrix <矩阵目录>");
    Console.WriteLine("  DocxValidation run-real-sample <输入docx> <输出docx>");
}
