using System.IO;
using System.Runtime.InteropServices;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;

namespace 事务所出报告批量排版WPF版;

/// <summary>
/// 统一入口：带参数按命令行排版，不带参数进图形界面。
/// WinExe 不会自动生成控制台输出，直接在终端里跑时需附加父控制台才能看到日志。
/// </summary>
public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length > 0)
            return RunCli(args);

        var app = new App();
        app.InitializeComponent();
        app.Run();
        return 0;
    }

    private static int RunCli(string[] args)
    {
        if (!Console.IsOutputRedirected)
            附加父控制台();

        if (args.Length == 1 && (args[0] == "-h" || args[0] == "--help"))
        {
            Console.WriteLine("用法：事务所出报告批量排版WPF版 <输入.docx> [输出.docx]");
            Console.WriteLine("输出路径省略时，按界面版同一规则在原目录生成“_已排版”后缀文件。");
            return 2;
        }

        var input = Path.GetFullPath(args[0]);
        var output = args.Length > 1 ? Path.GetFullPath(args[1]) : 输出文件命名规则.生成输出路径(input);

        var request = new RequestContract
        {
            InputPath = input,
            OutputPath = output,
            HasCoverOverride = null,
            IsPureCoverOverride = null,
            ScenarioName = "常规",
            RunSource = "CLI",
        };
        var 参数错误 = request.校验();
        if (参数错误 is not null)
        {
            Console.WriteLine(参数错误);
            return 2;
        }

        var pipeline = new DocumentPipeline(evt =>
            Console.WriteLine($"[{evt.Type}] [{evt.Stage}] {evt.Message}"));
        var result = pipeline.Run(request);

        if (result.Success)
        {
            Console.WriteLine($"Success=True Output={result.OutputPath}");
            if (result.AuditPath is not null)
                Console.WriteLine($"Audit={result.AuditPath}");
            return 0;
        }

        Console.WriteLine($"Success=False Code={result.ErrorCode} Message={result.Message}");
        return 1;
    }

    [DllImport("kernel32.dll")]
    private static extern bool AttachConsole(int dwProcessId);

    private static void 附加父控制台()
    {
        // ATTACH_PARENT_PROCESS；失败说明父进程没有控制台，静默跳过即可
        if (AttachConsole(-1))
        {
            try
            {
                Console.SetOut(new StreamWriter(
                    File.Open("CONOUT$", FileMode.Open, FileAccess.Write, FileShare.Write),
                    Console.OutputEncoding) { AutoFlush = true });
            }
            catch
            {
                // 附加成功但取不到控制台屏幕缓冲区时，保持原有输出通道
            }
        }
    }
}
