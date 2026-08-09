using System.IO;
using System.Windows;

namespace 事务所出报告批量排版WPF版;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : System.Windows.Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // 全局兜底：UI 线程未处理异常 + Task 未观察异常，都落盘一份错误日志并提示用户
        DispatcherUnhandledException += App_DispatcherUnhandledException;
        TaskScheduler.UnobservedTaskException += App_UnobservedTaskException;
        base.OnStartup(e);
    }

    private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
    {
        落盘错误日志("UI线程未处理异常", e.Exception);
        e.Handled = true;
        System.Windows.MessageBox.Show(
            $"发生未处理的错误，详细信息已写入错误日志：\n{错误日志路径}\n\n{e.Exception.Message}",
            "程序错误",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Error);
    }

    private void App_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        落盘错误日志("后台任务未观察异常", e.Exception);
        // 标记已观察，避免进程被直接终止
        e.SetObserved();
    }

    private static string 错误日志路径 =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "事务所出报告批量排版WPF版",
            "error.log");

    private static void 落盘错误日志(string 来源, Exception ex)
    {
        try
        {
            var path = 错误日志路径;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.AppendAllText(path,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {来源}{Environment.NewLine}{ex}{Environment.NewLine}{new string('-', 60)}{Environment.NewLine}");
        }
        catch
        {
            // 日志落盘本身失败时不再抛出，避免掩盖原始异常
        }
    }
}
