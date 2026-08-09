using System.IO;
using System.Windows.Media;
using System.Windows.Documents;
using System.ComponentModel;
using FirmFormatter.OpenXml.Contracts;
using FirmFormatter.OpenXml.Core;
using Wpf.Ui.Controls;

namespace 事务所出报告批量排版WPF版;

public partial class MainWindow : FluentWindow
{
    private bool _处理中;
    private CancellationTokenSource? _cts;
    private Task? _后台任务;
    private FlowDocument _logDoc = null!;

    private static readonly SolidColorBrush _最小化蓝 = new(System.Windows.Media.Color.FromArgb(0xFF, 0x1A, 0x73, 0xE8));
    private static readonly SolidColorBrush _最大化绿 = new(System.Windows.Media.Color.FromArgb(0xFF, 0x10, 0x7C, 0x10));
    private static readonly SolidColorBrush _关闭红 = new(System.Windows.Media.Color.FromArgb(0xFF, 0xE8, 0x11, 0x23));

    private static readonly SolidColorBrush _成功色;
    private static readonly SolidColorBrush _警告色;
    private static readonly SolidColorBrush _错误色;
    private static readonly SolidColorBrush _信息色;
    private static readonly SolidColorBrush _时间色;
    private static readonly System.Windows.Media.FontFamily _微软雅黑 = new("Microsoft YaHei");

    static MainWindow()
    {
        _最小化蓝.Freeze(); _最大化绿.Freeze(); _关闭红.Freeze();
        _成功色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(87, 166, 74)); _成功色.Freeze();
        _警告色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 165, 32)); _警告色.Freeze();
        _错误色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(231, 72, 86)); _错误色.Freeze();
        _信息色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(204, 204, 204)); _信息色.Freeze();
        _时间色 = new SolidColorBrush(System.Windows.Media.Color.FromRgb(128, 128, 128)); _时间色.Freeze();
    }

    public MainWindow()
    {
        InitializeComponent();
        _logDoc = _logBox.Document;
        
        _startButton.Content = "开始处理";
        _startButton.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;

        _logDoc.Blocks.Clear();
        AppendLog("系统就绪，准备处理文件。", LogType.Info);

        Loaded += (_, _) =>
        {
            foreach (var btn in FindVisualChildren<TitleBarButton>(this))
            {
                var capturedBtn = btn;
                var capturedColor = btn.ButtonType switch
                {
                    TitleBarButtonType.Minimize => _最小化蓝,
                    TitleBarButtonType.Maximize => _最大化绿,
                    TitleBarButtonType.Restore => _最大化绿,
                    TitleBarButtonType.Close => _关闭红,
                    _ => null
                };
                if (capturedColor == null) continue;

                DependencyPropertyDescriptor
                    .FromProperty(BackgroundProperty, typeof(TitleBarButton))
                    .AddValueChanged(capturedBtn, (_, _) =>
                    {
                        var bg = capturedBtn.Background as SolidColorBrush;
                        if (bg != null && bg.Color.A == 0)
                            capturedBtn.Background = capturedColor;
                    });
            }
        };
    }

    private void BrowseFile_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Title = "选择 Word 文件",
            Filter = "Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*"
        };
        if (dlg.ShowDialog() == true)
        {
            _selectedPathBox.Text = dlg.FileName;
            AppendLog($"已选择文件：{dlg.FileName}");
            _statusText.Text = $"已选择文件：{Path.GetFileName(dlg.FileName)}";
        }
    }

    private void BrowseFolder_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFolderDialog
        {
            Title = "选择待处理文件夹"
        };
        if (dlg.ShowDialog() == true)
        {
            _selectedPathBox.Text = dlg.FolderName;
            AppendLog($"已选择文件夹：{dlg.FolderName}");
            _statusText.Text = $"已选择文件夹：{Path.GetFileName(dlg.FolderName)}";
        }
    }

    private void Start_Click(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_处理中)
        {
            // 与后台 finally 中的 Dispose 存在竞态，防御 ObjectDisposedException
            try { _cts?.Cancel(); }
            catch (ObjectDisposedException) { }
            return;
        }

        var path = _selectedPathBox.Text.Trim();
        if (string.IsNullOrEmpty(path))
        {
            System.Windows.MessageBox.Show("请先通过浏览按钮选择文件或文件夹。", "提示",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }

        if (!File.Exists(path) && !Directory.Exists(path))
        {
            System.Windows.MessageBox.Show($"路径不存在：{path}", "错误",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            return;
        }

        _处理中 = true;
        _cts = new CancellationTokenSource();
        _startButton.Content = "取消处理";
        _startButton.Appearance = Wpf.Ui.Controls.ControlAppearance.Danger;
        _progressBar.Value = 0;
        _progressLabel.Text = "";
        _statusText.Text = "正在处理";
        _logDoc.Blocks.Clear();

        var includeSubfolders = _includeSubfoldersCheck.IsChecked == true;
        var token = _cts.Token;
        // 保留 Task 引用：便于关窗时等待后台任务结束，也避免未观察异常无声丢失
        _后台任务 = Task.Run(() => 后台处理(path, includeSubfolders, token));
    }

    private void 后台处理(string path, bool includeSubfolders, CancellationToken token)
    {
        try
        {
            if (File.Exists(path))
            {
                if (token.IsCancellationRequested) return;
                // 与文件夹模式保持一致：已排版输出文件不再重复排版，避免套娃命名
                if (输出文件命名规则.是已排版文件(path))
                {
                    Dispatcher.Invoke(() =>
                    {
                        AppendLog($"所选文件已是排版输出（_已排版），为避免重复排版已跳过：{Path.GetFileName(path)}", LogType.Warning);
                        _statusText.Text = "已跳过：所选文件为已排版输出";
                    });
                    return;
                }
                Dispatcher.Invoke(() => AppendLog($"开始处理文件：{path}"));
                var output = 处理单个文件(path);
                Dispatcher.Invoke(() =>
                {
                    AppendLog($"输出文件：{output}");
                    _progressBar.Value = 100;
                    _progressLabel.Text = "1/1";
                    _statusText.Text = $"处理完成：{Path.GetFileName(output)}";
                });
            }
            else
            {
                var searchOption = includeSubfolders
                    ? SearchOption.AllDirectories
                    : SearchOption.TopDirectoryOnly;
                var files = Directory.GetFiles(path, "*.docx", searchOption)
                    .Where(f => !输出文件命名规则.是已排版文件(f))
                    .ToList();

                if (files.Count == 0)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AppendLog("未找到需要处理的 docx 文件。");
                        _statusText.Text = "未找到文件";
                    });
                    return;
                }

                Dispatcher.Invoke(() => AppendLog($"开始批量处理：{path}（共 {files.Count} 个文件）"));

                var success = 0;
                var fail = 0;
                for (var i = 0; i < files.Count; i++)
                {
                    if (token.IsCancellationRequested) break;
                    var file = files[i];
                    var idx = i + 1;
                    try
                    {
                        var output = 处理单个文件(file);
                        success++;
                        if (token.IsCancellationRequested) break;
                        Dispatcher.Invoke(() =>
                        {
                            AppendLog($"[{idx}/{files.Count}] ✓ {Path.GetFileName(file)} → {Path.GetFileName(output)}", LogType.Success);
                            _progressBar.Value = (double)idx / files.Count * 100;
                            _progressLabel.Text = $"{idx}/{files.Count}";
                            _statusText.Text = $"正在处理：{Path.GetFileName(file)}";
                        });
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        fail++;
                        if (!token.IsCancellationRequested)
                        {
                            Dispatcher.Invoke(() =>
                                AppendLog($"[{idx}/{files.Count}] ⚠ 文件被占用：{Path.GetFileName(file)} - {ex.Message}", LogType.Warning));
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                    catch (Exception ex)
                    {
                        fail++;
                        if (!token.IsCancellationRequested)
                        {
                            Dispatcher.Invoke(() =>
                                AppendLog($"[{idx}/{files.Count}] ✗ {Path.GetFileName(file)} - {提炼异常日志(ex)}", LogType.Error));
                        }
                    }
                }

                if (token.IsCancellationRequested)
                {
                    Dispatcher.Invoke(() =>
                    {
                        AppendLog($"⚠ 用户已取消处理，已完成 {success + fail}/{files.Count} 个文件", LogType.Warning);
                        _statusText.Text = $"已取消：完成 {success + fail}/{files.Count}";
                    });
                }
                else
                {
                    Dispatcher.Invoke(() =>
                        _statusText.Text = $"批量处理完成：成功 {success}，失败 {fail}");
                }
            }
        }
        catch (Exception ex)
        {
            if (!token.IsCancellationRequested)
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog($"【严重错误】{提炼异常日志(ex)}", LogType.Error);
                    _statusText.Text = "处理失败";
                });
            }
        }
        finally
        {
            Dispatcher.Invoke(() =>
            {
                _处理中 = false;
                // 先原子置 null 再 Dispose，避免与 Start_Click/OnClosing 的 Cancel 形成 ObjectDisposedException 竞态
                var cts = Interlocked.Exchange(ref _cts, null);
                cts?.Dispose();
                _startButton.Content = "开始处理";
                _startButton.Appearance = Wpf.Ui.Controls.ControlAppearance.Primary;
                if (string.IsNullOrEmpty(_progressLabel.Text))
                    _progressLabel.Text = "完成";
            });
        }
    }

    private string 处理单个文件(string inputPath)
    {
        var outputPath = 输出文件命名规则.生成输出路径(inputPath);

        var request = new RequestContract
        {
            InputPath = inputPath,
            OutputPath = outputPath,
            HasCoverOverride = null,
            ScenarioName = "常规",
            RunSource = "GUI"
        };

        var pipeline = new DocumentPipeline(evt =>
        {
            // 日志级别直接取事件的 Type 字段，不再靠消息内容猜
            var type = evt.Type switch
            {
                "error" => LogType.Error,
                "warning" => LogType.Warning,
                _ => LogType.Info
            };
            var msg = $"[{evt.Stage}] {evt.Message}";
            // BeginInvoke 异步投递：逐条日志不再同步阻塞后台流水线线程
            Dispatcher.BeginInvoke(() => AppendLog(msg, type));
        });

        var result = pipeline.Run(request);

        if (!result.Success)
            throw new InvalidOperationException(
                $"流水线返回失败（{result.ErrorCode ?? "unknown"}）：{result.Message ?? "未知错误"}");

        return result.OutputPath;
    }

    private static string 提炼异常日志(Exception ex)
    {
        var text = (ex.Message ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
        if (string.IsNullOrEmpty(text))
            text = "文档结构异常或损坏";
        if (text.Contains("openxml_engine_failed"))
            text = "OpenXML引擎执行失败";
        if (text.Contains("bad_request_json"))
            text = "OpenXML引擎请求参数异常";
        if (text.Contains("invalid_request"))
            text = "OpenXML引擎请求参数不完整";
        return text;
    }

    public enum LogType { Info, Success, Warning, Error }

    private void AppendLog(string message, LogType type = LogType.Info)
    {
        var timeStr = DateTime.Now.ToString("HH:mm:ss");
        string prefix = $"> [{timeStr}] ";
        string icon = "";
        SolidColorBrush colorBrush = _信息色;

        // 级别一律由调用方显式传入（流水线日志取 LogEventContract.Type），不再靠消息内容猜，
        // 避免含"失败""错误"字样的正常日志被误染红

        switch (type)
        {
            case LogType.Success:
                icon = "🟢 ";
                colorBrush = _成功色;
                message = message.Replace("✓ ", "").Replace("成功：", "");
                break;
            case LogType.Warning:
                icon = "🟡 ";
                colorBrush = _警告色;
                message = message.Replace("⚠ ", "").Replace("警告：", "");
                break;
            case LogType.Error:
                icon = "🔴 ";
                colorBrush = _错误色;
                message = message.Replace("✗ ", "").Replace("错误：", "").Replace("【严重错误】", "");
                break;
            default:
                colorBrush = _信息色;
                break;
        }

        var paragraph = new Paragraph();
        paragraph.Inlines.Add(new Run(prefix) { Foreground = _时间色, FontFamily = _微软雅黑 });
        paragraph.Inlines.Add(new Run(icon + message) { Foreground = colorBrush, FontFamily = _微软雅黑 });

        _logDoc.Blocks.Add(paragraph);
        // 限制日志段落上限，裁剪最旧的，防止大批量时内存和滚动性能退化
        while (_logDoc.Blocks.Count > 5000)
        {
            _logDoc.Blocks.Remove(_logDoc.Blocks.FirstBlock);
        }
        _logBox.ScrollToEnd();
    }

    protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
    {
        if (_处理中)
        {
            // 后台任务仍在跑：先取消本次关闭，通知任务取消，等任务结束回调里再真正关窗，
            // 保证进程退出前写盘完成，也避免窗口销毁后 finally 里的 Dispatcher 调用抛异常逃逸
            e.Cancel = true;
            try { _cts?.Cancel(); }
            catch (ObjectDisposedException) { }
            AppendLog("正在等待后台任务结束后关闭窗口…", LogType.Warning);
            _后台任务?.ContinueWith(
                _ => Dispatcher.BeginInvoke(new Action(Close)),
                TaskScheduler.Default);
        }
        base.OnClosing(e);
    }

    private static IEnumerable<T> FindVisualChildren<T>(System.Windows.DependencyObject parent) where T : System.Windows.DependencyObject
    {
        if (parent == null) yield break;
        var children = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < children; i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is T typed) yield return typed;
            foreach (var desc in FindVisualChildren<T>(child)) yield return desc;
        }
    }
}
