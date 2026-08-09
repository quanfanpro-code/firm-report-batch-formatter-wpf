# 事务所出报告批量排版程序

这是一个 Windows 桌面程序，用 Open XML 批量整理事务所报告和财务报表附注的 `.docx` 排版。当前规则按四川华信报告样式编写，重点处理封面、正文、标题、表格、页眉页脚、页码和落款。

程序只修改复制后的 Word 文件，不依赖本机安装 Microsoft Word。原文件不会被覆盖，正式输出只有在处理、门禁和重新打开检查全部完成后才生成。

## 主要功能

- 处理单个文件或整个文件夹，可选择包含子文件夹。
- 自动识别封面和纯封面文档，封面不套用正文页眉页脚规则。
- 统一正文、标题、段落间距、字体和字号。
- 清理表格残留样式，统一数字、边框、对齐和表头格式；普通数字使用千分位并保留两位小数，表头为“序号”的整列保持原样。
- 处理多节文档、横向页面、页眉页脚和页码字段。
- 识别四川华信落款区域，处理事务所名称、注册会计师签字行、城市和日期。
- 在保存前执行 Open XML 结构校验、业务校验和可选场景校验。
- 处理失败时保留带 `_排版失败_时间戳.docx` 的失败件，便于复核，不占用正式输出名。
- 提供合成样例生成、快照导出和结构差异比较命令。

## 适用边界

- 只支持 Windows 和 `.docx`。
- 当前只有四川华信一套规则，不是可切换的多事务所模板平台。
- 宏、嵌入对象、复杂域代码和高度定制的第三方模板应先用副本测试。
- Open XML SDK 报出的 schema 兼容性差异只记警告，不阻断输出；正文缺失、页眉页脚引用损坏、业务门禁失败或最终文件无法重新打开仍会阻断。
- 自动校验不能替代最终人工复核，尤其是分页、跨页表格和签字盖章位置。

## 图形界面

需要安装 [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)。

```powershell
python .\run.py
```

或直接运行：

```powershell
dotnet run --project .\事务所出报告批量排版WPF版.csproj
```

打开后选择 Word 文件或文件夹，设置是否包含子文件夹，再点击“开始处理”。输出文件保存在原目录，文件名以 `_已排版.docx` 结尾。再次点击按钮可以请求取消尚未完成的批次。工具生成的正式输出和失败件不会被再次批量处理。

## 构建和测试

```powershell
dotnet restore .\事务所出报告批量排版WPF版.csproj
dotnet build .\事务所出报告批量排版WPF版.csproj -c Release
dotnet test .\开发测试\文档快照导出测试\文档快照导出测试.csproj -c Release
```

测试会在系统临时目录生成合成 Word 样例，不需要仓库内预存 DOCX。

当前自动测试覆盖真实窗口创建、封面与纯封面、标题和自动编号、表格与合并单元格、多节页眉页脚、脚注尾注、落款、复杂结构保留、失败件保留和输出门禁。正式交付前仍应在 Word 或 WPS 中检查分页和签章位置。

发布 Windows x64 单文件程序：

```powershell
dotnet publish .\事务所出报告批量排版WPF版.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o .\publish
```

## 文档验证工具

验证工具位于 `Tools\DocxValidation`。先构建：

```powershell
dotnet build .\Tools\DocxValidation\DocxValidation.csproj -c Release
```

常用命令：

```powershell
# 生成合成测试样例
dotnet run --project .\Tools\DocxValidation\DocxValidation.csproj -- generate-matrix C:\Temp\FirmFormatterMatrix

# 生成样例并执行完整排版矩阵
dotnet run --project .\Tools\DocxValidation\DocxValidation.csproj -- run-matrix C:\Temp\FirmFormatterMatrix

# 导出文档结构快照
dotnet run --project .\Tools\DocxValidation\DocxValidation.csproj -- export-snapshot C:\Temp\input.docx C:\Temp\snapshot.json

# 比较两份文档的结构快照
dotnet run --project .\Tools\DocxValidation\DocxValidation.csproj -- compare-snapshot C:\Temp\left.docx C:\Temp\right.docx C:\Temp\diff.json

# 执行一次排版
dotnet run --project .\Tools\DocxValidation\DocxValidation.csproj -- run-pipeline C:\Temp\input.docx C:\Temp\output.docx auto 常规
```

封面模式接受 `auto`、`true` 或 `false`。工具还提供 `verify-doc` 和 `run-real-sample` 命令，完整参数可直接运行工具查看。

## 主要目录

```text
Contracts/                         请求、响应和校验结果
Core/                              分类、排版、校验、快照和差异服务
Tools/DocxValidation/              合成样例与命令行验证工具
开发测试/文档快照导出测试/         xUnit 回归测试
MainWindow.xaml                    桌面界面
run.py                             本地启动脚本
```

## 数据和隐私

程序没有文档上传功能。公开仓库不包含客户文档、生成后的 DOCX、冻结结果、内部计划或本机路径。提交问题时请使用脱敏样例；最稳妥的做法是用验证工具生成合成样例后复现。

## 许可证

本项目采用 [GNU Affero General Public License v3.0](LICENSE)，许可证标识为 `AGPL-3.0-only`。

Open XML SDK 和 WPF-UI 使用 MIT 许可证；测试依赖使用 MIT 或 Apache-2.0 许可证。
