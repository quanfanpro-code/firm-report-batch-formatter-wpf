# DocxValidation

`DocxValidation` 是事务所排版程序附带的命令行验证工具。它复用主程序的分类、排版、校验、快照和差异服务，不维护第二套规则。

## 运行

在仓库根目录执行：

```powershell
dotnet run --project .\Tools\DocxValidation\DocxValidation.csproj -- <命令> <参数>
```

可用命令：

```text
generate-matrix <矩阵目录>
export-snapshot <输入docx> <输出json>
compare-snapshot <左侧docx> <右侧docx> <输出json>
run-pipeline <输入docx> <输出docx> [auto|true|false] [场景名]
verify-doc <docx路径> <场景名>
run-matrix <矩阵目录>
run-real-sample <输入docx> <输出docx>
```

`generate-matrix` 生成合成 Word 样例。`run-matrix` 会重新生成样例并逐份执行排版；已有专用断言的场景还会执行场景验证。生成的 DOCX、快照和差异 JSON 都是本地测试材料，不应提交到源码仓库。

每次矩阵复跑使用新的空目录，避免覆盖旧输入样例或因输出重名而失败。

`export-snapshot` 和 `compare-snapshot` 只允许新建 `.json` 文件；输出已经存在或扩展名错误时返回失败，保留原文件。

`verify-doc` 会执行结构、业务及指定场景检查，并输出 `GateSuccess`。只有一个输入文件时没有排版前基线，`BeforeAfterComparison=False` 表示没有验证前后内容保留，不能用它替代完整流水线。

`run-real-sample` 执行通用验收，不要求文件含有“公司的基本情况”等历史样本章节；合成矩阵的固定内容断言仍保留。
