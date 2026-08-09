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
