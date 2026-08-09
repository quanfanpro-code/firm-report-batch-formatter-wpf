from pathlib import Path
import subprocess


项目文件 = Path(__file__).with_name("事务所出报告批量排版WPF版.csproj")
subprocess.Popen(["dotnet", "run", "--project", str(项目文件)], cwd=项目文件.parent)
