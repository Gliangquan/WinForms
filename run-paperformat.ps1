$ErrorActionPreference = "Stop"

$env:DOTNET_CLI_HOME = "C:\Users\liangquan\.codex\memories\dotnet"
$env:DOTNET_SKIP_FIRST_TIME_EXPERIENCE = "1"
$env:DOTNET_NOLOGO = "1"
$env:MSBuildEnableWorkloadResolver = "false"

dotnet run --project "src\PaperFormat.WinForms\PaperFormat.WinForms.csproj" -m:1
