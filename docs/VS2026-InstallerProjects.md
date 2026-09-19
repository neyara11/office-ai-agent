# Visual Studio 2026 打开 OfficeAgent.vdproj 失败处理

> 构建链路总览（code vs installer）见 [build-and-installer.md](./build-and-installer.md)。
> 打 MSI 前请先跑 `.\build-installer-prep.bat`（Release 代码 + SourcePath 审计）。

## 现象

在 Visual Studio 2026 打开 `AiHelper.sln` 时，`OfficeAgent/OfficeAgent.vdproj` 可能提示：

```text
找不到此项目类型所基于的应用程序
54435603-dbb4-11d2-8724-00a0c9a8b90c
```

这个 GUID 是 Visual Studio Installer Projects 的 `.vdproj` 项目类型。它不是 `WordAi`、`ExcelAi`、`PowerPointAi`、`ShareRibbon` 的 VB.NET/VSTO 版本过低，也不是 `.NET Framework 4.7.2` 本身不兼容。

## 根因

`.vdproj` 不是 Visual Studio 标准内置项目类型，需要当前 Visual Studio 实例安装并启用 `Microsoft Visual Studio Installer Projects` 扩展。

如果 VS 2026 没有安装该扩展、扩展版本过旧、扩展被禁用，或者安装到了另一个 VS 实例，`OfficeAgent.vdproj` 就会显示不兼容。

当前解决方案中的普通代码项目仍然可以正常构建：

- `ShareRibbon/ShareRibbon.vbproj`
- `WordAi/WordAi.vbproj`
- `ExcelAi/ExcelAi.vbproj`
- `PowerPointAi/PowerPointAi.vbproj`

## 处理步骤

1. 在 Visual Studio 2026 打开 `Extensions -> Manage Extensions`。
2. 搜索并安装 `Microsoft Visual Studio Installer Projects`。
3. 确认扩展安装到 Visual Studio 2026 这个实例，而不是旧版 Visual Studio。
4. 关闭所有 Visual Studio 实例。
5. 重新启动 Visual Studio 2026。
6. 重新打开 `AiHelper.sln`。

如果仍然失败：

1. 在 `Extensions -> Manage Extensions -> Installed` 确认扩展已启用。
2. 更新扩展到 3.x 或更高版本。
3. 使用 `devenv.com .\AiHelper.sln /Rebuild Debug` 验证代码项目是否仍可构建。
4. 如果扩展短期不可用，先用 VS 2026 开发和调试代码项目，再用支持该扩展的 Visual Studio 实例构建 MSI。

## 不建议的修复方式

- 不要把 `OfficeAgent.vdproj` 的项目 GUID 改成普通 VB.NET 项目 GUID。
- 不要手工大范围重排 `.vdproj` 内容。
- 不要通过升级 `TargetFrameworkVersion` 来解决这个错误；它和 `.vdproj` 项目类型加载无关。
- 不要为了绕过加载错误删除 `OfficeAgent` 项目，除非明确决定迁移安装包方案。

## 长期建议

`.vdproj` 适合快速生成 MSI，但它依赖 Visual Studio 扩展，自动化能力和版本兼容性都比较弱。后续如果要继续做三合一安装包、减少重复 DLL、支持 CI 构建，建议规划迁移到 WiX Toolset 或 MSIX。

短期内保持 `OfficeAgent.vdproj` 最小改动；中长期把安装包瘦身、共享依赖去重、注册表写入和升级卸载逻辑迁移到可脚本化、可 CI 构建的安装链路。

## 命令行构建 MSI（验证过的步骤）

1. 安装扩展 `Microsoft Visual Studio Installer Projects 2022`（3.x，安装目标为 `Microsoft.VisualStudio.Community [17.0,19.0)`；旧版 1.x 只支持 VS 2017/2019，会被 VSIXInstaller 拒绝）。
2. 生成 Release 代码并审计输入：`.\build-installer-prep.bat`。
3. 命令行构建 `devenv.com` 会报 `ERROR: An error occurred while validating. HRESULT = '8000000A'`。这是 VS 2012+ 不支持进程外构建 setup 项目的已知问题，不是代码错误。设置注册表开关（HKCU，无需管理员）：

   ```powershell
   $instanceId = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -version "[17.0,18.0)" -property instanceId
   $regPath = "HKCU:\SOFTWARE\Microsoft\VisualStudio\17.0_${instanceId}_Config\MSBuild"
   New-Item -Path $regPath -Force | Out-Null
   Set-ItemProperty -Path $regPath -Name "EnableOutOfProcBuild" -Value 0 -Type DWord
   ```

   也可以直接运行 `Common7\IDE\CommonExtensions\Microsoft\VSI\DisableOutOfProcBuild\DisableOutOfProcBuild.exe`（需从 VS 安装目录上下文运行，否则找不到实例）。
4. 构建：`devenv.com AiHelper.sln /Build Release /Project OfficeAgent /Out <log>` → `OfficeAgent\Release\OfficeAgent.msi`。

### 松散 DLL 输入（易漏）

`OfficeAgent.vdproj` 中有以纯文件名（无目录）书写的 `SourcePath`，它们相对 `OfficeAgent\` 解析。仓库跟踪了其中一部分（`Markdig.dll`、`System.Buffers.dll`、`System.Memory.dll`、`System.Runtime.CompilerServices.Unsafe.dll`、`System.Threading.Tasks.Extensions.dll`），但下列四个曾缺失，导致预校验失败：

- `System.Text.Json.dll`
- `System.Diagnostics.DiagnosticSource.dll`
- `System.Numerics.Vectors.dll`
- `System.Threading.Tasks.Dataflow.dll`（不在 `packages/` 中；StreamJsonRpc 2.22.11 的 netstandard2.0 资产依赖它，需从 NuGet `System.Threading.Tasks.Dataflow` 8.0.1 取 `lib/net462`）

前三个可从 `WordAi\bin\Release\` 复制；第四个需单独获取。建议把它们与其余松散 DLL 一起纳入版本管理，保证可复现。

注意：`build\AuditInstallerInputs.ps1` 默认跳过纯文件名（无目录）的 `SourcePath`（`-IncludePlainFiles` 才检查），所以它报告 `PASS` 时仍可能缺少上述 DLL。修复 MSI 输入时不要只依赖该审计结果。

### 升级安装不会替换同版本文件

Windows Installer 只在新文件的版本号更高时才替换已安装的同名程序集。本项目程序集长期保持 `1.0.0.0`，因此用同一个 `ProductVersion` 重新打包后安装，磁盘上的 `ShareRibbon.dll` / 各插件 DLL 仍是旧文件（时间戳不变），表现为「代码改了但功能没变」。

处理方式（任选其一）：

1. 每次发版提升 `AssemblyInfo` 的文件/程序集版本（推荐，符合 MSI 语义）；或
2. 安装时强制覆盖：`msiexec /i OfficeAgent.msi REINSTALL=ALL REINSTALLMODE=amus`（本地验证可用）；或
3. 每次发布提升 `ProductVersion` 并生成新的 `ProductCode`（`UpgradeCode` 保持不变，`RemovePreviousVersions=TRUE`），保证是升级而不是就地修复。

本地验证旧文件是否被替换，可比较哈希：

```powershell
(Get-FileHash WordAi\bin\Release\ShareRibbon.dll).Hash -eq (Get-FileHash "C:\Program Files (x86)\it235\OfficeAiAgent\WordAi\ShareRibbon.dll").Hash
```
