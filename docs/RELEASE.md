# oWorkHelper 发布说明

## 当前版本

当前产品版本：`1.2.1`。

需要四段数值版本的字段使用 `1.2.1.0`，例如 `AssemblyVersion`、`AssemblyFileVersion` 和 VSTO `ApplicationVersion`。

## 构建环境

- Windows 10/11
- Visual Studio 2022，安装 Office/SharePoint 开发 (VSTO) 工作负载
- .NET Framework 4.8 Developer Pack
- Outlook 桌面版与 VSTO Runtime 4.0
- NuGet `packages.config` 依赖可还原

## 编译配置

| 配置 | 编译常量 | 在线 OCR | 输出目录 |
| --- | --- | --- | --- |
| Debug | `VSTO40,UseOfficeInterop` | 禁用 | `bin/Debug/` |
| Release-Intranet | `VSTO40,UseOfficeInterop,INTRANET_BUILD` | 禁用 | `bin/Release-Intranet/` |
| Release-Internet | `VSTO40,UseOfficeInterop,INTERNET_BUILD` | 启用（运行时） | `bin/Release-Internet/` |

> 注意：项目**没有** `Release` 配置（`oWorkhelper.vbproj` 只定义上述三种），
> 也不存在 `bin/Release/` 输出目录。

默认安全策略：未定义 `INTERNET_BUILD` 即禁用在线 OCR（运行时常量 `BuildFeatures.OnlineParserEnabled=False`）。

## 构建命令

> **必须提供清单签名证书指纹**，否则 VSTO 构建会以 `error MSB4044: The "SignFile" task was not given a value for the required parameter "CertificateThumbprint"` 失败。
> 通过 `/p:ManifestCertificateThumbprint=<thumbprint>` 传入，或设置同名环境变量
> `IWORKHELPER_MANIFEST_CERT_THUMBPRINT`（`oWorkhelper.vbproj` 会在未显式传参时回退到该变量）。

```powershell
msbuild .\oWorkhelper.sln /t:Restore,Rebuild /p:Configuration=Release-Intranet /p:Platform="Any CPU" /p:ManifestCertificateThumbprint=<thumbprint>
msbuild .\oWorkhelper.sln /t:Restore,Rebuild /p:Configuration=Release-Internet /p:Platform="Any CPU" /p:ManifestCertificateThumbprint=<thumbprint>
msbuild .\tools\OfflineTester\OfflineTester.vbproj /t:Restore,Rebuild /p:Configuration=Debug /p:Platform=AnyCPU
.\tools\OfflineTester\bin\Debug\OfflineTester.exe --selftest
```

`OfflineTester` 已加入 `oWorkhelper.sln`，但仅在 `Debug` 解决方案配置下参与构建
（`Release-Intranet` / `Release-Internet` 只映射 `ActiveCfg`，不含 `.Build.0`，因此发布构建不会带上测试工具）。

VSTO 构建需要清单签名证书。源码仓库不得提交 PFX、私钥、证书指纹或用户级发布配置；构建时通过受保护证书存储、环境变量或 MSBuild 参数提供。

## 发布前检查

- `AssemblyInformationalVersion` 使用 `1.2.1`。
- `AssemblyVersion`、`AssemblyFileVersion`、`ApplicationVersion` 使用合法四段版本 `1.2.1.0`。
- `Release-Intranet` 与 `Release-Internet` 均应构建通过且无警告。
- `OfflineTester --selftest` 必须通过。
- 不提交 `bin/`、`obj/`、`packages/`、日志、本地配置、样例票据、OCR 响应、密钥或证书材料。

## 变更记录

### v1.2.1

- 更新 oWorkHelper 产品版本至 `1.2.1`。

### v1.2.0

- 统一当前发布版本为 `1.2.0`。
- 保留 Outlook PDF 附件归档、滴滴发票合并、常规发票识别、未识别 PDF 保留、命名模板和 OfflineTester。
- 修复启动性能追踪器 ByRef 局部变量初始化警告，使发布构建达到 0 Error / 0 Warning。
- 移除项目文件中的硬编码开发证书绑定，改为构建环境显式提供。
- 精简 docs，保留架构、开发、配置、排障与发布说明。

### 早期基线

- 建立 Outlook VSTO 加载项、PDF 文本识别、Baidu OCR 可选回退、DPAPI 密钥保护和离线测试工具。
