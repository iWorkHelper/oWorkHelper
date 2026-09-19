# oWorkHelper

`oWorkHelper` 是 [iWorkHelper Organization](https://github.com/iWorkHelper) 旗下的 Outlook VSTO 加载项，用于批量处理邮件中的 PDF 附件（发票、行程单），完成自动识别、合并、命名和归档。

**当前版本**：`v1.3.0`

下载：[最新 Release](https://github.com/iWorkHelper/oWorkHelper/releases/latest)

## 主要功能

- 在 Outlook 功能区提供"归档"和"设置"入口。
- 批量读取选中邮件的 PDF 附件，按邮件类型自动分流处理。
- **滴滴发票**：同一封邮件的发票与行程单 PDF 合并为一个文件，识别后统一命名归档。
- **常规发票**：每张增值税发票 PDF 单独识别、命名、归档。
- **未识别 PDF**：无法识别的 PDF 按 `未识别_{原始文件名}` 保留归档，不丢弃文件。
- 本地解析基于 PdfPig 文本抽取与规则识别（坐标行重建、候选评分）。
- 外网版支持百度 OCR 在线解析，本地字段不足时自动回退。
- 自定义命名模板，支持变量说明与预览。
- 归档前预检查、运行锁防并发、进度展示、结果报告。
- Secret Key 通过 Windows DPAPI 加密保护。
- 提供 OfflineTester 离线测试与诊断工具。

## 运行环境

- Windows 10/11
- Outlook 桌面版（Microsoft 365 或 Office 2016+）
- .NET Framework 4.8
- VSTO Runtime 4.0

## 技术栈

| 技术 | 说明 |
|------|------|
| VB.NET | 开发语言 |
| .NET Framework 4.8 | 目标框架 |
| VSTO 4.0 | Outlook 加载项框架 |
| WinForms | 设置窗体与进度窗口 |
| PdfPig 0.1.14 | PDF 文本抽取与合并 |
| 百度 OCR | 在线发票识别（外网版可选） |
| Windows DPAPI | Secret Key 加密存储 |
| MSBuild / VS2022 | 构建工具 |

## 项目目录结构

```
oWorkHelper/
├── Core/                     # 业务核心
│   ├── Archive/              # 命名规则、模板引擎、归档规划与执行
│   ├── Common/               # 结果对象、路径工具、错误处理、编译期开关
│   ├── Configuration/        # OCR 配置读取
│   ├── Diagnostics/          # 启动性能跟踪
│   ├── Invoice/              # 发票/行程数据模型
│   ├── Logging/              # 线程安全文件日志
│   ├── Mail/                 # 邮件附件读取与分组
│   ├── Ocr/Baidu/            # 百度 OCR 接入
│   ├── Pdf/                  # PDF 文本抽取、布局分析、合并
│   ├── Recognition/          # 识别管道、本地/在线识别器
│   ├── Security/             # DPAPI 密钥加密
│   └── Workflow/             # 批量归档编排、分流、预检查、运行锁
├── docs/                     # 项目文档
├── My Project/               # 程序集信息与设置
├── Resources/                # Ribbon 图标
├── tools/
│   ├── OfflineTester/        # 离线测试控制台工具
│   └── OutlookResiliency/    # 加载项诊断脚本
├── oWorkhelper.sln           # 解决方案
├── oWorkhelper.vbproj        # 主工程文件
├── MainRibbon.vb             # 功能区（归档/设置按钮）
├── SettingsForm.vb           # 设置窗体
├── ProgressForm.vb           # 进度窗口
├── ThisAddIn.vb              # 加载项入口
└── packages.config           # NuGet 依赖清单
```

## 内网版与外网版

| 特性 | Release-Intranet（内网版） | Release-Internet（外网版） |
|------|---------------------------|---------------------------|
| 编译常量 | `INTRANET_BUILD` | `INTERNET_BUILD` |
| 在线 OCR | 禁用 | 可选启用 |
| 设置界面 | 隐藏 OCR 配置和在线解析选项 | 完整显示 |
| 版本标识 | 显示"（内网版）" | 无额外标识 |
| 外网请求 | 不产生任何外网请求 | 启用 OCR 时调用百度 API |

**默认安全策略**：未定义 `INTERNET_BUILD` 即禁用在线解析。内网版即使配置文件残留在线参数也不调用 OCR。

## 开发与编译

开发环境需要 Visual Studio 2022、“Office/SharePoint 开发（VSTO）”工作负载和 .NET Framework 4.8 开发包。

### 编译配置

| 配置 | 编译常量 | 输出目录 | 在线 OCR |
|------|---------|---------|---------|
| Debug | `VSTO40,UseOfficeInterop` | `bin\Debug\` | 禁用 |
| Release-Intranet | `VSTO40,UseOfficeInterop,INTRANET_BUILD` | `bin\Release-Intranet\` | 禁用 |
| Release-Internet | `VSTO40,UseOfficeInterop,INTERNET_BUILD` | `bin\Release-Internet\` | 运行时启用 |

> 项目**没有** `Release` 配置，也没有 `bin\Release\` 输出目录。

### 编译步骤

> `oWorkhelper.sln` 的构建**必须提供清单签名证书指纹**，否则 VSTO 构建报
> `error MSB4044`（`SignFile` 缺少 `CertificateThumbprint`）。
> 用 `/p:ManifestCertificateThumbprint=<thumbprint>` 传入，或设置环境变量
> `IWORKHELPER_MANIFEST_CERT_THUMBPRINT`（工程文件会在未传参时回退到该变量）。

```bash
# 内网版
MSBuild oWorkhelper.sln /p:Configuration="Release-Intranet" /p:Platform="Any CPU" /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>

# 外网版
MSBuild oWorkhelper.sln /p:Configuration="Release-Internet" /p:Platform="Any CPU" /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>

# OfflineTester（调试工具，已加入解决方案，仅 Debug 配置参与构建）
MSBuild oWorkhelper.sln /p:Configuration=Debug /p:Platform="Any CPU" /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>
# 或单独构建：
MSBuild tools\OfflineTester\OfflineTester.vbproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU
```

## 安装与使用

VSTO 加载项通过 ClickOnce 发布或手动注册安装到 Outlook。加载行为设置为 `LoadBehavior=3`（随 Outlook 自动加载）。

签名使用本地临时证书，证书文件由 `.gitignore` 排除；正式部署应替换为受信任的代码签名证书。

### 基本使用方法

1. 在 Outlook 中选中一封或多封邮件。
2. 点击功能区 **工作助手 → 发票归档 → 归档**。
3. 系统执行预检查后显示进度窗口。
4. 完成后弹出汇总，详细结果见归档报告和日志。

首次使用前需在 **设置** 中配置归档目录。

### 基本配置

通过 Outlook 功能区 **工作助手 → 设置** 打开设置窗口：

- **归档目录**：PDF 归档后存放位置。
- **命名模板**：滴滴发票、常规发票、未识别 PDF 三类模板可分别自定义。
- **在线 OCR**（仅外网版）：配置百度 API Key / Secret Key。

设置保存后直接关闭窗口。

## 安全与隐私

- Secret Key 使用 Windows DPAPI（CurrentUser 作用域）加密存储，不明文保存。
- 日志和报告中不输出 AK、SK 或 Access Token。
- 内网版通过运行时常量 `BuildFeatures.OnlineParserEnabled=False` 收口在线解析（fail-closed）：网络栈仍被编译进程序集，但不会发起任何外网请求。
- `.gitignore` 已覆盖 `user.config`、`baidu-ocr.config.xml`、样例 PDF 等敏感文件。

## 命名规则与默认行为

| 类型 | 默认模板 | 示例 |
|------|---------|------|
| 滴滴发票 | `{乘车日期}_{金额}_{出发地点}_{到达地点}` | `20260518_138.46_上海虹桥站_上海市徐汇区.pdf` |
| 常规发票 | `{开票日期}_{金额}_{销售方名称}` | `20260526_138.46_某某科技有限公司.pdf` |
| 未识别 PDF | `未识别_{原始文件名}` | `未识别_某某扫描件.pdf` |

- `{乘车日期}` 优先取行程出发日期，其次取开票日期。
- `{金额}` 对滴滴取行程金额 > 价税合计；对常规发票取价税合计 > 金额。
- 空字段自动跳过，非法字符自动清理，同名文件自动追加序号。
- 字段严重不足时回退为 `未识别票据_{邮件主题}_{时间戳}`。

## 常见问题

| 问题 | 解决方案 |
|------|---------|
| 加载项被 Outlook 禁用 | 文件 → 选项 → 加载项 → 已禁用项目 → 启用 |
| "尚未配置归档目录" | 在设置中选择归档文件夹 |
| "已有归档任务正在运行" | 等待当前任务完成 |
| PDF 命名为"未识别_..." | 该 PDF 未能识别为发票，已按未识别归档 |
| 换机器后 OCR 失效 | DPAPI 仅当前用户可解密，需重新输入 Secret Key |
| 常规发票被误当未识别 | 该 PDF 缺发票特征，可用 OfflineTester --classify 复核 |

详细排查见 [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)。

## 版本号规则

对外产品版本使用 Major.Minor.Patch，当前版本为 1.3.0。需要四段数值版本的 VSTO / Assembly 字段使用 1.3.0.0。详细规则见 [docs/RELEASE.md](docs/RELEASE.md)。

## 文档入口

| 文档 | 内容 |
|------|------|
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | 系统架构、代码结构、核心流程、安全设计 |
| [DEVELOPMENT.md](docs/DEVELOPMENT.md) | 开发环境、调试、OfflineTester 和设计决策 |
| [CONFIGURATION.md](docs/CONFIGURATION.md) | 用户设置、OCR 配置、命名模板变量 |
| [TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 故障排查、诊断工具、日志位置 |
| [RELEASE.md](docs/RELEASE.md) | 版本、构建、发布检查和变更记录 |

## Release

- [最新版本下载](https://github.com/iWorkHelper/oWorkHelper/releases/latest)
- [全部 Releases](https://github.com/iWorkHelper/oWorkHelper/releases)
- [问题反馈](https://github.com/iWorkHelper/oWorkHelper/issues)

## 许可证

本项目采用 [MIT License](LICENSE) 开源许可证。
