# 开发指南

> 本文面向 oWorkHelper 的开发者与维护者，涵盖开发环境、调试方法、测试工具和关键设计决策。

## 1. 开发环境搭建

- **IDE**：Visual Studio 2022，安装"Office/SharePoint 开发（VSTO）"工作负载。
- **框架**：.NET Framework 4.8 开发包。
- **Outlook**：需安装 Outlook 桌面版（调试时 F5 启动 Outlook）。
- **NuGet**：依赖通过 `packages.config` 管理，构建前自动或手动还原 `packages/` 目录。

## 2. 项目约定

- `Option Strict Off`、`Option Explicit On`、`Option Infer On`（见 `oWorkhelper.vbproj`）。
  切换 `Option Strict On` 的成本已实测：`msbuild oWorkhelper.sln /t:Rebuild /p:Configuration=Release-Intranet /p:OptionStrict=On`
  只产生 **16 个错误**（`BC30574` 后期绑定 ×15、`BC32023` Object 不是集合类型 ×1），集中在
  `Core/Common/ExplorerFolderService.vb` 与 `MainRibbon.vb` 两个文件；本轮不切换（需专项排期），详见 `REVIEW_TRACKING.md` O-46。
- 控件/成员命名采用 PascalCase 与 camelCase 混用（WinForms 惯例）。
- UI 文本与注释使用中文。
- `Core/` 业务代码不依赖 Outlook 或 `My.Settings`，因此 `OfflineTester` 可直接链接同一批源文件做离线测试。

## 3. 调试

- **Outlook 调试**：F5 启动 Outlook（`DebugInfoExeName` 指向 `outlook.exe`），在 Outlook 中操作触发断点。
- **离线调试**：使用 `tools/OfflineTester` 调试 Core 层逻辑，无需启动 Outlook。
- **启动性能**：`StartupPerformanceTracker` 在日志中记录各阶段耗时，用于定位启动缓慢。

## 4. OfflineTester 离线测试工具

独立控制台应用，不依赖 Outlook，链接 Core 源文件。

### 构建

```
MSBuild tools\OfflineTester\OfflineTester.vbproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU
```

### 常用命令

| 命令 | 用途 |
|------|------|
| `--selftest` | 内置自测（DPAPI 往返、命名模板、OCR JSON 解析、识别合并等） |
| `<pdf> --local-only` | 仅本地解析 |
| `<pdf> --general-invoice` | 常规发票识别诊断 |
| `<pdf> --general-invoice --dump-candidates` | 显示字段候选与评分 |
| `<pdf> --classify` | 显示 PDF 分流分类 |
| `--preflight <目录>` | 归档前预检查 |
| `--simulate-error <码名\|list>` | 查看错误码文案 |
| `<pdf> --ocr` | 本地优先 + OCR 兜底（需 INTERNET_BUILD 构建，见下） |
| `<pdf> --force-ocr` | 强制 OCR（需 INTERNET_BUILD 构建，见下） |
| `--save-baidu-config` | 用环境变量写入本机加密配置 |
| `--parse-ocr-json <文件>` | 解析脱敏 OCR JSON |
| `--compare-expected <文件>` | 回归对比（json/csv 期望文件） |

> **`--ocr` / `--force-ocr` 在默认构建下不可用。** `OfflineTester.vbproj` 未定义 `INTERNET_BUILD`，
> 因此 `BuildFeatures.OnlineParserEnabled=False`：`BaiduOcrInvoiceRecognizer.IsAvailable()` 返回 False，
> `Recognize()` 直接返回 `ConfigurationMissing`（“当前构建版本未启用在线解析功能”），**不会发出任何网络请求**。
> 需要真实联调时用下面的命令重建测试工具：
>
> ```
> MSBuild tools\OfflineTester\OfflineTester.vbproj /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU /p:DefineConstants=INTERNET_BUILD
> ```
>
> 注意该构建会启用在线解析（仅用于联调，勿用于发布门禁的 `--selftest`）。

### 退出码

| 退出码 | 含义 |
|--------|------|
| 0 | 正常 |
| 1 | 运行异常 / selftest 有失败项 |
| 2 | 参数错误 / 文件不存在 |

### OCR 联调

```
set BAIDU_OCR_AK=你的APIKey
set BAIDU_OCR_SK=你的SecretKey
OfflineTester.exe <pdf> --force-ocr --save-response ocr-raw
```

前提：用 `/p:DefineConstants=INTERNET_BUILD` 构建过测试工具（见上表说明），否则该命令只会得到
`ConfigurationMissing` 而不会联网。

工具仅打印脱敏 AK（如 `ab****yz`），不打印完整密钥。

## 5. Outlook 端到端测试

在 Outlook 环境中人工验证的关键项：

1. 加载项随 Outlook 启动正常加载，Ribbon 显示"工作助手"。
2. 点击"归档"正常执行完整流程。
3. 点击"设置"打开设置窗口，保存后直接关闭。
4. 内网版本隐藏在线解析 UI，外网版本显示。
5. 版本号显示正确（当前正式版本为 `1.3.0`）。
6. 进度窗口正常显示并更新。
7. 单封邮件失败不影响其他邮件处理。

## 6. 关键设计决策

### 归档运行锁

`ArchiveRunGuard` 使用 `Interlocked.CompareExchange` 原子获取运行锁，`ArchiveRunToken` 通过 `Using` 语义保证释放。预检查不检查运行状态（避免自我阻断）。

### 识别管道

`RecognitionPipeline` 先本地后在线：
- 本地识别（`LocalTextInvoiceRecognizer`）：常规增值税发票委派 `GeneralInvoiceLocalRecognizer`，滴滴/行程单由自身解析
- `PreferLocalParse=True` 且“命名核心字段”充分 → 直接采用本地结果
- `PreferLocalParse=False`、`AutoFallbackToOcr=True` 或文本疑似图片型 → 调用 `BaiduOcrInvoiceRecognizer`
  （注意：`PreferLocalParse=False` 时**即使本地字段充分也会调用 OCR**）
- OCR 失败但本地有字段 → 返回本地部分成功；两者都不可用 → 返回失败 / 需要 OCR
- 多 PDF 识别结果通过 `InvoiceRecognitionMerger` 合并（滴滴：先逐 PDF 识别，再合并识别结果与 PDF）

### 常规发票候选评分

`GeneralInvoiceLocalRecognizer` 对每个字段产生多个候选，由 `GeneralInvoiceCandidateScorer` 按位置、格式、长度等评分择优。商品明细通过 `PdfTableRegionDetector` + `GeneralInvoiceLineItemParser` 解析。

### 默认安全策略

未定义 `INTERNET_BUILD` 则在线解析禁用（运行时常量收口，网络栈仍被编译进程序集，**不是编译期剔除**）。
`BaiduOcrInvoiceRecognizer` 在 `OnlineParserEnabled=False` 时直接拒绝请求。内网版即使配置文件残留在线参数也不调用 OCR。

### Secret Key 保护

`SecretProtector` 使用 Windows DPAPI（CurrentUser 作用域）加密 Secret Key（API Key 不加密），密文存储带 `DPAPI:` 前缀，
主存储在 `My.Settings`（`user.config`）。明文迁移在两个时机触发：**打开设置窗口时**，以及**每次点击“归档”时**
（`MainRibbon` 在预检查前调用 `MigratePlaintextIfNeeded`），避免用户从不打开设置而长期残留明文。

## 7. 样例与脱敏

- 样例 PDF 放 `sample/`（`.gitignore` 已忽略），不提交真实票据。
- OCR 脱敏返回放 `ocr-raw/`（`.gitignore` 已忽略）。
- 回归期望文件与真实样例均不提交。
- 校准 `BaiduInvoiceFieldMapper` 时对照脱敏 JSON 更新映射表。

## 8. 诊断脚本

`tools/OutlookResiliency/` 包含两个 PowerShell 诊断脚本：
- `check_iworkhelper_addin_registration.ps1`：检查加载项注册状态。
- `get_iworkhelper_details.ps1`：获取加载项详细信息。
