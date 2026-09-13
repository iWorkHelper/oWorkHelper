# 开发指南

> 本文面向 oWorkHelper 的开发者与维护者，涵盖开发环境、调试方法、测试工具和关键设计决策。

## 1. 开发环境搭建

- **IDE**：Visual Studio 2022，安装"Office/SharePoint 开发（VSTO）"工作负载。
- **框架**：.NET Framework 4.8 开发包。
- **Outlook**：需安装 Outlook 桌面版（调试时 F5 启动 Outlook）。
- **NuGet**：依赖通过 `packages.config` 管理，构建前自动或手动还原 `packages/` 目录。

## 2. 项目约定

- `Option Strict Off`、`Option Explicit On`、`Option Infer On`（见 `oWorkhelper.vbproj`）。
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
| `<pdf> --ocr` | 本地优先 + OCR 兜底 |
| `<pdf> --force-ocr` | 强制 OCR |
| `--save-baidu-config` | 用环境变量写入本机加密配置 |
| `--parse-ocr-json <文件>` | 解析脱敏 OCR JSON |
| `--compare-expected <文件>` | 回归对比（json/csv 期望文件） |

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

工具仅打印脱敏 AK（如 `ab****yz`），不打印完整密钥。

## 5. Outlook 端到端测试

在 Outlook 环境中人工验证的关键项：

1. 加载项随 Outlook 启动正常加载，Ribbon 显示"工作助手"。
2. 点击"归档"正常执行完整流程。
3. 点击"设置"打开设置窗口，保存后直接关闭。
4. 内网版本隐藏在线解析 UI，外网版本显示。
5. 版本号显示正确（当前正式版本为 `1.2.0`）。
6. 进度窗口正常显示并更新。
7. 单封邮件失败不影响其他邮件处理。

## 6. 关键设计决策

### 归档运行锁

`ArchiveRunGuard` 使用 `Interlocked.CompareExchange` 原子获取运行锁，`ArchiveRunToken` 通过 `Using` 语义保证释放。预检查不检查运行状态（避免自我阻断）。

### 识别管道

`RecognitionPipeline` 先本地后在线：
- 本地识别（`LocalTextInvoiceRecognizer`）→ 非滴滴则委派 `GeneralInvoiceLocalRecognizer`
- 本地 Success/PartialSuccess → 直接使用
- 本地 NeedsOcr + 允许在线 → `BaiduOcrInvoiceRecognizer`
- 多 PDF 识别结果通过 `InvoiceRecognitionMerger` 合并

### 常规发票候选评分

`GeneralInvoiceLocalRecognizer` 对每个字段产生多个候选，由 `GeneralInvoiceCandidateScorer` 按位置、格式、长度等评分择优。商品明细通过 `PdfTableRegionDetector` + `GeneralInvoiceLineItemParser` 解析。

### 默认安全策略

未定义 `INTERNET_BUILD` 则在线解析禁用。`BaiduOcrInvoiceRecognizer` 在 `OnlineParserEnabled=False` 时直接拒绝请求。内网版即使配置文件残留在线参数也不调用 OCR。

### Secret Key 保护

`SecretProtector` 使用 Windows DPAPI（CurrentUser 作用域）加密 Secret Key，存储带 `DPAPI:` 前缀。设置窗口打开时自动迁移明文为密文。

## 7. 样例与脱敏

- 样例 PDF 放 `sample/`（`.gitignore` 已忽略），不提交真实票据。
- OCR 脱敏返回放 `ocr-raw/`（`.gitignore` 已忽略）。
- 回归期望文件与真实样例均不提交。
- 校准 `BaiduInvoiceFieldMapper` 时对照脱敏 JSON 更新映射表。

## 8. 诊断脚本

`tools/OutlookResiliency/` 包含两个 PowerShell 诊断脚本：
- `check_iworkhelper_addin_registration.ps1`：检查加载项注册状态。
- `get_iworkhelper_details.ps1`：获取加载项详细信息。
