# oWorkHelper 架构文档

## 1. 项目定位

oWorkHelper 是 iWorkHelper Organization 旗下的 Outlook VSTO 加载项，基于 VB.NET / .NET Framework 4.8 开发。其核心功能是批量处理邮件中的 PDF 附件，完成发票识别、PDF 合并、文件命名和归档操作。用户在 Outlook 中选中一批邮件后，一键即可将附件中的发票 PDF 按照识别结果自动命名并归档到指定目录。

## 2. 技术栈

| 组件 | 技术 |
|------|------|
| 语言 | VB.NET |
| 运行时 | .NET Framework 4.8 |
| Office 集成 | VSTO 4.0 (Visual Studio Tools for Office) |
| UI 框架 | WinForms (Ribbon + 设置窗体 + 进度窗体) |
| PDF 处理 | PdfPig 0.1.14 (文本提取、坐标行重建、表格区域检测、PDF 合并) |
| OCR 识别 | 百度 OCR 多票识别 API (仅外网版本) |
| 密钥保护 | Windows DPAPI (Data Protection API, CurrentUser 范围) |
| 构建工具 | MSBuild / Visual Studio 2022 |

## 3. 代码结构

### 顶层文件

- **ThisAddIn.vb** — VSTO 加载项入口，负责加载项生命周期管理
- **MainRibbon.vb** — Outlook Ribbon 界面，提供用户操作按钮
- **SettingsForm.vb** — 设置窗体，管理百度 OCR 配置及归档路径等选项
- **ProgressForm.vb** — 进度窗体，展示批量处理进度

### Core 模块

#### Core/Common/ (通用基础设施，12 个文件)

- **Result.vb** — 统一结果类型（`ProcessStatus` + 消息），封装操作成功/失败状态及错误信息
- **PathHelper.vb** — 路径处理辅助方法（应用数据目录、日志目录、temp 目录、唯一文件名）
- **FileNameSanitizer.vb** — 文件名清理，移除非法字符
- **PrivacySafeFormatter.vb** — 隐私安全格式化器，对邮件主题和附件名进行脱敏处理
- **ExceptionFormatter.vb** — 异常格式化，生成结构化异常信息
- **ErrorSeverity.vb** — 错误严重级别枚举
- **AppErrorCode.vb** — 应用错误代码枚举，标识各类错误场景
- **AppError.vb** — 应用错误类，结合错误代码与严重级别
- **UserFriendlyMessageProvider.vb** — 面向用户的友好错误消息生成器
- **BuildFeatures.vb** — 编译特性开关（`#If INTERNET_BUILD` → 运行时常量），控制内外网版本差异
- **ComRelease.vb** — COM 对象统一释放工具（`Marshal.ReleaseComObject` 包装）
- **ExplorerFolderService.vb** — 归档后打开或激活归档目录，使用 Shell.Application.Windows() 枚举已有窗口

#### Core/Logging/ (日志，1 个文件)

- **AppLogger.vb** — 线程安全的文件日志记录器

#### Core/Diagnostics/ (诊断，1 个文件)

- **StartupPerformanceTracker.vb** — 启动性能追踪器，记录加载项各阶段耗时

#### Core/Configuration/ (配置管理，3 个文件)

- **BaiduOcrOptions.vb** — 百度 OCR 配置选项（API Key、Secret Key、API 端点等）
- **BaiduXmlConfigStore.vb** — 基于 XML 文件的配置持久化存储
- **OcrConfigProvider.vb** — OCR 配置提供者，统一配置读取入口

#### Core/Invoice/ (发票数据模型，4 个文件)

- **InvoiceFieldNames.vb** — 发票字段名称常量定义
- **InvoiceTripInfo.vb** — 行程信息（出发地、到达地、乘车日期等，用于滴滴发票）
- **InvoiceLineItem.vb** — 发票明细行项目
- **InvoiceInfo.vb** — 发票信息聚合模型，包含金额、日期、销售方、行程明细等

#### Core/Pdf/ (PDF 处理，7 个文件)

基于 PdfPig 库实现 PDF 文本提取与处理：

- **PdfTextExtractor.vb** — PDF 文本提取器，从 PDF 中提取原始文本块
- **PdfTextLayoutExtractor.vb** — PDF 文本布局提取器，按坐标重建文本行
- **PdfTextBlock.vb** — PDF 文本块数据模型
- **PdfTextLine.vb** — PDF 文本行数据模型
- **PdfTextExtractResult.vb** — PDF 文本提取结果
- **PdfTableRegionDetector.vb** — 表格区域检测器，识别 PDF 中的表格结构
- **PdfMergeService.vb** — PDF 合并服务，将多个 PDF 文件合并为一个

#### Core/Recognition/ (识别引擎，16 个文件)

识别管线，负责从 PDF 文本或 OCR 结果中提取发票关键信息：

- **RecognitionPipeline.vb** — 识别调度器（本地优先 / 在线兜底，标记识别来源）
- **LocalTextInvoiceRecognizer.vb** — 本地文本识别器；**滴滴/行程单逻辑就在此类内**（`DetectDocumentType` + 行程单表头/表格解析），不存在独立的“滴滴行程发票识别器”
- **GeneralInvoiceLocalRecognizer.vb** — 常规增值税发票识别器（候选评分 + 分区解析 + 商品明细）
- **GeneralInvoiceCandidateScorer.vb** — 字段候选评分器
- **GeneralInvoiceFieldCandidate.vb** — 字段候选数据模型
- **GeneralInvoiceLineItemParser.vb** — 商品明细行解析器
- **GeneralInvoiceParseResult.vb** — 常规发票解析中间结果
- **KeyFieldEvaluator.vb** — 关键字段 / 命名核心字段完整度评估
- **InvoiceRecognitionMerger.vb** — 多来源识别结果合并器
- **BaiduOcrInvoiceRecognizer.vb** — 百度 OCR 识别器，调用百度多票识别 API
- **IInvoiceRecognizer.vb** — 识别器接口
- **InvoiceDocumentType.vb** — 票据类型枚举
- **InvoiceField.vb** — 扁平字段模型
- **InvoiceRecognitionResult.vb** — 识别结果与识别状态
- **RecognitionContext.vb** — 识别上下文（PDF 路径、抽取文本）
- **LocalTextNormalizer.vb** — 本地解析前的文本归一化

#### Core/Ocr/Baidu/ (百度 OCR 集成，7 个文件)

- **BaiduAccessTokenProvider.vb** — 百度 OCR 访问令牌获取与缓存
- **BaiduAccessTokenResult.vb** — 令牌请求结果
- **BaiduOcrHttpClient.vb** — 百度 OCR HTTP 客户端，发送 PDF 文件并获取识别结果
- **BaiduOcrRawResponse.vb** — 百度 OCR 原始响应数据模型
- **BaiduMultipleInvoiceResponseParser.vb** — 多票识别响应解析器
- **BaiduInvoiceTypeMapper.vb** — 百度发票类型到内部类型的映射
- **BaiduInvoiceFieldMapper.vb** — 百度发票字段到内部字段的映射

#### Core/Mail/ (邮件处理，4 个文件)

- **MailAttachmentReader.vb** — 邮件附件读取器，从 Outlook MailItem 中提取 PDF 附件
- **MailAttachmentItem.vb** — 邮件附件数据模型
- **MailPdfGroup.vb** — 单封邮件的 PDF 分组
- **MailPdfGroupingResult.vb** — 邮件 PDF 分组结果

#### Core/Archive/ (归档处理，8 个文件)

- **NamingTemplates.vb** — 命名模板配置：**只使用统一模板 `UnifiedTemplate`**（旧的三套模板 Invoice/Trip/Unknown 属性已标记 `Obsolete`，仅为兼容旧配置保留）
- **NamingTemplateEngine.vb** — 模板引擎，将占位符替换为实际识别值（按字段充分性决定是否回退）
- **ArchiveNamingRule.vb** — 统一归档命名规则 + 内置未识别 fallback
- **UnknownPdfNamingRule.vb** — 未识别 PDF 的命名规则
- **ArchivePlanner.vb** — 归档计划器，根据识别结果生成归档计划
- **ArchiveExecutor.vb** — 归档执行器，按计划将文件复制到目标目录
- **ArchiveResult.vb** — 归档结果数据模型
- **ArchiveReportWriter.vb** — 归档报告生成器，汇总处理结果

#### Core/Security/ (安全模块，2 个文件)

- **SecretProtector.vb** — 基于 DPAPI 的密钥加密/解密
- **ProtectedSettingsProvider.vb** — 受保护的配置读写提供者，自动处理加密/解密

#### Core/Workflow/ (工作流编排，12 个文件)

- 批量归档工作流编排器，协调整个处理流程
- 邮件分类器，将邮件附件分为滴滴/常规/未识别/无 PDF 四类
- 预检检查器，在流程启动前验证前置条件（目标目录存在性等）
- 运行锁（RunGuard），使用 Interlocked 原子操作防止并发执行
- 进度报告器，向 ProgressForm 推送处理进度

## 4. 核心处理流程

整个归档流程如下：

```
用户在 Outlook 中选中邮件 → 点击 Ribbon 上的归档按钮
    │
    ▼
RunGuard 获取运行锁（Interlocked 原子操作，防止重复执行）
    │
    ▼
预检检查（Preflight Check）
  - 验证是否选中了邮件
  - 验证归档目标目录：目录不存在则**自动创建**，再以“写入并删除探针文件”验证写权限
  - 验证临时目录 / 日志目录可写（日志目录不可写仅告警，不阻断）
  - 验证命名模板非空、OCR 配置是否完整（仅提示）
  - 校验完整路径长度（按最长可能产出名对比 259 上限，超限仅告警）
  - 校验归档磁盘可用空间（低于 100 MB 仅告警，不阻断）
    │
    ▼
BatchArchiveWorkflow 启动批量处理
    │
    ├─ 1. 读取邮件 PDF 附件
    │     MailAttachmentReader 从每封邮件中提取 PDF 附件
    │
    ├─ 2. 分类（Classification）
    │     将附件分为四类：
    │     - 滴滴发票（Didi）：根据文件名或内容特征识别
    │     - 常规发票（General）：一般增值税发票
    │     - 未识别（Unknown）：无法分类的 PDF
    │     - 无 PDF（NoPDF）：邮件中无 PDF 附件
    │
    ├─ 3. 识别（Recognition）
    │     │
    │     ├─ 先对每个 PDF **逐个独立识别**（识别管线）
    │     │
    │     ├─ 滴滴发票：全部 PDF 识别完成后，先合并**识别结果**
    │     │   （InvoiceRecognitionMerger），再合并 PDF 文件
    │     │   （PdfMergeService，顺序为发票在前、行程单在后），随后统一命名
    │     │   ——只合并滴滴成员 PDF；同封邮件内的非滴滴 PDF 仍各自独立归档
    │     │
    │     └─ 常规发票：逐个 PDF 独立识别
    │         └─ 识别管线
    │
    │     识别管线执行顺序：
    │     ① PdfPig 提取文本
    │     ② 本地文本识别器（正则匹配）
    │        - 滴滴：提取乘车日期、金额、出发地、到达地
    │        - 常规：提取开票日期、金额、销售方名称
    │     ③ 关键字段评估：判断本地识别结果是否充分
    │     ④ 若不充分，按策略回退到百度 OCR（外网版运行时启用；`PreferLocalParse=False` 时也会调用）
    │     ⑤ 合并多来源识别结果
    │
    ├─ 4. 命名（Naming）
    │     模板驱动的文件命名：
    │     - 滴滴发票：{乘车日期}_{金额}_{出发地点}_{到达地点}.pdf
    │     - 常规发票：{开票日期}_{金额}_{销售方名称}.pdf
    │     - 未识别：未识别_{原始文件名}.pdf
    │     NamingTemplateEngine 负责占位符替换，FileNameSanitizer 清理非法字符
    │
    └─ 5. 归档（Archive）
          ArchivePlanner 生成归档计划
          ArchiveExecutor 将文件复制到目标目录（含同名冲突解决）
          ArchiveReportWriter 生成处理报告
          ExplorerFolderService 打开或激活归档目录
          RunGuard 释放运行锁
```

### 执行线程与取消（现状）

批量归档在 Outlook 的 **UI 线程上同步执行**：`MainRibbon.ButtonArchive_Click` 直接调用
`BatchArchiveWorkflow.Run`，中途通过 `Application.DoEvents()` 泵消息，使 `ProgressForm` 能重绘并让“取消”按钮可点击。
取消采用**协作式轮询**：工作流在“邮件边界”和“每个 PDF 之前”检查 `IArchiveProgressReporter.IsCancellationRequested`
（由 `ProgressForm.CancelRequested` 提供），命中后停止处理剩余邮件并保留已归档文件，不做强制中断。
因此处理大批量邮件时 Outlook UI 仍会被阻塞（批量工作线程化见 REVIEW_TRACKING.md O-02）。

## 5. 内外网版本差异

项目通过 `BuildFeatures.vb` 作为唯一的编译特性开关，控制内外网版本差异。

### 编译配置

项目定义了 3 种构建配置（`oWorkhelper.vbproj`）：

| 配置名 | 编译常量 | 用途 |
|--------|----------|------|
| Debug | `VSTO40,UseOfficeInterop`（无版本常量） | 开发调试，默认离线模式 |
| Release-Intranet | `VSTO40,UseOfficeInterop,INTRANET_BUILD` | 内网发布版本 |
| Release-Internet | `VSTO40,UseOfficeInterop,INTERNET_BUILD` | 外网发布版本，运行时启用百度 OCR |

### 安全默认原则

`BuildFeatures.vb` 通过 `#If INTERNET_BUILD` 条件编译产生的是**运行时常量** `OnlineParserEnabled`
（未定义 `INTERNET_BUILD` 的构建为 `False`）。未定义该常量时，`BaiduOcrInvoiceRecognizer.IsAvailable()`
返回 `False`，`Recognize()` 直接返回 `ConfigurationMissing`，**不会发起任何网络请求（fail-closed）**。

需要明确的是：这**不是“编译期禁止在线 OCR”**。网络栈（`Core/Ocr/Baidu/*` 与百度 OCR 识别器）仍然被编译进内网版本，
只是被运行时常量收口。如需真正的编译期隔离，可评估把网络相关类包进 `#If INTERNET_BUILD`（尚未实施，见 REVIEW_TRACKING.md O-47）。

### 条件编译控制

`BuildFeatures.vb` 中使用 `#If INTERNET_BUILD` 条件编译指令，暴露只读属性供其他模块查询。各模块无需关心编译常量细节，统一通过 `BuildFeatures` 类判断当前运行模式。

受影响的功能点：
- 百度 OCR 在线识别功能的启用与禁用
- 设置窗体中 OCR 相关配置项的显示与隐藏
- 预检检查中 OCR 配置完整性验证的开关

## 6. 错误处理设计

### 错误建模

错误处理采用结构化设计，由三个核心类型组成：

- **AppErrorCode 枚举** — 定义所有已知错误场景的唯一标识码，每个错误码对应一种明确的失败原因
- **AppError 类** — 组合错误码与严重级别，携带上下文信息，支持链式追溯原始异常
- **UserFriendlyMessageProvider** — 将 AppError 转换为面向用户的中文友好消息，隔离技术细节与用户展示

### 预检机制

在批量处理流程启动前执行预检检查（Preflight Check），提前验证各项前置条件是否满足。预检阶段不依赖运行时状态，仅检查静态配置和环境条件。预检失败时直接向用户报告问题，不进入处理流程。

### 运行锁

RunGuard 使用 `Interlocked.CompareExchange` 原子操作实现运行锁，确保同一时刻只有一个批量处理流程在运行。
原子性只覆盖**锁字**（`_state`）与持有者代号（`_generation`）；批次标识、线程 ID、获取时间这些
**持有者诊断信息是普通字段赋值，并非原子**，仅用于日志定位（`DescribeHolder`），不参与并发判定。
释放只接受当前持有者代号，过期/伪造 token 的释放请求为 no-op。

### 逐项故障隔离

批量处理中，单个邮件或单个附件的处理失败不会中断整个流程。每个处理项独立捕获异常，记录错误信息后继续处理下一项。最终通过归档报告汇总所有成功和失败的处理结果。

## 7. 安全设计

### 密钥保护

**只有 Secret Key（SK）使用 Windows DPAPI 加密**，API Key（AK）按普通设置项明文保存：

- **加密范围**：DPAPI CurrentUser 范围，仅当前 Windows 用户可解密
- **存储格式**：密文带 `DPAPI:` 前缀，用于区分明文与密文
- **主要存储位置**：`My.Settings.BaiduSecretKey` → `user.config`（User scope）；
  `%AppData%\iWorkHelper\baidu-ocr.config.xml`（`BaiduXmlConfigStore`）是**兼容用的外部回退**，
  `OcrConfigProvider` 读取时以 `My.Settings` 优先
- **SecretProtector** 负责加密和解密操作；**ProtectedSettingsProvider** 封装读写与透明加解密
- 加密失败时**拒绝写入**（不落明文），并在 `Save()` 后通过 `VerifyPersisted()` 断言存储值确为密文

### 明文自动迁移

检测到明文遗留 Secret Key（无 `DPAPI:` 前缀）时自动加密回写。触发点有两处：

1. 打开设置窗体时（`SettingsForm`）；
2. **每次点击“归档”时**（`MainRibbon` 在预检查之前调用 `MigratePlaintextIfNeeded`），
   以免用户从不打开设置就直接归档而始终残留明文。

### 日志脱敏

- AK/SK 会以**部分掩码**形式写入日志：`BaiduOcrOptions.ToSafeSummary()` → `MaskSecret()`（保留前 2、后 2 字符），
  完整密钥与 Access Token 不写日志
- **PrivacySafeFormatter** 对邮件主题和附件名称进行脱敏处理，防止个人信息泄露到日志文件

## 8. 百度 OCR 集成

百度 OCR 集成使用百度多票识别 API（`multiple_invoice`）实现在线发票识别。
网络栈被编译进所有构建配置，是否可用由运行时常量 `BuildFeatures.OnlineParserEnabled` 收口
（仅 `INTERNET_BUILD` 版本为 `True`，见第 5 节）。

### 调用流程

```
BaiduAccessTokenProvider 获取 Access Token
  - 使用 API Key + Secret Key 换取 Token
  - Token 带有效期缓存，过期自动刷新
      │
      ▼
BaiduOcrHttpClient 发送识别请求
  - HTTP POST 请求
  - 将 PDF 文件以 Base64 编码作为 pdf_file 参数提交
  - 调用百度 multiple_invoice API 端点
      │
      ▼
BaiduMultipleInvoiceResponseParser 解析响应
  - 解析 JSON 响应体
  - 提取每张发票的识别结果
      │
      ▼
BaiduInvoiceTypeMapper 映射发票类型
  - 将百度返回的发票类型编码映射为内部发票类型枚举
      │
      ▼
BaiduInvoiceFieldMapper 映射发票字段
  - 将百度返回的字段名称映射为内部 InvoiceInfo 模型字段
  - 统一数据格式（日期格式、金额格式等）
```

### 回退策略

百度 OCR 仅在**外网版本**（`INTERNET_BUILD`）且 `BaiduOcrOptions.IsConfigured()` 为真时才有机会调用；
是否真的调用由 `RecognitionPipeline` 决定（`PreferLocalParse` / `AutoFallbackToOcr` / 疑似图片型）：

1. `PreferLocalParse=True` 且本地命名核心字段充分 → 直接采用本地结果，不调用 OCR；
2. `PreferLocalParse=False`（设置里选择“在线优先”）→ **即使本地字段充分也会调用 OCR**；
3. `AutoFallbackToOcr=True` 或文本疑似图片型 → 本地不充分时回退 OCR；
4. OCR 不可用或失败时，本地有字段则返回本地部分成功，否则返回需要 OCR 的提示。

因此“本地充分就一定不发网络请求”只在 `PreferLocalParse=True` 时成立。
