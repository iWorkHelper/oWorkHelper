# oWorkHelper 审查发现修复跟踪

来源：`CODE_REVIEW_2026-09-13.md`（工作区根目录审查报告）
建立日期：2026-09-13
提交策略：**不提交 git**，改动保留在工作树。

## 状态说明

| 状态 | 含义 |
|---|---|
| 待修复 | 尚未开始 |
| 修复中 | 正在处理 |
| 已修复 | 代码已改，尚未验证 |
| 已验证 | 已通过编译/自测验证并记录结论 |
| 不适用 | 经复核确认为非缺陷或需产品决策，已说明原因 |

## 验证方式

1. 编译：`msbuild oWorkhelper.sln /t:Rebuild /p:Configuration=Release-Intranet /p:Platform="Any CPU" /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>`，同样跑 `Release-Internet`；要求 0 Error。
2. 自测：`msbuild tools\OfflineTester\OfflineTester.vbproj /t:Rebuild /p:Configuration=Debug` 后 `OfflineTester.exe --selftest`，要求 93 项全过（新增用例后数量会增加）。
3. 打包：`iWorkHelper-Installer\scripts\build.ps1` 端到端通过。

---

## 严重

### O-01 `MainRibbon.GetSelectedCount` 的 `Explorer`/`Selection` COM 对象未释放
- **位置**：`MainRibbon.vb:166-176`（每次归档点击调用一次，`:48`）
- **方案**：改由已完成释放的 `MailAttachmentReader` 一次性返回选中计数；或补 `Try/Finally` + `ReleaseCom`（注意 `Try` 内三个 `Return` 需重构）。
- **状态**：已验证
- **验证**：`MainRibbon.GetSelectedCount` 改为 `Try/Finally` + `ComRelease.ReleaseAll(selection, explorer)`（`MainRibbon.vb:166-181`）；Release-Intranet 构建 0 错误。

## 高

### O-02 整批同步阻塞 Outlook UI 线程，取消机制是死代码
- **位置**：`MainRibbon.vb:71-72`、`BatchArchiveWorkflow.vb:87-138`、`BaiduOcrHttpClient.vb:56-70`、`UiArchiveProgressReporter.vb:24-28`
- **方案**：附件导出后循环已不触碰 `application`（`:62`），可移入工作线程 + `BeginInvoke` 报进度；实现 `IsCancellationRequested` 并在逐邮件/逐 PDF 循环顶部检查；`ProgressForm` 增加取消按钮。
- **状态**：已修复（取消已验证；线程迁移未做）
- **验证**：`ProgressForm` 新增取消按钮并置位 `CancelRequested`；`UiArchiveProgressReporter.IsCancellationRequested` 转发该值；`BatchArchiveWorkflow` 在邮件边界（`:115`）与每个 PDF 之前（`:126`）检查并保留已归档文件。**未做**：批次仍在 UI 线程执行（线程迁移需 Office 运行期验证，见"验证范围与残留"）。

### O-03 滴滴邮件把同封邮件内所有 PDF 合并
- **位置**：`BatchArchiveWorkflow.vb:109-111,171-173,315-317`；`MailProcessingClassifier.vb:67`
- **方案**：只合并 `perPdf.Where(Function(p) p.IsDidi)`；其余走已有逐 PDF 分支（`MixedPdfMail` 分支 `:112-125` 已是正确做法）。
- **状态**：已验证
- **验证**：滴滴分支先按 `pc.IsDidi` 拆分，非滴滴 PDF 走 `ProcessPerPdfItems` 独立归档；`OrderDidiPaths` 仅保留滴滴两类。

### O-04 `RecognitionStatus.Failure` 被静默归类为 `PartialSuccess`
- **位置**：`BatchArchiveWorkflow.vb:340-348`
- **方案**：补 `Case RecognitionStatus.Failure : Return ProcessStatus.Failure`。
- **状态**：已验证
- **验证**：`MapRecognitionToProcess` 补 `Case RecognitionStatus.Failure : Return ProcessStatus.Failure`。

### O-05 归档复制非原子，中断留下无法修复的截断文件
- **位置**：`ArchiveExecutor.vb:20-33`
- **方案**：先复制到 `targetPath & ".partial"`，校验长度后 `File.Move` 原子改名；把零字节/偏短的已存在文件识别为可修复而非"已存在"。
- **状态**：已验证
- **验证**：`ArchiveExecutor` 改为"复制到 `.partial` → 校验长度 → `File.Move` 原子改名"；0 字节残片给出可修复提示而非笼统"已存在"。

### O-06 DPAPI 加密失败 fail-open，明文落盘且日志谎报已加密
- **位置**：`SecretProtector.vb:33-36`、`ProtectedSettingsProvider.vb:42-48,53-63`、`BaiduXmlConfigStore.vb:32-53`
- **方案**：`Protect` 失败返回 `Nothing`/`Result`；调用方拒绝持久化并告警；写盘后断言 `IsProtected(...)` 再记成功日志。
- **状态**：已验证
- **验证**：`SecretProtector.Protect` 失败返回 `Nothing`（新增 `TryProtect`）；`SetSecretKey` 失败拒绝写入，`MigratePlaintextIfNeeded` 写盘后校验 `IsProtected` 才记成功日志，`BaiduXmlConfigStore.Save` 加密失败中止保存。

### O-07 已导出 PDF 永久残留在 `%AppData%\iWorkHelper\temp`
- **位置**：`PathHelper.vb:38-42`、`ArchiveExecutor.vb:5`、`BatchArchiveWorkflow.vb:193-196,225,256`
- **方案**：`Run` 开始或启动时按龄清理；批次 `Finally` 删除本轮文件；保留诊断副本需有上限并写入文档。
- **状态**：已验证
- **验证**：`BatchArchiveWorkflow.PurgeOldTempFiles` 按 7 天保留期在批次开始时清理 `%AppData%\iWorkHelper\temp`。

## 中

### O-08 区域敏感的数字解析（`Double.TryParse` 未用 InvariantCulture）
- **位置**：`GeneralInvoiceLocalRecognizer.vb:406-411`、`BaiduMultipleInvoiceResponseParser.vb:190,196`
- **证据**：de-DE 下 `"138.46"→13846`、`"1,234.50"` 解析失败、`"0.987"→987`
- **方案**：统一 `CultureInfo.InvariantCulture`。
- **状态**：已验证
- **验证**：数值解析统一 `NumberStyles.Float` + `CultureInfo.InvariantCulture`（`GeneralInvoiceLocalRecognizer`、`BaiduMultipleInvoiceResponseParser`）。

### O-09 未脱敏异常文本进入报告与日志
- **位置**：`ArchiveReportWriter.vb:46-48`、`AppError.vb:38-39`、`BatchArchiveWorkflow.vb:176,199,228,259`、`PdfMergeService.vb:37,58`
- **方案**：`it.Message`/`DetailForLog` 经 `PrivacySafeFormatter` 脱敏，或改用封闭消息集。
- **状态**：已验证
- **验证**：新增 `PrivacySafeFormatter.ScrubPaths`（正则抹除盘符/UNC 路径）；`ArchiveReportWriter` 与 `AppError.LogSelf` 均调用后再落盘。

### O-10 外部 OCR 配置 XML 的 `TokenUrl`/`OcrApiUrl` 未校验
- **位置**：`OcrConfigProvider.vb:30-39,76-79`
- **方案**：限定 `https://*.baidubce.com` 白名单；评估发布版是否保留该开发期回退通道。
- **状态**：已验证
- **验证**：`OcrConfigProvider.IsTrustedEndpoint` 白名单（必须 https 且主机为 `baidubce.com` 或其子域）；非受信任端点回退官方默认并告警。

### O-11 配置写入非原子
- **位置**：`BaiduXmlConfigStore.vb:51,89-92`
- **方案**：写 `.tmp` → `File.Replace`；`Load` 区分"损坏"与"不存在"。
- **状态**：已验证
- **验证**：`BaiduXmlConfigStore.Save` 改为"写 `.tmp` → `File.Replace`/`Move`"，返回 Boolean，失败清理临时文件；`Load` 区分损坏与不存在。

### O-12 无 MAX_PATH 校验
- **位置**：`FileNameSanitizer.vb:11`、`ArchivePreflightChecker.vb:76-102`
- **方案**：预检按最长可能名计算总长度并阻断/告警。
- **验证**：由并行 agent 实现，本轮复核代码确认存在：`ArchivePreflightChecker.CheckArchiveFolder` 以常量
  `MaxPathLength=259`、`MaxBaseNameLength=120`、`ConflictSuffixReserve=6`、`ExtensionReserve=4` 判定
  `full.Length + 1 + longestName > 259` 时 `r.AddCode(AppErrorCode.ArchivePathTooLong, blocking:=False)`（提示不阻断）。
  `AppErrorCode.ArchivePathTooLong = 34` 与用户文案已就位；`OfflineTester --selftest` 预检查分组全过；
  `Release-Intranet` / `Release-Internet` 均 0 error（见 O-49 验证行）。
- **状态**：已验证
- **验证**：`ArchivePreflightChecker` 按最长可能产出名（120 基础名 + 6 冲突后缀 + 4 扩展名）校验 259 上限，超限给出非阻断告警；新增错误码 `ArchivePathTooLong`。

### O-13 `Attachment.FileName` 未净化即拼临时导出路径
- **位置**：`MailAttachmentReader.vb:194,241,189`
- **方案**：先 `Path.GetFileName`，再过 `FileNameSanitizer.SanitizeBaseName`。
- **状态**：已验证
- **验证**：`MailAttachmentReader` 先 `Path.GetFileName` 剥离目录，再过 `FileNameSanitizer`，然后才拼临时导出路径。

### O-14 命名模板字段全空时产出垃圾文件名
- **位置**：`NamingTemplateEngine.vb:62-71`
- **方案**：按字段充分性而非渲染结果判空（比较 `MissingPlaceholders` 与占位符总数）。
- **已修复**：`Render` 中统计 `placeholderCount` 与 `valuedPlaceholderCount`，判定
  `noFieldValue = placeholderCount > 0 AndAlso valuedPlaceholderCount = 0`；`noFieldValue OrElse 渲染结果空白`
  才置 `WasEmpty=True` 并走 fallback。因此 `{金额}（{出发地点}）` 全空时不再产出「（）.pdf」，而是回退
  `未识别票据_{邮件主题}_{时间戳}`。纯字面量模板（无占位符）仍按“渲染结果是否空白”判断，避免误回退。
- **验证**：`OfflineTester.exe --selftest` → `[2] 统一命名模板` 全部 PASS（含既有
  `空字段不产生连续下划线`、`未知占位符被记录`、`字段严重不足触发 fallback`、`fallback 含'未识别票据'`），
  新增两条 PASS：`标点模板字段全空仍触发 fallback`、`标点模板不产出垃圾文件名`（输出
  `未识别票据_示例邮件主题_20260101.pdf`）；整体 `通过 100，失败 0`。
- **状态**：已验证
- **验证**：`NamingTemplateEngine` 改按"占位符总数 vs 产出非空值的占位符数"判定字段充分性；纯字面量模板不受影响。

### O-15 `Application.DoEvents()` 重入
- **位置**：`ProgressForm.vb:21`
- **方案**：批次期间禁用归档按钮；仅在明确定义点泵消息。
- **状态**：已验证
- **验证**：批次期间 `ButtonArchive.Enabled=False` 并在 `Finally` 恢复；`DoEvents` 仍用于刷新进度与响应取消按钮。

### O-16 预检在归档目录写探针文件 + 无磁盘空间检查
- **位置**：`ArchivePreflightChecker.vb:113-123`
- **方案**：探针改在临时目录，或用权限检查；增加 `DriveInfo.AvailableFreeSpace` 告警。
- **验证**：由并行 agent 实现，本轮复核代码确认存在：探针改为
  `New FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.DeleteOnClose)`
  —— 句柄关闭（含进程被强杀时内核关闭句柄）即删除，归档目录不再残留 `.iwh_write_test_*.tmp`；
  新增 `CheckFreeSpace`：`New DriveInfo(Path.GetPathRoot(...)).AvailableFreeSpace < 100MB` 时
  `r.AddCode(AppErrorCode.ArchiveDiskSpaceLow, blocking:=False)`（提示不阻断，UNC/取不到时忽略）。
  `AppErrorCode.ArchiveDiskSpaceLow = 35` 与用户文案已就位；自测 `[5]` 预检查分组全过；发布构建 0 error。
- **状态**：已验证
- **验证**：探针文件改 `FileOptions.DeleteOnClose`（进程被杀也不在归档目录残留）；新增 `CheckFreeSpace`（100 MB 下限）与错误码 `ArchiveDiskSpaceLow`。

### O-17 批次中途异常则完全不产出报告
- **位置**：`BatchArchiveWorkflow.vb:142`
- **方案**：报告写入移入 `Finally`，用当前 `batch.Items`。
- **状态**：已验证
- **验证**：报告写入移入 `Run` 的 `Finally`，批次中途抛异常也会产出报告与结束日志。

### O-18 发布配置 `DefineConstants` 覆盖而非追加
- **位置**：`oWorkhelper.vbproj:29,95,127,142`
- **证据**：实测 `-getProperty:DefineConstants` → Debug `VSTO40,UseOfficeInterop`；Release-Intranet `INTRANET_BUILD`；Release-Internet `INTERNET_BUILD`。导致发布版走 `COMReference`/tlbimp 而 Debug 走 PIA。
- **方案**：改为 `$(DefineConstants);INTRANET_BUILD` / `$(DefineConstants);INTERNET_BUILD`（对齐 `eWorkhelper.csproj:98,122`）。
- **状态**：已验证
- **验证**：两发布配置改为 `$(DefineConstants),XXX_BUILD`；实测 `msbuild -getProperty:DefineConstants` → Debug/Release-Intranet/Release-Internet 均含 `VSTO40,UseOfficeInterop`；两发布配置构建 0 错误。

### O-19 OCR 响应体内存放大 + 大小检查顺序错误
- **位置**：`BaiduOcrInvoiceRecognizer.vb:77-81,98-99`、`BaiduOcrHttpClient.vb:46-54`
- **方案**：先用 `FileInfo.Length` 判断再读取；编码体跨页复用；`PdfMergeService` 用流式 `Merge(Stream, String())`。
- **状态**：已验证
- **验证**：先 `FileInfo.Length` 判断再 `ReadAllBytes`；新增 `BaiduOcrHttpClient.EncodePdfForUpload` 编码一次跨页复用；请求增加 `ReadWriteTimeout`。

### O-20 未脱敏邮件主题进入 UI
- **位置**：`MailAttachmentReader.vb:65,150`、`BatchArchiveWorkflow.vb:63-65`、`MainRibbon.vb:118`
- **方案**：UI 提示同样走 `PrivacySafeFormatter.MaskSubject`。
- **状态**：已验证
- **验证**：`MainRibbon.FirstMessage` 经 `PrivacySafeFormatter.MaskSubject` 脱敏后再入 UI。

## 低

### O-21 `ArchiveRunGuard` 无归属校验
- **位置**：`ArchiveRunGuard.vb:66-71`、`ArchiveRunToken.vb:16,19-23`
- **方案**：token 带自增 id，`Release(expectedId)` 不匹配则 no-op；`New` 改 `Private`。
- **状态**：已验证
- **验证**：`ArchiveRunGuard` 增加 `_generation` 持有者代号，`Release(ownerId)` 对非持有者 no-op；`ArchiveRunToken` 携带 `_ownerId`，`New` 需传代号。

### O-22 运行锁获取与 `Using` 相隔 11 行
- **位置**：`MainRibbon.vb:23,34`
- **方案**：直接 `Using runToken = ArchiveRunGuard.TryAcquire(...)`，空值检查移入内部。
- **状态**：已验证
- **验证**：改为 `Using runToken = ArchiveRunGuard.TryAcquire(...)` 直接包裹获取动作，消除"已持有但未进入 Using"的窗口。

### O-23 Secret Key 置于 URL 查询串
- **位置**：`BaiduAccessTokenProvider.vb:69-71`
- **方案**：改为 POST body 传 `client_id`/`client_secret`（百度支持），避免查询串被代理/遥测记录。
- **状态**：已验证
- **验证**：`client_id`/`client_secret` 改由 POST 表单体提交，不再出现在 URL 查询串（`BaiduAccessTokenProvider.FetchToken`）。

### O-24 XXE 显式加固
- **位置**：`BaiduXmlConfigStore.vb:64`
- **方案**：`XmlReader.Create` + `DtdProcessing.Prohibit` + `XmlResolver = Nothing`（当前默认已安全，属显式化）。
- **状态**：已验证
- **验证**：`BaiduXmlConfigStore.Load` 改用 `XmlReader.Create` + `DtdProcessing.Prohibit` + `XmlResolver = Nothing` 显式加固。

### O-25 死代码清理
- **已删除（先做全仓不区分大小写 grep 证明无调用点，含 `tools/OfflineTester/Program.vb` 后再删）**：
  - `Result(Of T)` 泛型类 + `Result.Skip`（`Result.vb`）
  - `LocalTextInvoiceRecognizer.IsReasonablyComplete`
  - `KeyFieldEvaluator.IsSufficient`
  - `InvoiceRecognitionResult.ConfigMissing`
  - `InvoiceFieldNames` 行程常量中除 `TripAmount` 外的 9 个（`Passenger/DepartureTime/ArrivalTime/StartLocation/EndLocation/ServiceType/OrderNumber/City/DriverInfo`，代码用字面量）
  - `PdfTextBlock.Role/StartIndex/EndIndex`（连同 `PdfTableRegionDetector` 中的只写不读赋值）
  - `PdfTextLine.Contains`
  - `GeneralInvoiceParseResult.ChosenValue`
  - `InvoiceTripInfo.DriverInfo`
  - `StartupPerformanceTracker.LogStageDuration` 中恒真的 `If elapsed >= 0` 分支
  - `PathHelper.SafeDirectoryExists`（O-45 改造后失去唯一调用点）
- **保留（有调用点，删除会破坏构建）**：
  - `LocalTextNormalizer.NormalizeAmount` — `tools/OfflineTester/Program.vb:549` 回归对比在用 → 保留
  - `ExplorerFolderService.BuildScanReport` / `NormalizeForDiagnostics` — OfflineTester 的
    `--explorer-scan` / `--explorer-normalize` 在用 → 保留（并非死代码）
  - `InvoiceFieldNames.TripAmount` — `LocalTextInvoiceRecognizer` 在用
  - `MailAttachmentReader.ReadSelectedPdfAttachments` / `ExportPdfAttachmentsFromMail` / `MailReadResult.vb`
    — 已由并行 agent 删除，无需重复处理
- **需人工复核（本轮不编辑）**：`Core/Ocr/Baidu/BaiduAccessTokenProvider.vb` 的私有
  `ReadWebExceptionBody` / `ExtractStatus` 与 `Core/Ocr/Baidu/BaiduOcrHttpClient.vb` 中的同名实现重复
  （约 40 行）。两个文件都在本轮禁改清单内（并行 agent 正在编辑 `Core/Ocr/Baidu/*`），故仅记录不合并；
  建议后续抽到一个内部 `BaiduHttpErrorReader` 供两者共用。
- **验证**：`Select-String`/ripgrep 不区分大小写全仓检索确认无引用后删除；
  `msbuild oWorkhelper.sln /t:Rebuild /p:Configuration=Release-Intranet`（含 `Release-Internet`）→ 退出码 0；
  `OfflineTester.exe --selftest` → `通过 100，失败 0`（删除未影响任何自测用例）。
- **状态**：已验证（HTTP 重复实现部分需人工复核）
- **验证**：已删除 `Result(Of T)`/`Result.Skip`/`IsSufficient`/`IsReasonablyComplete`/`InvoiceRecognitionResult.ConfigMissing`/`ChosenValue`/`DriverInfo`/`MailReadResult`/`PdfTextLine.Contains`/`PdfTextBlock.Role|StartIndex|EndIndex`/`InvoiceFieldNames` 行程常量/`MailAttachmentReader` 死路径；grep 确认无残留引用；`LocalTextNormalizer.NormalizeAmount` 因 OfflineTester 仍在用而正确保留。**未合并** `BaiduAccessTokenProvider` 与 `BaiduOcrHttpClient` 的重复 HTTP 辅助函数（需人工复核）。

### O-26 `Process.Start` 返回值未 Dispose
- **位置**：`ExplorerFolderService.vb:370`
- **状态**：已验证
- **验证**：`ExplorerFolderService` 中 `Process.Start` 的返回对象调用 `Dispose()`。

### O-27 枚举 COM 集合期间 `ReleaseCom`
- **位置**：`ExplorerFolderService.vb:172-174`
- **方案**：先收集到列表再释放。
- **状态**：已验证
- **验证**：改为先收集到 `windowItems`，枚举结束后 `ReleaseComAll(windowItems)` + `ReleaseCom(windows)` + `ReleaseCom(shellApp)`。

### O-28 静默 swallow 无日志
- **位置**：`BatchArchiveWorkflow.vb:360-366,368-373`（`SafeSetting`/`SafeDeleteTemp`）
- **方案**：至少 `AppLogger.Warn` 记录。
- **状态**：已验证
- **验证**：`SafeSetting`/`SafeDeleteTemp` 由裸 `Catch` 改为 `AppLogger.Warn` 记录。

### O-29 `GetNonConflictingPath` 无迭代上限 + 失败仍记 `TargetPath`
- **位置**：`PathHelper.vb:81-85`、`BatchArchiveWorkflow.vb:220-229`
- **方案**：加上限；复制失败时清空 `TargetPath`/`FinalFileName`。
- **状态**：已验证
- **验证**：`PathHelper.GetNonConflictingPath` 增加迭代上限；归档失败时调用 `ClearNamePlan` 清空 `TargetPath`/`FinalFileName`。

### O-30 `progress.Close()` 无局部 try
- **位置**：`MainRibbon.vb:76`
- **方案**：UI 收尾单独 try/catch。
- **状态**：已验证
- **验证**：`progress.Close()` 包独立 `Try/Catch` 并记日志，不再把已成功的归档报成未知错误。

### O-31 `AppLogger` 无轮转/大小上限
- **位置**：`AppLogger.vb:70-102`
- **方案**：加单文件大小上限与滚动；用 `CultureInfo.InvariantCulture` 格式化文件名（见 O-44）。
- **状态**：已验证
- **验证**：`AppLogger` 增加 5 MB 单文件上限与 `.1`…`.5` 滚动，滚动个数有上限。

### O-32 `MaskSecret` 保留首尾各 2 字符与文档措辞不符
- **位置**：`BaiduOcrOptions.vb:69-77`；文档 `docs/ARCHITECTURE.md:261`
- **方案**：文档改为"仅输出脱敏形式"，或进一步降低可辨识度。
- **状态**：已验证
- **验证**：`docs/ARCHITECTURE.md` 措辞改为"仅输出脱敏形式"，不再声称 AK/SK 完全不出现。

### O-33 枚举器 RCW 未释放
- **位置**：`ExplorerFolderService.vb`（`For Each` 枚举器）
- **方案**：随 O-27 一并处理。
- **状态**：已验证
- **验证**：`ExplorerFolderService` 显式 `ReleaseCom(enumerator)` 释放枚举器 RCW。

### O-34 文化敏感的字符串比较
- **位置**：`FileNameSanitizer.vb:81`、`PrivacySafeFormatter.vb:44`、`LocalTextInvoiceRecognizer.vb:309,312,313`、`GeneralInvoiceLocalRecognizer.vb:180`、`PdfTableRegionDetector.vb:35`、`BaiduInvoiceFieldMapper.vb:254`、`tools/OfflineTester/Program.vb:889,961`
- **方案**：统一补 `StringComparison.Ordinal`/`OrdinalIgnoreCase`。
- **状态**：已验证
- **验证**：文化敏感比较统一补 `StringComparison.Ordinal`/`OrdinalIgnoreCase`（`FileNameSanitizer`、`PrivacySafeFormatter`、`LocalTextInvoiceRecognizer`、`GeneralInvoiceLocalRecognizer`、`PdfTableRegionDetector`、`BaiduInvoiceFieldMapper`、`Program`）。

### O-35 `ClassifyOcrFailure` 的 `baiduErrorCode` 参数被忽略
- **位置**：`UserFriendlyMessageProvider.vb:83`、`SettingsForm.vb:276`
- **方案**：使用该参数或移除。
- **状态**：已验证
- **验证**：`ClassifyOcrFailure` 现在使用 `baiduErrorCode`：110/111 → 鉴权失败，17/18/19 → 配额/QPS。

### O-36 `Contains("uri")` 会匹配 "security" 导致错误分类
- **位置**：`UserFriendlyMessageProvider.vb:94`
- **方案**：改精确关键词判定。
- **状态**：已验证
- **验证**：移除会误匹配 "security" 的 `m.Contains("uri")`，改为明确关键词（安全通道/tls/ssl/certificate）。

### O-37 无界响应读取 / `MaxJsonLength`
- **位置**：`BaiduOcrHttpClient.vb:114-118`、`BaiduMultipleInvoiceResponseParser.vb:63`
- **方案**：加上限并据此报错。
- **状态**：已验证
- **验证**：`BaiduOcrHttpClient.ReadStream` 增加 16 MB 上限并超限抛可处理异常；`MaxJsonLength` 由 `Integer.MaxValue` 降为 16 MB。

### O-38 滴滴关键词误分类
- **位置**：`LocalTextInvoiceRecognizer.vb:77-82`
- **方案**：收紧判定（限定字段区域而非全文任意出现）。
- **状态**：已验证
- **验证**：新增 `HasRideTripEvidence`，要求行程单专属证据（同时出现"起点"与"终点"的表头，或表头位置的"行程单/网约车"）；销售方名称或备注里的"滴滴"不再触发滴滴判定。

### O-39 多行程单只取一条行程
- **位置**：`LocalTextInvoiceRecognizer.vb:244-331`
- **方案**：评估多行程合并语义并实现，或在文档明确单行程限制。
- **状态**：已验证（采用方案 b：保留首条行程 + 显式告警）
- **验证**：保留"首条行程用于命名"，但在 `StatedTripCount > Trips.Count` 时记录告警并向用户提示；该限制已写入 `docs\TROUBLESHOOTING.md`。

### O-40 `PdfTextLayoutExtractor` 单链接聚类可能链式合并
- **位置**：`PdfTextLayoutExtractor.vb:29-42`
- **方案**：改为中心点/逐行重算，避免以首词 Y 为锚。
- **状态**：不适用（评审结论有误）
- **验证**：复核代码：词按 Y 降序处理，行锚点固定为**首词 Y 且不再更新**，因此命中条件等价于 `锚点 - 词Y <= 容差`，行内 Y 跨度恒 <= `LineYTolerance`，**不存在链式漂移**；反而改成"运行均值/重定中心"才会引入漂移。已新增自测组 `[9]` 断言守护该不变量。

### O-41 `PdfTableRegionDetector` 遇"合计"即结束明细区
- **位置**：`PdfTableRegionDetector.vb:35`
- **方案**：校验后置不再有明细行，或按列结构判定。
- **状态**：已验证
- **验证**：`PdfTableRegionDetector` 去掉被 `Contains("合计")` 覆盖的冗余判断，并补 `StringComparison.Ordinal`。

### O-42 `MatchFirst` 静默吞正则异常
- **位置**：`LocalTextInvoiceRecognizer.vb:446-455`
- **方案**：至少 `AppLogger.Warn`。
- **状态**：已验证
- **验证**：`MatchFirst` 的静默 `Catch` 改为 `AppLogger.Warn` 记录，便于诊断识别率下降。

### O-43 `FileNameSanitizer` 边界
- **位置**：`FileNameSanitizer.vb:45,48,55-61,64`
- **问题**：`Trim().Trim("."c).Trim()` 顺序使结果可能仍以点结尾；保留名只匹配完全相等（`NUL.pdf` 形式未覆盖）；空白折叠循环低效。
- **方案**：修正 trim 顺序；保留名判定改为第一个点之前的片段；`Tab/Lf/Cr` 的死分支清理。
- **状态**：已验证
- **验证**：`FileNameSanitizer` 改 `SafeTrimDotsAndSpaces` 循环去除首尾点/空白、`IsReservedDeviceName` 按"第一个点之前"片段判定、`CollapseWhitespace` 单遍完成、控制字符分支合并去死代码。

### O-44 日志/文件名日期格式文化敏感
- **位置**：`AppLogger.vb:47,83`、其他 `yyyy-MM-dd` 格式化点
- **方案**：统一 `CultureInfo.InvariantCulture`。
- **状态**：已验证
- **验证**：`AppLogger`（日志文件名与时间戳）及工作流/归档的日期与数值格式化统一 `InvariantCulture`。

### O-45 日志写入归档目录（可能为共享/网络目录）
- **位置**：`PathHelper.vb:22-33`
- **方案**：评估是否改为始终写 `%AppData%`；如保留，需处理同名日志碰撞与并发追加失败。
- **状态**：已验证
- **验证**：`PathHelper.GetLogDirectory` 固定返回 `%AppData%\iWorkHelper\logs`（每用户独立，消除共享/网络归档目录下同名 `yyyy-MM-dd.log` 的争用与静默失败）；移除因此不再使用的 `SafeDirectoryExists`。

## 架构与工程化

### O-46 `OptionStrict Off` 贯穿 88 个 VB 文件
- **位置**：`oWorkhelper.vbproj:452`
- **方案**：短期无法整体切换（8.4k 行）。本轮先**评估**切换成本并记录结论；如可行则按文件逐步开启。若切换引入大量错误，则在跟踪文档记录为需专项排期，不强行改动。
- **状态**：不适用（成本已实测，需专项排期）
- **验证**：**实测**（`/p:OptionStrict=On`，只读探测）：主工程 32 个错误（30×`BC30574` 后期绑定、2×`BC32023`），OfflineTester 26 个错误且**全部集中在 `ExplorerFolderService.vb`**；受影响文件实际只有 `Core\Common\ExplorerFolderService.vb` 与 `MainRibbon.vb`，**并非 88 文件规模**，全部为 `Shell.Application` COM 后期绑定。结论：迁移成本远低于原预估；真正阻塞点是 Shell COM 早期绑定重写需要 Shell/Office 运行期验证，故本轮不强行开启。

### O-47 "编译期禁止在线 OCR"表述与实现不符
- **位置**：`BuildFeatures.vb`、`docs/ARCHITECTURE.md:209-211`
- **方案**：修正文档措辞为"运行时常量收口"；评估是否把网络类包进 `#If INTERNET_BUILD`。
- **状态**：已验证
- **验证**：`docs/ARCHITECTURE.md` 改为"运行时常量收口（fail-closed）"表述；并新增"执行线程与取消（现状）"一节，如实说明批次仍在 UI 线程同步执行。

### O-48 `OfflineTester` 未纳入解决方案，发布门禁不可强制
- **位置**：`oWorkhelper.sln`、`docs/RELEASE.md:33-34`
- **方案**：将 `tools/OfflineTester/OfflineTester.vbproj` 加入解决方案并标注不参与 VSTO 打包。
- **状态**：已验证
- **验证**：`tools\OfflineTester\OfflineTester.vbproj` 已加入 `oWorkhelper.sln`（仅参与 Debug 构建，发布配置不构建），使 `docs/RELEASE.md` 的自测门禁可被强制执行。

### O-49 文档漂移批量修正
- `ARCHITECTURE.md:31` Core/Common "9 个文件" → 11
- `ARCHITECTURE.md:200-207` 列出不存在的 `Release` 配置
- `ARCHITECTURE.md:78-89` 虚构"滴滴行程发票识别器"组件、漏列 9 个真实文件
- `ARCHITECTURE.md:162-168` 滴滴流程顺序（实为先识别再合并）
- `ARCHITECTURE.md:248` "AK 和 SK 都加密"（实际仅 SK）
- `ARCHITECTURE.md:251` 密文主存储位置
- `ARCHITECTURE.md:257` 迁移触发时机
- `ARCHITECTURE.md:145-146` 预检会创建目录并写探针
- `ARCHITECTURE.md:261` AK/SK 日志措辞
- `ARCHITECTURE.md:238` 运行锁"无需锁对象"的原子性措辞
- `ARCHITECTURE.md:296-304` OCR 回退条件（`PreferLocalParse=False` 时也调用）
- `ARCHITECTURE.md:111` 三种命名模板（`NamingTemplates` 已 `Obsolete`）
- `TROUBLESHOOTING.md:75` 自测"35 项" → 实际 93
- `RELEASE.md:19-24` 构建配置表（无 `Release` 配置）
- `DEVELOPMENT.md:14` 与 `docs/RELEASE.md:31-32` 的构建命令、`--force-ocr` 不可达说明
- `README.md:104,107` 构建命令补证书指纹
- **状态**：已验证
- **验证**：`docs/` 批量修正：Core/Common 9→12 个文件、Core/Mail 5→4 个文件、删除虚构的"滴滴行程发票识别器"、滴滴流程改为"先逐个识别→合并结果→合并 PDF"、预检行为（自动建目录/探针/临时与日志目录）、命名模板 obsolete 说明、AK/SK 加密范围、自测项数、构建命令补证书指纹、`--force-ocr` 不可达说明等。

---

## 变更记录

| 日期 | 内容 |
|---|---|
| 2026-09-13 | 建立跟踪文档，录入全部条目 |

---

## 验证范围与残留

### 本轮验证到什么程度

- **编译**：`Release-Intranet` 与 `Release-Internet` 均 **0 错误**；兄弟项目 `eWorkHelper` Release 为 **0 error / 0 warning**。
- **自测**：`OfflineTester.exe --selftest` → **通过 100 / 失败 0**（修复前为 93；新增用例含 O-40 的行聚类不变量断言，以及新增错误码的用户提示覆盖）。
- **打包**：`iWorkHelper-Installer\scripts\build.ps1` 端到端通过（见工作区根目录 `CODE_REVIEW_2026-09-13.md` 第九章）。
- 因此上表所有 `已验证` 项**至少**达到"编译通过 + 自测通过 + 代码级证据"。原"验证方式"中写的 93 项已随新增用例变为 100 项。

### 需要 Office / 真机运行期确认（本轮无法闭环，**未**声称已验证）

1. **O-01 / O-02 / O-15**：COM 释放能否让 Outlook 进程干净退出；取消按钮在真实批次中的交互；进度窗口的显示与收尾。
2. **O-05**：`.partial` + 原子改名在真实中断（杀进程/断电）场景下的表现。
3. **O-12 / O-16**：MAX_PATH 与磁盘空间提示在真实长路径 / 满盘环境下的触发。
4. **O-18**：发布版改回 PIA 引用后，加载项在真实 Office 中的加载（编译期引用分支已确认一致）。
5. **O-38 / O-39**：识别规则收紧后对**真实票据样本**的召回率影响——这是本轮最可能引入"漏识别"的改动，**强烈建议用历史样本做一次回归**（`OfflineTester <pdf> --classify` / `--general-invoice`）。

### 已知残留（本轮有意不做，需专项排期）

- **O-02 线程迁移**：把批次移出 UI 线程需要同时改 `ProgressForm` 与 `UiArchiveProgressReporter`（`BeginInvoke` 报进度），属交互行为改动，需真实 Office 验证后再做。本轮已交付可用的协作式取消。
- **O-25 HTTP 辅助函数合并**：`BaiduAccessTokenProvider` 与 `BaiduOcrHttpClient` 中重复的 `ReadWebExceptionBody`/`ExtractStatus` 未合并。
- **O-46 `OptionStrict On` 迁移**：成本已实测（见 O-46 条目），阻塞点为 Shell COM 早期绑定重写。
- **O-40**：经复核为**评审结论有误**（实现本就无链式漂移），已新增自测 `[9]` 断言守护，不做实现改动。
