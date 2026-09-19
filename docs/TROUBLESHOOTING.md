# 故障排查与诊断

## 1. Outlook 加载项相关

### 加载项被禁用

| 现象 | 原因 | 处理 |
|------|------|------|
| 启动时提示"加载项导致启动缓慢，已被禁用" | Outlook Resiliency 检测启动超时(>1000ms) | 见下方恢复步骤 |
| 启用后下次仍被禁用 | 启动代码仍然耗时 | 查看日志中 Startup 总耗时 |

### 恢复被禁用的加载项

1. Outlook → 文件 → 选项 → 加载项
2. 底部"管理"下拉选"已禁用项目" → 转到
3. 选择 oWorkHelper → 启用
4. 重启 Outlook

也可通过注册表恢复：打开 `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems`，删除包含 iWorkHelper 的条目。

### 启动性能

- `ThisAddIn.Startup` 已优化为轻量初始化（仅初始化性能跟踪器）
- `StartupPerformanceTracker` 记录各阶段耗时到日志
- 诊断脚本位于 `tools/OutlookResiliency/` 目录：
  - `check_iworkhelper_addin_registration.ps1` — 检查加载项注册状态
  - `get_iworkhelper_details.ps1` — 获取加载项详细信息

## 2. 配置问题

| 现象 | 原因 | 处理 |
|------|------|------|
| "尚未配置归档目录" | 未设置归档目录 | 打开设置选择归档文件夹 |
| "归档目录不存在" | 目录被删除 | 重选目录或保存时同意创建 |
| "归档目录无写入权限" | 目录只读 | 更换目录或检查权限 |
| "归档目录路径格式非法" | 路径含非法字符 | 重选有效目录 |
| "配置不完整" | 启用OCR但缺AK/SK | 填写API Key和Secret Key |
| "接口/Token地址无效" | 非http(s)地址 | 填写正确地址 |
| "超时时间需在3000-120000" | 超时值不合理 | 设为合理范围 |
| 测试OCR提示密钥无效 | AK/SK错误或已重置 | 核对更新密钥 |
| 测试OCR提示网络异常 | 无网络/代理 | 检查网络 |
| 换机器后OCR失效 | DPAPI仅当前用户可解密 | 重新输入Secret Key |

## 3. 归档运行问题

| 现象 | 原因 | 处理 |
|------|------|------|
| "未选择邮件" | 未在Outlook选中邮件 | 先选中邮件 |
| "已跳过（无PDF附件）" | 邮件无PDF | 正常 |
| "PDF合并失败" | PDF加密/损坏 | 检查该附件，其他邮件不受影响 |
| "需要配置OCR" | 本地字段不足且OCR未配置 | 配置在线OCR |
| 文件名含多个下划线 | 非法字符被清理 | 正常，Windows限制 |
| 同名文件未覆盖 | 自动追加(1)(2)序号 | 正常 |
| "已有归档任务正在运行" | 确有任务未结束 | 等待完成 |

### 分流处理说明

| 现象 | 说明 |
|------|------|
| 滴滴邮件发票+行程单合并 | 正常，滴滴按合并归档 |
| 普通发票每张单独归档 | 正常，常规发票单独归档 |
| PDF命名为"未识别_原名.pdf" | 未能识别为发票，按未识别归档 |
| 一封邮件多个归档文件 | 正常（多张常规发票/含未识别PDF） |
| 普通发票被误当未识别 | 缺发票特征，可用OfflineTester --classify复核 |

## 4. 识别问题

- 本地识别不稳定：多为图片型PDF或异版式，启用OCR兜底
- OCR成功但字段空：版式不受支持，文件仍会归档
- 坐标行重建已对同版式滴滴行程单稳定
- **多行程单只取首条行程用于命名**：一份行程单声明多笔行程（如“共3笔行程”）时，
  当前只把**首条行程**的起点/终点/金额用于文件命名（表头级“合计”金额仍会完整解析）。
  声明笔数与实际解析条数不一致时，`LocalTextInvoiceRecognizer` 会写入结果消息并记 `Warn` 日志
  （日志中搜索“行程单声明”即可定位）。需要严格区分多笔行程时应人工核对原 PDF。
  另外，滴滴判定已收紧：仅销售方名称或备注里出现“滴滴/行程单/网约车”不会再被误判为行程单，
  需要“上车时间 / 共N笔行程 / 行程起止 / 行程明细 / 起点+终点 / 表头短行含行程单”等证据。

## 5. 离线诊断工具 (OfflineTester)

```
OfflineTester.exe --selftest                     # 内置自测（当前 100 项）
OfflineTester.exe --preflight <归档目录>           # 预检查
OfflineTester.exe --simulate-error <错误码|list>   # 查看错误文案
OfflineTester.exe <pdf> --local-only              # 本地识别诊断
OfflineTester.exe <pdf> --general-invoice         # 常规发票诊断
OfflineTester.exe <pdf> --classify                # 分类诊断
OfflineTester.exe --dump-local-debug              # 转储本地解析诊断
```

> `--selftest` 打印的“通过 N，失败 0”会随用例增加而变化，以实际输出为准；
> 发布门禁要求**失败数为 0**（不锁死通过数）。

## 6. 日志与报告位置

- 日志：`%AppData%\iWorkHelper\logs\yyyy-MM-dd.log`（固定写入当前用户 AppData；不再写入归档目录，
  以免多个用户共用共享/网络归档目录时争用同一个 `yyyy-MM-dd.log`）
- 报告：`%AppData%\iWorkHelper\logs\archive-report-yyyyMMdd-HHmmss.txt`
  （`ArchiveReportWriter` 使用同一个日志目录，因此日志目录变化后报告位置随之变化）
- 汇总弹窗显示报告路径和日志路径
- 日志不含完整 AK/SK/token（AK/SK 仅以 `前2位****后2位` 的掩码形式出现）
