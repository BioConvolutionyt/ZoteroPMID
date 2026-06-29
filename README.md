# ZoteroPMID - 基于 PMID 自动插入 Zotero 参考文献

## 适用环境

- **Zotero** 7 / 8（已安装 Word 插件）
- **Microsoft Word** 2016+（Windows）
- 文档中以 `[PMID]` 或 `[PMID, PMID]` 格式标注引用位置（PMID 为 7-9 位纯数字）

## 文件说明

| 文件 | 用途 |
|------|------|
| `ZoteroPMID_Macro.bas` | Word VBA 宏模块（Step1 + Step3） |
| `ZoteroPMID_Zotero.js` | Zotero 映射脚本（Step2b） |

运行过程中，所有中间文件保存在 `下载\ZoteroPMID\` 文件夹中（可通过 `OUTPUT_BASE` 配置修改）：

| 中间文件 | 内容 |
|----------|------|
| `pmid_list.txt` | 从文档中提取的唯一 PMID 列表 |
| `citations.nbib` | 从 PubMed 下载的 MEDLINE 格式文献数据 |
| `pmid_mapping.txt` | PMID → Zotero 条目的映射关系 |

---

## 方案选择
提供两种可选方案
| | 方案 A：混合方案（推荐） | 方案 B：全自动方案 |
|---|---|---|
| **Zotero 集成度** | 完整（Zotero 原生创建引用） | 较低（VBA 直接创建域代码） |
| **切换引用样式** | 正常工作 | 可能需手动清除残留格式 |
| **引用高亮** | 与 Zotero 手动插入一致 | 普通 Word 域 |
| **自动化程度** | 每个引用需在 Zotero 弹窗中确认 | 全自动一键完成 |
| **所需步骤** | Step1 → 导入 → Step3（逐个） | Step1 → 导入 → Zotero JS → Step3 → Refresh |

---

## 完整操作流程

### 前置准备（两种方案通用）

1. 确保 Zotero 已运行，Word 工具栏中有 Zotero 选项卡
2. 在 Zotero 中创建好目标集合（如 `我的文库 > Research > Citation`）
3. 正确配置`ZoteroPMID_Zotero.js`中的COLLECTION_PATH参数（[见可配置参数](#可配置参数)）

---

### Step 1：提取 PMID + 下载 MEDLINE（两种方案通用）

1. 打开 Word 文档
2. `Alt+F11` 打开 VBA 编辑器
3. **插入 > 模块**，将 `ZoteroPMID_Macro.bas` 的内容粘贴进去
   - 注意：如果第一行是 `Attribute VB_Name = "ZoteroPMID"`，**删掉这行**
4. 关闭 VBA 编辑器
5. `Alt+F8` → 选择 `Step1_ExtractAndDownload` → 运行

**宏自动完成：**
- 扫描全文，提取所有 `[PMID]` 格式的引用
- 从 PubMed E-utilities 下载 MEDLINE 文献数据
- 将 PMID 列表复制到剪贴板
- 输出保存至 `下载\ZoteroPMID\`（由 `OUTPUT_BASE` 控制）

---

### Step 2：导入文献到 Zotero（两种方案通用，约 1 分钟）

**方法 A（推荐）：**
1. 打开 Zotero，左侧选中目标集合（如 `Research > Citation`）
2. 点击工具栏**魔术棒图标**（通过标识符添加条目）
3. `Ctrl+V` 粘贴 PMID 列表 → 回车
4. 等待所有文献导入完成

**方法 B（备选）：**
1. Zotero → 文件 → 导入 → 选择 `下载\ZoteroPMID\citations.nbib`
2. 导入后，将条目拖入目标集合

---

### Step 3 — 方案 A：逐个插入

> 此方案调用 Zotero 原生对话框插入引用，集成度完整。

1. **先保存文档备份**
2. `Alt+F8` → 选择 `InsertNextCitation` → 运行
3. 宏自动执行以下操作：
   - 找到文档中第一个 `[PMID]` 占位符
   - 若为多 PMID（如 `[A, B]`），自动拆分为 `[A][B]`
   - 将当前 PMID 复制到剪贴板
   - 删除占位符（含前导空格）
   - 弹出 Zotero 引用对话框
4. 在 Zotero 对话框的搜索栏中 **粘贴 PMID 或输入文献标题**，选中条目，按回车确认
5. **重复运行**宏处理下一个引用，直到提示"All done"
6. 全部完成后：点击 Zotero 工具栏 → **Add/Edit Bibliography** 插入参考文献列表

**强烈建议：** 为 `InsertNextCitation` 分配快捷键，每次只需 **按快捷键 → Zotero弹窗粘贴确认 → 按快捷键** 即可连续处理。

快捷键设置方法：
1. **文件 → 选项 → 自定义功能区**
2. 左下角点击 **键盘快捷方式：自定义**
3. 类别选 **"宏"**
4. 修改 **"将更改保存在(V)"** 为当前文档
4. 右侧找到 `InsertNextCitation`
5. 在"请按新快捷键"框中按下 `Ctrl+Shift+Z`（或其他未占用组合）
6. 点击 **指定 → 关闭**

> 注意：由于 `ZoteroAddEditCitation` 是非阻塞调用，循环自动化不可行。
> 快捷键 + 逐个确认是方案 A 唯一可靠的工作方式。

---

### Step 3 — 方案 B：全自动插入

> 此方案速度快但 Zotero 集成度较低。需额外运行 Zotero JS 映射脚本。

1. 在 Zotero 中运行 `ZoteroPMID_Zotero.js` 映射脚本（见下方说明）
2. 回到 Word，**先保存文档备份**
3. `Alt+F8` → 选择 `AutoInsertAll` → 运行
4. 点击 Zotero 工具栏 → **Refresh**（首次选择 Vancouver 样式）

**Zotero JS 映射脚本运行方法：**
1. Zotero → 工具 > 开发者 > Run JavaScript
2. 粘贴 `ZoteroPMID_Zotero.js` 全部内容 → Run
3. 确认输出 `DONE` 且匹配数正确

---

### 最后：手动处理非 PMID 引用

以下格式的引用不会被自动处理，需手动通过 Zotero 工具栏的 **Add/Edit Citation** 逐个插入：

- DOI 格式：`[10.17863/CAM.48429]`
- URL 格式：`[https://doi.org/...]`、`[https://cran.r-project.org/...]` 等

---

## 可配置参数

### VBA 宏 (`ZoteroPMID_Macro.bas`)

```vba
Private Const WORK_FOLDER   As String = "ZoteroPMID"      ' 工作文件夹名
Private Const FILE_MAPPING  As String = "pmid_mapping.txt" ' 映射文件名
Private Const FILE_NBIB     As String = "citations.nbib"   ' MEDLINE 文件名
Private Const FILE_PMIDS    As String = "pmid_list.txt"    ' PMID 列表文件名

' 输出根目录："Downloads" = 下载文件夹, "Desktop" = 桌面
Private Const OUTPUT_BASE   As String = "Downloads"

' 替换时是否删除 [PMID] 前的空格（Vancouver 等上标样式设为 True）
Private Const REMOVE_SPACE_BEFORE_CITE As Boolean = True

' PMID 正则匹配规则（当前：7-9位纯数字）
Private Const RE_BRACKET    As String = "\[(\d{7,9}(?:,\s*\d{7,9})*)\]"
Private Const RE_SINGLE     As String = "\d{7,9}"
```

> - 如果 PMID 位数不在 7-9 之间，修改正则中的 `{7,9}` 即可（如 `{6,9}` 支持 6 位 PMID）。
> - `OUTPUT_BASE` 控制中间文件的输出位置，VBA 和 JS 两边需保持一致。
> - `REMOVE_SPACE_BEFORE_CITE` 设为 `True` 时，Step3 会自动删除 `[PMID]` 前的一个空格，适用于 Vancouver 等上标引用样式；设为 `False` 保留空格，适用于方括号编号样式。

### Zotero JS (`ZoteroPMID_Zotero.js`)

```javascript
var WORK_FOLDER     = "ZoteroPMID";           // 工作文件夹名（需与 VBA 一致）
var MAPPING_FILE    = "pmid_mapping.txt";      // 映射文件名（需与 VBA 一致）
var PMID_LIST_FILE  = "pmid_list.txt";         // PMID 列表文件名（需与 VBA 一致）
var NBIB_FILE       = "citations.nbib";        // MEDLINE 文件名（需与 VBA 一致）
var COLLECTION_PATH = ["Research", "Citation"];     // Zotero 集合路径层级（我的文库 > Research > Citation）
var OUTPUT_BASE     = "Downloads";             // 输出根目录（需与 VBA 一致）
```

> - `COLLECTION_PATH` 按层级填写集合名称。如文献在 `我的文库 > MyProject` 下，改为 `["MyProject"]`。
> - `OUTPUT_BASE` 必须与 VBA 宏中的值一致，否则两边找不到对方的文件。

---

## 常见问题

| 问题 | 原因 | 解决方案 |
|------|------|----------|
| VBA 报 `Attribute VB_Name` 语法错误 | 手动粘贴代码时不需要此行 | 删掉第一行 `Attribute VB_Name = "ZoteroPMID"` |
| 引用上标前有多余空格 | 原文 `[PMID]` 前有空格被保留 | 将 `REMOVE_SPACE_BEFORE_CITE` 设为 `True` |
| Step3 后显示 JSON 原始文本 | 域的 ShowCodes 属性为 True | 宏已自动修复；若仍有问题，在立即窗口运行 `For Each f In ActiveDocument.Fields: f.ShowCodes = False: Next` |
| 部分 PMID 未被替换 | PMID 不在映射文件中 | 检查该 PMID 是否成功导入 Zotero 且有 DOI |

---

## 技术原理

**方案 A（混合方案）：**
```
文档 [PMID]  ──VBA Step1──>  pmid_list.txt + citations.nbib
                                    │
                              手动导入 Zotero
                                    │
                 VBA Step3_InsertNextCitation (逐个):
                   找到 [PMID] → 复制PMID → 删除占位符
                        → 调用 ZoteroAddEditCitation
                        → 用户在Zotero对话框中选择条目
                        → Zotero 原生创建引用域（完整集成）
                   重复直到全部完成
                        → 手动 Add/Edit Bibliography
```

**方案 B（全自动方案）：**
```
文档 [PMID]  ──VBA Step1──>  pmid_list.txt + citations.nbib
                                    │
                              手动导入 Zotero
                                    │
                 Zotero JS: PMID ──MEDLINE──> DOI ──Zotero──> 条目URI
                                    │
                              pmid_mapping.txt
                                    │
                 VBA Step3_AutoInsertAll: [PMID] → ADDIN ZOTERO_ITEM 域代码
                                    │
                 Zotero Refresh: 域代码 → 格式化引用 [1] + 参考文献列表
```
