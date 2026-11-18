# wps-table-unwrap

> 把 WPS 表格里一个单元格的多行文字，自动拆成「一行一格」。

一个 PowerShell 脚本，通过 COM 接口连上正在运行的 **WPS 文字（WPS Writer）**，把表格中
"一格多行"的内容拆开：第 1 行留在原格，其余各行**下移到紧接其下新插入的行**；同一行里有
多列同时溢出时，按最大溢出数一次性插够行数，**各列保持对齐**。

![处理前](docs/before.png)

![处理后](docs/after.png)

## 它能解决什么

从网页、PDF、聊天记录里复制来的表格，经常出现"一个单元格里塞了好几行"的情况——按回车
换行的、列宽不够自动折行的都有。手工一行行剪贴很费时，这个脚本几秒钟就能全部拆开。

## 环境要求

- Windows + WPS Office（已安装在默认位置即可）
- 无需安装 Python、无需安装 Office；只需要系统自带的 PowerShell

## 快速开始

### 方式一：双击 / 拖放（推荐）

双击 `run.cmd`，处理 **WPS 当前打开的文档**，结果另存为 `<原文件名>_单行.docx`。

或者把 `.docx` 文件**拖到 `run.cmd` 上**，处理那个文件。

### 方式二：命令行

```powershell
# 处理当前打开的文档（另存副本，原文件不动）
.\wps-table-unwrap.ps1

# 先看看会拆多少行，不改文档
.\wps-table-unwrap.ps1 -Path "D:\报告\实验报告.docx" -DryRun

# 只处理第 1 张表、只处理第 2 列，其它列的内容复制到新行
.\wps-table-unwrap.ps1 -Path .\a.docx -Tables 1 -Columns 2 -FillOther Copy

# 就地修改（自动留 .bak 备份）
.\wps-table-unwrap.ps1 -Path .\a.docx -InPlace
```

> 首次运行如果被策略拦住，在 PowerShell 里执行一次：
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

## 参数

| 参数 | 说明 |
| --- | --- |
| `-Path` | 要处理的 `.docx/.doc`。省略则操作 WPS 当前激活的文档 |
| `-OutPath` | 输出路径。默认 `<原文件名>_单行<扩展名>`，与原文件同目录 |
| `-Tables` | 只处理指定表格序号（1 起），如 `-Tables 1,3` |
| `-Columns` | 只处理指定列（1 起），如 `-Columns 2` |
| `-Mode` | `Auto`（默认）：手动换行 + 自动折行都拆；`Hard`：只拆手动换行/分段；`Wrap`：只拆自动折行 |
| `-FillOther` | `None`（默认）：新行的其它列留空；`Copy`：复制原行内容（适合"第一列是标签"） |
| `-InPlace` | 就地修改原文件（先备份 `.bak`） |
| `-DryRun` | 只检查并报告，不改动文档 |
| `-MaxSplit` | 单格最多拆成几行，默认 50，防止异常文档行数爆炸 |

## 工作原理

1. **连接**：优先 `GetActiveObject('Word.Application')` 连到已运行的 WPS；没有运行实例时
   新建一个并打开 `-Path`。WPS 注册的 `Word.Application` 兼容 Word 对象模型。
2. **判定"多行"**：分两种情况处理——
   - 手动换行（`Shift+Enter`，`\v`）与段落标记（`Enter`，`\r`）：直接按字符切分；
   - 列宽不够的自动折行：把文字写回单元格，用 `Range.Information(10)`（字符所在显示行号）
     配合二分法量出真实折行位置，再切分。
3. **落位**：从最后一行往前扫，插入新行不会打乱尚未处理的索引；写入完成后复核，仍为多行的
   单元格会提示人工确认。

### 两个实现上的坑

- WPS 的 `Range.ComputeStatistics(wdStatisticLines)` 在某些列上会稳定返回 `0`，直接拿来判断
  会漏判。实测改用 `Range.Information(10)` 取首末字符行号之差才准确，取不到时再退回
  `ComputeStatistics`。
- 测量行数时要把文字临时写进单元格，函数退出前必须还原，否则 `-DryRun` 会破坏原文。

## 安全性

- 默认**另存为新文件**，原文件不动；只有加 `-InPlace` 才覆盖，且先备份 `.bak`。
- 只关闭脚本自己打开的那个文档，**绝不关闭你已打开的文件**。

## 已知限制

- 单元格内部的局部格式（加粗、变色、字号混排）在文字迁移时会丢失，只保留段落格式。
- 横向合并的单元格会跳过，并在复核阶段提示。
- 需要 WPS/Word 的桌面版；WPS 网页版无法通过 COM 操作。

## 致谢与许可

MIT License，见 [LICENSE](LICENSE)。

---

### English

`wps-table-unwrap` is a small PowerShell tool that connects to a running **WPS Office Writer**
(via the `Word.Application` COM interface) and splits multi-line table cells into one line per
row — the first line stays, the remaining ones are moved into newly inserted rows, columns stay
aligned. Handy for tables pasted from web pages or PDFs where several values ended up crammed into
a single cell. No Python, no Office required — just Windows + WPS Office.
