<#
.SYNOPSIS
    连接 WPS 文字（WPS Writer），把表格中"一个单元格里的多行文字"自动拆成"一行一格"。

.DESCRIPTION
    工作方式：
      1. 若本机已有正在运行的 WPS，则直接连上去；没有则用 -Path 启动一个新的实例。
      2. 逐表、逐行、逐单元格检查文字行数。一个单元格里出现多行时：
           · 第 1 行留在原单元格；
           · 第 2 行及以后的文字"下传"到紧接其下新插入的行（同一列）；
           · 一行里有多列同时溢出时，按最大溢出数一次性插入足够行数，各列对齐。
      3. 默认另存为新文件（原文件不动）；-InPlace 就地修改（自动留 .bak 备份）。

    行数判定：WPS 的 ComputeStatistics 在某些列上会错误地返回 0，
    因此改用 Range.Information(wdFirstCharacterLineNumber) 取首末字符的行号之差，
    实测稳定；取不到时再退回 ComputeStatistics。

    能识别的两种"多行"：
       · 手动换行（Shift+Enter，\v）与段落标记（Enter，\r）—— 直接按字符切分；
       · 列宽不够造成的自动折行 —— 写回单元格后用二分法找出真实折行位置再切分。

.PARAMETER Path
    要处理的 .docx/.doc 路径。省略则操作 WPS 当前激活的文档。

.PARAMETER OutPath
    输出路径。省略时默认为 "<原文件名>_单行<扩展名>"，放在原文件同一目录。

.PARAMETER Tables
    只处理指定序号的表格（1 起），如 -Tables 1,3。省略则处理全部表格。

.PARAMETER Columns
    只处理指定列（1 起），如 -Columns 2。省略则处理全部列。

.PARAMETER Mode
    Auto（默认）= 先按手动换行/段落切分，再处理仍会折行的文字；
    Hard        = 只按手动换行/段落切分，不拆自动折行；
    Wrap        = 只按自动折行切分（Enter 产生的段落先合并成空格）。

.PARAMETER FillOther
    None（默认）= 新插入的行里，其它列留空；
    Copy        = 新插入的行里，其它列复制原行的内容（适合"第一列是标签"的表）。

.PARAMETER InPlace
    就地修改原文件（先复制一份 .bak 备份）。

.PARAMETER DryRun
    只检查并报告，不改动文档。

.PARAMETER MaxSplit
    单个单元格最多拆成几行，防止异常文档行数爆炸。默认 50。

.EXAMPLE
    .\WPS-表格单行化.ps1 -Path "D:\报告\实验报告.docx" -DryRun

.EXAMPLE
    .\WPS-表格单行化.ps1                       # 处理 WPS 当前打开的文档，另存副本

.EXAMPLE
    .\WPS-表格单行化.ps1 -Path .\a.docx -Tables 1 -Columns 2 -FillOther Copy
#>
[CmdletBinding()]
param(
    [string]$Path,
    [string]$OutPath,
    [int[]]$Tables,
    [int[]]$Columns,
    [ValidateSet('Auto', 'Hard', 'Wrap')][string]$Mode = 'Auto',
    [ValidateSet('None', 'Copy')][string]$FillOther = 'None',
    [switch]$InPlace,
    [switch]$DryRun,
    [int]$MaxSplit = 50
)

$ErrorActionPreference = 'Stop'

# ---------- WPS/Word 对象模型常量 ----------
$wdStatisticLines   = 1      # ComputeStatistics 的"行数"统计项
$wdFirstCharLine    = 10     # Information：该字符所在的显示行号
$wdDoNotSaveChanges = 0

function Write-Step([string]$msg, [string]$color = 'Gray') {
    Write-Host $msg -ForegroundColor $color
}

# ---------- 行数判定 ----------
# 取一段 range 首末字符的行号差；失败时退回 ComputeStatistics
function Get-RangeLineCount($doc, $range) {
    $s = [int]$range.Start; $e = [int]$range.End
    if (($e - $s) -le 1) { return 1 }          # 空单元格
    $a = -1; $b = -1
    try { $a = [int]$doc.Range($s, $s + 1).Information($wdFirstCharLine) } catch { }
    try { $b = [int]$doc.Range($e - 1, $e).Information($wdFirstCharLine) } catch { }
    if ($a -gt 0 -and $b -ge $a) { return ($b - $a + 1) }
    try { return [math]::Max(1, [int]$range.ComputeStatistics($wdStatisticLines)) } catch { return 1 }
}

function Get-CellLineCount($doc, $cell) {
    return (Get-RangeLineCount $doc $cell.Range)
}

# 单元格文字（去掉末尾的段落标记 + 单元格结束符）
function Get-CellText($cell) {
    $t = [string]$cell.Range.Text
    return ($t -replace "[\r\a]+$", '')
}

# 按"实际显示的行"把一段文字切开（依赖当前单元格的列宽与字体）
function Split-ByDisplayWidth($doc, $cell, [string]$text) {
    $n = $text.Length
    if ($n -le 0) { return @() }
    $start = [int]$cell.Range.Start

    function Local-Lines([int]$from, [int]$len) {
        if ($len -le 0) { return 0 }
        $a = $start + $from
        $a1 = -1; $b1 = -1
        try { $a1 = [int]$doc.Range($a, $a + 1).Information($wdFirstCharLine) } catch { }
        try { $b1 = [int]$doc.Range($a + $len - 1, $a + $len).Information($wdFirstCharLine) } catch { }
        if ($a1 -gt 0 -and $b1 -ge $a1) { return ($b1 - $a1 + 1) }
        try { return [int]$doc.Range($a, $a + $len).ComputeStatistics($wdStatisticLines) } catch { return 1 }
    }

    $segments = New-Object System.Collections.Generic.List[string]
    $pos = 0
    $guard = 0
    while ($pos -lt $n -and $guard -lt $MaxSplit) {
        $guard++
        $rest = $n - $pos
        if ((Local-Lines $pos $rest) -le 1) {
            $segments.Add($text.Substring($pos, $rest).TrimEnd())
            break
        }
        $lo = 1; $hi = $rest; $best = 1
        while ($lo -le $hi) {
            $mid = [int](($lo + $hi) / 2)
            if ((Local-Lines $pos $mid) -le 1) { $best = $mid; $lo = $mid + 1 }
            else { $hi = $mid - 1 }
        }
        $segments.Add($text.Substring($pos, $best).TrimEnd())
        $pos += $best
    }
    return $segments.ToArray()
}

# 把一个单元格的文字拆成"每行一段"
function Split-CellContent($doc, $cell, [string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return @('') }
    $orig = Get-CellText $cell        # 记住原内容；测量时会把文字临时写进单元格
    $result = New-Object System.Collections.Generic.List[string]

    # ① 手动换行 / 段落标记
    if ($Mode -ne 'Wrap') {
        $pieces = @($text -split "[\r\v]" | Where-Object { $_.Trim() -ne '' })
        if ($pieces.Count -eq 0) { $pieces = @($text) }
    }
    else {
        $pieces = @($text -replace "[\r\v]", ' ')
    }

    # ② 仍会折行的，按显示宽度继续切
    foreach ($piece in $pieces) {
        $p = ([string]$piece).Trim()
        if ($p -eq '') { continue }
        if ($Mode -eq 'Hard') { $result.Add($p); continue }

        $cell.Range.Text = $p        # 借用单元格本身来测量
        if ((Get-CellLineCount $doc $cell) -le 1) {
            $result.Add($p)
        }
        else {
            $sub = Split-ByDisplayWidth $doc $cell $p
            if (-not $sub -or $sub.Count -eq 0) { $result.Add($p) }
            else { foreach ($s in $sub) { $result.Add($s) } }
        }
    }

    $cell.Range.Text = $orig          # 还原，使本函数无副作用（DryRun 才不会破坏原文）
    if ($result.Count -eq 0) { $result.Add('') }
    return $result.ToArray()
}

# ---------- 保存 ----------
function Save-DocumentAs($doc, [string]$target) {
    $ext = [System.IO.Path]::GetExtension($target).ToLower()
    $fmt = if ($ext -eq '.doc') { 0 } else { 16 }   # 0=Word97-2003  16=docx
    try { $doc.SaveAs2($target, $fmt); return } catch { }
    try { $doc.SaveAs([ref]$target, [ref]$fmt); return } catch { }
    $doc.SaveAs($target)
}

# ---------- 连接 WPS ----------
$ownApp = $false
$openedHere = $false      # 只关闭本脚本自己打开的文档，绝不动用户已打开的文件
$app = $null
$doc = $null

try {
    if (-not $Path) {
        try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
        catch { throw "没有检测到正在运行的 WPS。请先打开文档，或用 -Path 指定文件路径。" }
        Write-Step "已连接到正在运行的 WPS。" 'Green'
    }
    else {
        if (-not (Test-Path -LiteralPath $Path)) { throw "文件不存在：$Path" }
        try {
            $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
            Write-Step "已连接到正在运行的 WPS。" 'Green'
        }
        catch {
            Write-Step "启动 WPS…" 'Gray'
            $app = New-Object -ComObject Word.Application
            $app.Visible = $false
            $ownApp = $true
        }
        $Path = (Resolve-Path -LiteralPath $Path).Path
        $doc = $app.Documents | Where-Object { $_.FullName -eq $Path } | Select-Object -First 1
        if (-not $doc) { $doc = $app.Documents.Open($Path); $openedHere = $true }
    }

    if (-not $doc) { $doc = $app.ActiveDocument }
    if (-not $doc) { throw "WPS 中没有打开的文档。" }

    $app.DisplayAlerts = 0
    $srcName = $doc.Name
    $srcFull = ''
    try { $srcFull = $doc.FullName } catch { }
    Write-Step "当前文档：$srcName" 'Cyan'

    if ($doc.Tables.Count -eq 0) { throw "当前文档里没有表格。" }

    $doc.Repaginate()          # 强制重新排版，保证行号准确

    $tableIdx = if ($Tables) { $Tables } else { 1..$doc.Tables.Count }
    Write-Step ("共 {0} 个表格，本次处理：{1}" -f $doc.Tables.Count, ($tableIdx -join ', '))

    $statRows = 0; $statCells = 0; $statNewRows = 0

    foreach ($ti in $tableIdx) {
        if ($ti -lt 1 -or $ti -gt $doc.Tables.Count) { Write-Step "跳过：第 $ti 个表格不存在" 'Yellow'; continue }
        $tbl = $doc.Tables.Item($ti)
        $rowCount = $tbl.Rows.Count
        try { $colCount = $tbl.Columns.Count } catch { $colCount = 0 }
        Write-Step ("`n[表 {0}] {1} 行 × {2} 列" -f $ti, $rowCount, $colCount) 'Cyan'

        # 从最后一行往前处理，插入新行不会打乱尚未处理的索引
        for ($r = $rowCount; $r -ge 1; $r--) {
            try { $rowObj = $tbl.Rows.Item($r) } catch { continue }

            $plan = @{}          # 列号 -> 该列拆出来的各行文字
            $originals = @{}     # 列号 -> 原内容（FillOther = Copy 时用）
            $maxExtra = 0

            if ($FillOther -eq 'Copy') {
                foreach ($cell in $rowObj.Cells) {
                    try { $originals[[int]$cell.ColumnIndex] = (Get-CellText $cell) } catch { }
                }
            }

            foreach ($cell in $rowObj.Cells) {
                $c = $null
                try { $c = [int]$cell.ColumnIndex } catch { continue }
                if ($Columns -and ($Columns -notcontains $c)) { continue }

                $text = Get-CellText $cell
                if ((Get-CellLineCount $doc $cell) -le 1) { continue }

                $segs = @(Split-CellContent $doc $cell $text)
                if ($segs.Count -gt $MaxSplit) { $segs = $segs[0..($MaxSplit - 1)] }
                if ($segs.Count -le 1) { continue }

                $cell.Range.Text = $segs[0]
                $plan[$c] = $segs
                if (($segs.Count - 1) -gt $maxExtra) { $maxExtra = $segs.Count - 1 }
                $statCells++
            }

            if ($plan.Count -eq 0) { continue }
            $doc.Repaginate()

            if ($DryRun) {
                Write-Step ("  第 {0} 行：将拆出 {1} 行，涉及列 {2}" -f $r, $maxExtra, (($plan.Keys | Sort-Object) -join ',')) 'Yellow'
                $statRows++; $statNewRows += $maxExtra
                continue
            }

            # 一次插入 $maxExtra 行（紧跟第 $r 行之后）
            for ($k = 1; $k -le $maxExtra; $k++) {
                if ($r + $k -le $tbl.Rows.Count) { [void]$tbl.Rows.Add($tbl.Rows.Item($r + 1)) }
                else { [void]$tbl.Rows.Add() }
            }
            $statNewRows += $maxExtra

            foreach ($c in ($plan.Keys | Sort-Object)) {
                $segs = $plan[$c]
                for ($j = 1; $j -lt $segs.Count; $j++) {
                    try { $tbl.Cell($r + $j, $c).Range.Text = $segs[$j] }
                    catch { Write-Step ("    警告：第 {0} 行第 {1} 列写入失败" -f ($r + $j), $c) 'Yellow' }
                }
                if ($FillOther -eq 'Copy') {
                    for ($j = 1; $j -le $maxExtra; $j++) {
                        if ($j -lt $segs.Count) { continue }
                        foreach ($oc in $originals.Keys) {
                            if ($oc -eq $c) { continue }
                            try { $tbl.Cell($r + $j, $oc).Range.Text = $originals[$oc] } catch { }
                        }
                    }
                }
            }

            Write-Step ("  第 {0} 行已拆出 {1} 行" -f $r, $maxExtra) 'DarkGray'
            $statRows++
        }
    }

    $doc.Repaginate()

    Write-Step ("`n==== 统计 ====") 'Cyan'
    Write-Step ("被拆分的行：{0}；被拆分的单元格：{1}；新增行：{2}" -f $statRows, $statCells, $statNewRows)

    if ($DryRun) {
        Write-Step "`n[DryRun] 未修改文档。" 'Yellow'
    }
    else {
        if ($InPlace) {
            if ($srcFull) {
                $bak = "$srcFull.bak"
                Copy-Item -LiteralPath $srcFull -Destination $bak -Force
                Write-Step "已备份原文件：$bak" 'Gray'
            }
            $doc.Save()
            Write-Step "已就地保存：$srcFull" 'Green'
        }
        else {
            if (-not $OutPath) {
                $dir = if ($srcFull) { Split-Path -Parent $srcFull } else { (Get-Location).Path }
                $base = [System.IO.Path]::GetFileNameWithoutExtension($srcName)
                $ext = [System.IO.Path]::GetExtension($srcName); if (-not $ext) { $ext = '.docx' }
                $OutPath = Join-Path $dir ("{0}_单行{1}" -f $base, $ext)
            }
            $OutPath = [System.IO.Path]::GetFullPath($OutPath)
            $outDir = Split-Path -Parent $OutPath
            if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
            Save-DocumentAs $doc $OutPath
            Write-Step "已另存为：$OutPath" 'Green'
        }

        # 复核
        $left = 0
        foreach ($ti in $tableIdx) {
            if ($ti -lt 1 -or $ti -gt $doc.Tables.Count) { continue }
            $tbl = $doc.Tables.Item($ti)
            for ($r = 1; $r -le $tbl.Rows.Count; $r++) {
                try { $cells = $tbl.Rows.Item($r).Cells } catch { continue }
                foreach ($cell in $cells) { if ((Get-CellLineCount $doc $cell) -gt 1) { $left++ } }
            }
        }
        if ($left -eq 0) { Write-Step "复核通过：目标表格中已没有多行单元格。" 'Green' }
        else { Write-Step ("复核：仍有 {0} 个单元格为多行（可能受合并单元格或列宽限制），请人工确认。" -f $left) 'Yellow' }
    }
}
catch {
    Write-Step ("`n[出错] " + $_.Exception.Message) 'Red'
    exit 1
}
finally {
    if ($doc -and $openedHere -and -not $InPlace) {
        try { $doc.Close($wdDoNotSaveChanges) } catch { }
    }
    if ($ownApp -and $app) { try { $app.Quit() } catch { } }
    if ($app) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { } }
    [GC]::Collect()
}
