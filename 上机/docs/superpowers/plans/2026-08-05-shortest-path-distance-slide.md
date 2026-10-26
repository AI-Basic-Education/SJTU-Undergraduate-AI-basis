# Shortest-Path Distance Slide Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace slide 10's placeholder with a polished shortest-path metric example whose two key formulas are native editable PowerPoint equations.

**Architecture:** Use PowerPoint's Windows COM API so the existing master, logo, title area, footer, and theme remain intact. Copy the source deck to a new filename, replace only slide 10's body shapes, create the graph and explanatory text with native PowerPoint shapes, and convert linear UnicodeMath text into built-up equation math zones.

**Tech Stack:** PowerShell 7/Windows PowerShell, Microsoft PowerPoint COM, Office Math, bundled PPTX OOXML validation scripts.

---

### Task 1: Create the slide-editing script

**Files:**
- Create: `.ppt-edit/apply-shortest-path-example.ps1`
- Read: `ppt/第9-10节-传统无监督学习.pptx`
- Create: `ppt/第9-10节-传统无监督学习_增加例子版.pptx`

- [ ] **Step 1: Create a PowerShell script with reusable shape helpers**

The script will:

1. Copy the source presentation without overwriting it.
2. Open the copy in PowerPoint.
3. Keep the slide 10 title and slide-number placeholder but remove its temporary body placeholder.
4. Add the red rounded frame, two-column content, weighted graph, proof, and application line.
5. Add both key expressions through `TextRange2.MathZones.Add(...).BuildUp()` so they become native Office Math.

Use these exact helper interfaces:

```powershell
param(
  [string]$Source = 'C:\nut\research\mycourse\全校人工智能大课\pptv10\ppt\第9-10节-传统无监督学习.pptx',
  [string]$Output = 'C:\nut\research\mycourse\全校人工智能大课\pptv10\ppt\第9-10节-传统无监督学习_增加例子版.pptx'
)

$red = 0x1616B6
$dark = 0x2D2C24
$blue = 0xAF6937

function Add-Text($slide, $text, $left, $top, $width, $height, $size = 18, $color = $dark, $bold = $false) {
  $shape = $slide.Shapes.AddTextBox(1, $left, $top, $width, $height)
  $shape.TextFrame2.MarginLeft = 2
  $shape.TextFrame2.MarginRight = 2
  $shape.TextFrame2.MarginTop = 1
  $shape.TextFrame2.MarginBottom = 1
  $shape.TextFrame2.WordWrap = -1
  $shape.TextFrame2.TextRange.Text = $text
  $shape.TextFrame2.TextRange.Font.NameFarEast = 'Microsoft YaHei'
  $shape.TextFrame2.TextRange.Font.Name = 'Book Antiqua'
  $shape.TextFrame2.TextRange.Font.Size = $size
  $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $color
  $shape.TextFrame2.TextRange.Font.Bold = [int]$bold * -1
  return $shape
}

function Add-Equation($slide, $linear, $left, $top, $width, $height, $size = 20) {
  $shape = $slide.Shapes.AddTextBox(1, $left, $top, $width, $height)
  $shape.TextFrame2.MarginLeft = 0
  $shape.TextFrame2.MarginRight = 0
  $shape.TextFrame2.MarginTop = 0
  $shape.TextFrame2.MarginBottom = 0
  $shape.TextFrame2.TextRange.Text = $linear
  $shape.TextFrame2.TextRange.Font.Name = 'Cambria Math'
  $shape.TextFrame2.TextRange.Font.Size = $size
  $shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
  $range = $shape.TextFrame2.TextRange
  $zone = $range.MathZones.Add($range)
  $zone.BuildUp()
  return $shape
}

function Add-Node($slide, $label, $left, $top) {
  $node = $slide.Shapes.AddShape(9, $left, $top, 34, 34)
  $node.Fill.ForeColor.RGB = $blue
  $node.Line.ForeColor.RGB = $blue
  $node.TextFrame2.TextRange.Text = $label
  $node.TextFrame2.TextRange.Font.Name = 'Cambria Math'
  $node.TextFrame2.TextRange.Font.Size = 16
  $node.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
  $node.TextFrame2.TextRange.ParagraphFormat.Alignment = 2
  $node.TextFrame2.VerticalAnchor = 3
  return $node
}

function Add-Edge($slide, $x1, $y1, $x2, $y2, $weight, $labelX, $labelY, $accent = $false) {
  $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
  $line.Line.ForeColor.RGB = $(if ($accent) { $red } else { 0x777777 })
  $line.Line.Weight = $(if ($accent) { 3 } else { 1.5 })
  Add-Text $slide $weight $labelX $labelY 30 18 12 $dark $false | Out-Null
}

Copy-Item -LiteralPath $Source -Destination $Output -Force
$app = New-Object -ComObject PowerPoint.Application
$presentation = $app.Presentations.Open($Output, $false, $false, $false)

try {
  $slide = $presentation.Slides.Item(10)

  foreach ($shape in @($slide.Shapes)) {
    if ($shape.Name -like '内容占位符*') { $shape.Delete() }
  }

  $title = $slide.Shapes.Title
  $title.TextFrame2.TextRange.Text = '例：图上的最短路径距离'

  $frame = $slide.Shapes.AddShape(5, 34, 76, 892, 438)
  $frame.Fill.Visible = 0
  $frame.Line.ForeColor.RGB = $red
  $frame.Line.Weight = 2.25

  Add-Text $slide '定义与图示' 62 96 360 28 20 $red $true | Out-Null
  Add-Text $slide '给定连通无向图 G=(V,E)，每条边权 w(e)>0。' 62 132 390 30 16 $dark $false | Out-Null
  Add-Equation $slide 'd_G(u,v)=min_(P:u→v)∑_(e∈P)w(e)' 72 168 390 44 21 | Out-Null
  Add-Text $slide '即所有从 u 到 v 的路径中，长度最小者。' 62 215 390 28 15 $dark $false | Out-Null

  Add-Edge $slide 112 300 218 268 '2' 157 274 $true
  Add-Edge $slide 218 268 344 300 '3' 275 273 $true
  Add-Edge $slide 112 300 220 382 '6' 153 338 $false
  Add-Edge $slide 220 382 344 300 '2' 277 338 $false
  Add-Edge $slide 218 268 220 382 '4' 223 316 $false
  Add-Node $slide 'u' 95 283 | Out-Null
  Add-Node $slide 'a' 201 251 | Out-Null
  Add-Node $slide 'b' 203 365 | Out-Null
  Add-Node $slide 'v' 327 283 | Out-Null
  Add-Text $slide '最短路径 u→a→v，距离为 2+3=5' 85 420 330 30 15 $red $true | Out-Null

  $divider = $slide.Shapes.AddLine(485, 112, 485, 475)
  $divider.Line.ForeColor.RGB = 0xD8D8D8
  $divider.Line.Weight = 1

  Add-Text $slide '为什么它是距离？' 520 96 350 28 20 $red $true | Out-Null
  Add-Text $slide '✓ 非负性：边权为正，路径长度非负' 520 140 360 30 16 $dark $false | Out-Null
  Add-Text $slide '✓ 同一性：d_G(u,v)=0 当且仅当 u=v' 520 181 360 30 16 $dark $false | Out-Null
  Add-Text $slide '✓ 对称性：无向路径可反向行走' 520 222 360 30 16 $dark $false | Out-Null
  Add-Text $slide '✓ 三角不等式：经 z 的两段最短路径可拼接' 520 263 370 34 16 $dark $false | Out-Null
  Add-Equation $slide 'd_G(u,v)≤d_G(u,z)+d_G(z,v)' 522 307 350 44 20 | Out-Null
  Add-Text $slide '应用：地图导航 · 网络路由 · ISOMAP 测地距离' 520 390 360 52 17 $blue $true | Out-Null

  $presentation.Save()
}
finally {
  $presentation.Close()
  $app.Quit()
  [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
  [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
```

- [ ] **Step 2: Run the script**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File ".ppt-edit\apply-shortest-path-example.ps1"
```

Expected: `ppt/第9-10节-传统无监督学习_增加例子版.pptx` exists and PowerPoint exits cleanly.

### Task 2: Validate structure and native equations

**Files:**
- Read: `ppt/第9-10节-传统无监督学习_增加例子版.pptx`
- Create: `.ppt-edit/final-unpacked/`

- [ ] **Step 1: Unpack and validate the output**

Run:

```powershell
python "C:\Users\xuzhi\.claude\skills\pptx\ooxml\scripts\unpack.py" `
  "C:\nut\research\mycourse\全校人工智能大课\pptv10\ppt\第9-10节-传统无监督学习_增加例子版.pptx" `
  ".ppt-edit\final-unpacked"

python "C:\Users\xuzhi\.claude\skills\pptx\ooxml\scripts\validate.py" `
  ".ppt-edit\final-unpacked" `
  --original "C:\nut\research\mycourse\全校人工智能大课\pptv10\ppt\第9-10节-传统无监督学习.pptx"
```

Expected: validation exits with code 0 and reports no new OOXML errors.

- [ ] **Step 2: Confirm native Office Math exists on slide 10**

Search `final-unpacked/ppt/slides/slide10.xml` for Office Math markup such as `m:oMath` or `a14:m`.

Expected: at least two math objects corresponding to the distance definition and triangle inequality.

### Task 3: Render and visually inspect slide 10

**Files:**
- Create: `.ppt-edit/final-rendered/slide10.png`

- [ ] **Step 1: Export slide 10 through PowerPoint**

Use PowerPoint COM to export at 1600×900:

```powershell
$ppt = 'C:\nut\research\mycourse\全校人工智能大课\pptv10\ppt\第9-10节-传统无监督学习_增加例子版.pptx'
$out = (Resolve-Path '.ppt-edit').Path + '\final-rendered'
New-Item -ItemType Directory -Force -Path $out | Out-Null
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($ppt, $true, $true, $false)
$pres.Slides.Item(10).Export((Join-Path $out 'slide10.png'), 'PNG', 1600, 900)
$pres.Close()
$app.Quit()
```

- [ ] **Step 2: Inspect and correct layout defects**

Check:

- title, logo, frame, and footer match adjacent slides;
- both equations are built up and legible;
- no text overlaps the graph, divider, frame, title, or footer;
- graph edge labels do not collide with nodes;
- proof lines and application text fit inside the right column;
- page number remains `10/40`.

If any check fails, adjust the coordinates or font sizes in `.ppt-edit/apply-shortest-path-example.ps1`, rerun Tasks 1–3, and retain only the corrected output.
