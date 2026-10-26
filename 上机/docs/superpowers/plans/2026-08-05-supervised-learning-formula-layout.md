# Supervised Learning Formula Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produce a new copy of the supervised-learning deck whose slides 10, 11, 15, and 16 use editable native PowerPoint equations and the approved teaching-hierarchy layout.

**Architecture:** A PowerShell automation script drives the locally installed PowerPoint application through COM. It opens a copy of the source deck, removes only the editable content shapes on the four target slides, rebuilds those shapes with existing theme colors, and inserts equations through PowerPoint's own equation commands. A verification script checks slide count, native Office Math XML, and PowerPoint rendering.

**Tech Stack:** PowerShell, Microsoft PowerPoint COM, Office Math, Python OOXML validation utilities.

---

## File map

- Create `tools/format_supervised_learning_ppt.ps1`: deterministic slide reconstruction and native-equation insertion.
- Create `tools/verify_supervised_learning_ppt.ps1`: structural checks and target-slide image export.
- Create `ppt/第7-8节-传统监督学习_公式排版版.pptx`: final presentation; source deck remains unchanged.

### Task 1: Build reusable PowerPoint helpers

**Files:**
- Create: `tools/format_supervised_learning_ppt.ps1`

- [ ] **Step 1: Define source/output paths and visual constants**

Use the source deck under `ppt`, write to the suffixed output filename, and define the existing palette: title red `#C3151B`, section navy `#003C68`, step green `#08751B`, class blue `#0066C9`, equation purple `#812080`, body black, and explanation gray.

- [ ] **Step 2: Implement regular text creation**

Create `Add-Text($slide, $text, $left, $top, $width, $height, $size, $color, $bold, $font, $align)` using `Shapes.AddTextbox`, `TextFrame2`, fixed internal margins, and no automatic bulleting.

- [ ] **Step 3: Implement native equation creation**

Create `Add-Equation($app, $slide, $linearText, $left, $top, $width, $height, $size, $color, $align)`:

```powershell
$slide.Select()
$app.CommandBars.ExecuteMso("InsertBuildingBlocksEquationsGallery")
$shape = $app.ActiveWindow.Selection.ShapeRange.Item(1)
$shape.TextFrame.TextRange.Text = $linearText
$app.CommandBars.ExecuteMso("EquationProfessional")
$shape.Left = $left
$shape.Top = $top
$shape.Width = $width
$shape.Height = $height
```

Then apply font size, color, alignment, transparent fill, and no outline. This route is required because `TextRange2.MathZones` is read-only and cannot create an equation.

- [ ] **Step 4: Implement safe shape cleanup**

Create `Remove-ContentShapes($slide, $keepNames)` that deletes shapes in reverse z-order while preserving the title placeholder, slide-number placeholder, slide master graphics, and—on slide 15—the existing LDA diagram.

### Task 2: Rebuild slide 10

**Files:**
- Modify: `tools/format_supervised_learning_ppt.ps1`

- [ ] **Step 1: Preserve the original title and page number**

Delete only `TextBox 3`; retain `Title 1` and `Slide Number Placeholder 2`.

- [ ] **Step 2: Add teaching sections**

Add navy `【题目】`, the Chinese experiment description, a centered native equation for the four ordered pairs, the task sentence, navy `【提示】`, and the explanatory sentence.

- [ ] **Step 3: Add the least-squares objective**

Insert a centered native equation equivalent to

```text
min_(a,b) sum_(i=1)^4 (y_i-(a+b x_i))^2
```

with red emphasis and enough height for professional fraction/superscript layout.

### Task 3: Rebuild slide 11

**Files:**
- Modify: `tools/format_supervised_learning_ppt.ps1`

- [ ] **Step 1: Replace the single body textbox**

Delete `TextBox 2`, preserve the title and page number, and add navy `【求解过程】`.

- [ ] **Step 2: Add Step 1**

Add a green step heading and four native equations for the sums. Align their equals signs by using identical left edges and widths.

- [ ] **Step 3: Add Step 2**

Add a green heading and a two-line purple native equation system for the normal equations.

- [ ] **Step 4: Add Step 3**

Add a green heading, three aligned native equations for substitution/elimination, and a red native equation for the final values `a=0.15, b=1.94`.

### Task 4: Rebuild slide 15

**Files:**
- Modify: `tools/format_supervised_learning_ppt.ps1`

- [ ] **Step 1: Preserve structural elements**

Keep `Title 1`, page number, and `Picture 12`; remove all other slide-local text/group shapes.

- [ ] **Step 2: Restore the concept statement**

Place the red LDA concept sentence below the title without entering the right-side logo area.

- [ ] **Step 3: Build the right-side statistics section**

Add two concise Chinese labels and native equations for `(mu_i,Sigma_i)` and `(w^T mu_i,w^T Sigma_i w)`.

- [ ] **Step 4: Build the objective section**

Add the red heading `最大化目标函数`, a centered Fisher objective equation, aligned gray explanations for numerator/denominator, and a native equation for the optimal direction.

- [ ] **Step 5: Restore the reading prompt**

Place the red reading prompt above the footer, within the content-safe area.

### Task 5: Rebuild slide 16

**Files:**
- Modify: `tools/format_supervised_learning_ppt.ps1`

- [ ] **Step 1: Remove overlapping source shapes**

Delete `TextBox 4` and `TextBox 5`; retain the title and page number.

- [ ] **Step 2: Build the left problem column**

Add navy `【题目】`, the sample description, red positive-class samples, blue negative-class samples, the native equation `w=[1,2]^T`, and the three requested quantities.

- [ ] **Step 3: Build the right formula column**

Add a thin red vertical separator, navy `【公式】`, and native equations for projection, mean, and variance:

```text
z=w^T x=w_1 x_1+w_2 x_2
mu=1/n sum_(i=1)^n z_i
sigma^2=1/n sum_(i=1)^n (z_i-mu)^2
```

### Task 6: Save and verify

**Files:**
- Create: `tools/verify_supervised_learning_ppt.ps1`
- Create: `ppt/第7-8节-传统监督学习_公式排版版.pptx`

- [ ] **Step 1: Save to a new filename**

Run the formatter with PowerPoint visible to keep equation commands reliable. Close the presentation and PowerPoint cleanly in `finally`.

- [ ] **Step 2: Validate OOXML**

Unpack the output and run:

```powershell
python "C:\Users\xuzhi\.claude\skills\pptx\ooxml\scripts\validate.py" `
  ".cursor-ppt-work\supervised-learning-unpacked" `
  --original "ppt\第7-8节-传统监督学习.pptx"
```

Expected: no validation errors.

- [ ] **Step 3: Verify native equations**

Search slides 10, 11, 15, and 16 for `<a14:m>` and `<m:oMath>` and require at least one Office Math zone on every target slide.

- [ ] **Step 4: Render target slides**

Open the output through PowerPoint COM and export slides 10, 11, 15, and 16 to `.cursor-ppt-work/final-slide-*.png` at 1600×900.

- [ ] **Step 5: Visually inspect and iterate**

Check title/logo clearance, equation readability, overlapping shapes, content/footer clearance, consistent colors, and spacing. Adjust the formatter and rerun until all four slides pass.
