---
name: python-docx-photo-grid
description: |
  Python-based Word (.docx) 报告生成，配「自适应 2 栏相片网格」技术 + 工程竣工报告
  系列（安装完成 / 设备测试 / 系统验收）+ 勘察备忘录 统一排版（蓝色系、承建商抬头图页眉、
  统一签署区、章节结构）。触发词：生成 docx 报告 with 自动相片排版、相片表格、photo
  auto-layout、竣工报告相片、test report with photos、自适应相片排版、A4 纵向 2 栏相片、
  docx 报告生成、Word 相片插入、python-docx Pillow、工程竣工报告、安装完成报告、
  设备测试报告、系统验收报告、承建商抬头图、统一签署、蓝色系报告。
  提供：完整 A4 几何常数、自适应网格算法（核心技术：按行数动态缩放）、add_picture_fit
  工具、MD5 去重、正确的 caption 生成方式（避开 _01 后缀 bug）、统一签署区 + 页眉抬头图
  helper、验证清单、坑位大全。可与 officecli-workflow、electrical-test-report-generator、
  site-inspection-memo-generator 等报告生成技能组合使用。
---

# python-docx-photo-grid — Word 报告 + 自适应相片排版 + 工程竣工报告系列

> **这个技能的核心使命：解决「换任务就不识、换会话又不识、换技能又不识」的复发问题。**
> 凡是「生成 docx 报告 + 自动塞相片 + 自动排版 + 统一签署/抬头」的任务，**先看这里**。
>
> 📎 **统一模板/SOP/helper 完整版**：`references/工程竣工报告_模板与SOP.md`
> 

---

## 1. 适用场景（When to use）

| 场景 | 典型任务 |
|---|---|
| **工程竣工报告** | 适用于各类图文并茂测试竣工报告（安装完成 / 设备测试 / 系统验收）|
| 竣工/验收/测试报告 | 设备验收报告、电箱竣工详情 |
| 巡查/勘察记录 | 现场勘察备忘录（蓝色系 Letter 版）、质量检验报告 |
| 任何「文字+大量相片」报告 | 工作总结、施工日志、售后报告 |
| 需要相片排版好看 | 标书附件、客户报告、年报 |

**不适用**：纯文字文档（用 word-typography-guide）；修改现有 docx（用 officecli-workflow）；表格数据为主（用 faithful-xlsx-template）；勘察备忘录专用流程（用 site-inspection-memo-generator，但相片网格技术照用本技能）。

---

## 2. 环境搭建（必做，不要跳过）

### 2.1 Managed Python venv

永远用 **managed venv**，不要污染系统 Python：

```bash
# managed python 位置
PY="<WORKBUDDY_DIR>/binaries/python/versions/3.13.12/python.exe"

# 建 venv（首次）
$PY -m venv "<WORKBUDDY_DIR>/binaries/python/envs/default"

# 装依赖
"<WORKBUDDY_DIR>/binaries/python/envs/default/Scripts/pip.exe" install python-docx Pillow
```

### 2.2 Windows vs Git Bash 路径坑 ⚠️

| 工具 | 路径格式 | 例子 |
|---|---|---|
| **Windows python.exe** | **Windows 反斜杠 `D:\...`** | `r"<WORK_DIR>\xxx.docx"` |
| Git Bash shell 命令 | POSIX `/d/...` | `/d/工作文件/xxx.docx` |

**规则：**
- 喂给 `python -c` 或 `.py` 的路径 → **必须用 `D:\...`**（python-docx 跑在 Windows 下，POSIX 路径会 PackageNotFoundError）
- 在 bash heredoc/cp/ls 用的路径 → POSIX `/d/...`
- venv 内 `python.exe` 在 **`Scripts\python.exe`**，不是 `bin/python`

---

## 3. ⭐ 统一排版规范（工程竣工报告 / 勘察备忘）

> 完整版见 `references/工程竣工报告_模板与SOP.md`。以下是**必须记住的核心**。

### 3.1 纸型（两套，不要捞乱）⚠️

| 文档类型 | 纸型 | 边距 |
|:---|:---|:---|
| **勘察备忘录** | **Letter 21.59 × 27.94 cm** | 上/下 2.6、左/右 2.8 |
| **工程竣工报告** | **A4 21.0 × 29.7 cm** | 左 1.5 固定；上 2.54、下 1.2、右 1.1 |

### 3.2 配色（蓝色系，一文一色系）

深蓝 `#1F4E79`（标题/表头）｜中蓝 `#2E75B6`（副标题）｜浅蓝 `#D6E4F0`｜绿 `#C6EFCE`（合格）｜橙 `#FCE4D6`（取消/待定）｜红 `#F8CBAD`（独立供电）｜灰 `#808080`（图注）｜斑马 `#F2F2F2`。

### 3.3 字体 / 字号

- 中文 eastAsia：**PMingLiU**；ascii/hAnsi：**Times New Roman**（`set_run` 三属性一齐设）。
- 封面 22pt Arial 粗；章节标题 16pt 粗；子标题 13pt；正文 12pt；表格 10pt；图注 10pt 灰。
- 页码：footer `PAGE / NUMPAGES` 域（OxmlElement `w:fldChar` 三件套）。

### 3.4 页眉：承建商公司抬头图

- section header 第一段落 run 内嵌公司 logo，**实际渲染 ~7.5 × 2.0 cm**（交付件实测 6.5~7.7 × 1.8~2.1）。
- ⚠️ 首次用源图 15.92 × 4.32 cm 会**过大**压正文；插入后缩到 ~7.5×2.0，或用大图时上边距调到 ~4.6 cm。
- helper：`add_header_logo(doc, logo_path, width_cm=7.5, height_cm=2.0)`（见附录 §五）。

### 3.5 统一签署区（权威格式，三份报告一致）

```
[16pt 粗体]  签署
测试/安装人员：工程承建商　　　日期：2026-08-28
业主/用户：________________　　　日期：________________
```

- 报告开头资讯区另有「承建商：工程承建商」。
- 签署章节白底黑字（不填色）。helper：`add_signing_block(doc)`（见附录 §六）。

### 3.6 章节结构 — 工程竣工报告（标准模板）

| 报告类型 | 章节结构 |
|:---|:---|
| **安装完成报告** | 标题 → 一、工程概况 → 二、安装完成明细 → 三、现场安装照片 → 签署 |
| **设备测试报告** | 标题 → 一、测试概述 → 二、标准依据 → 三、测试数据 → 四、测试点照片 → 签署 |
| **系统验收报告** | 标题 → 一、工程概况 → 二、验收项目明细 → 三、现场验收照片 → 签署 |

---

## 4. A4 纵向页面几何（核心常数）

```python
# ===== A4 纵向（margin 2.54 cm）=====
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN = 2.54
CONTENT_W = PAGE_W - 2 * MARGIN      # 15.92 cm
CONTENT_H = PAGE_H - 2 * MARGIN      # 24.62 cm
```

如果用其他纸型/边距（如勘察备忘 Letter），改 `PAGE_W/PAGE_H/MARGIN`，后面所有计算自动接着变。

---

## 5. ⭐ 自适应 2 栏相片网格（核心算法）

> **这个是整个技能最值钱的技术。** 解决「固定 2×2 有大片空白」、「固定每页 N 张放不下」的矛盾。

### 5.1 原理

1. 每页最多 `MAX_ROWS × GRID_COLS` 张（默认 3×2 = 6 张）
2. **每页按实际行数动态计算相片高度上限**：
   ```
   h_allow = (PAGE_BUDGET / rows) - CAPTION_H - CELL_PAD - SAFETY
   h_allow = clamp(h_allow, MIN_IMG_H, ABS_MAX_IMG_H)
   ```
3. 该页所有相片按 `h_allow` 为上限缩放（保持长宽比）
4. 因为相片永远可以缩细，所以 3 行 6 张**必定放得落**

### 5.2 为何「固定行数」会失败

常见错误：用「贪婪塞行」（放不落先换页）— 如果相片多是手机直拍（直向图），渲染高度顶到上限，一页怎么算都只放到 2 行 = 4 张，无改善。

**正确做法：行数固定，但相片大小接着行数变。** 3 行就自动缩细到 ~5.9 cm 高；2 行就放大到 9 cm；1 行（尾页剩少）就放到上限。

### 5.3 完整代码模板

```python
# ===== 网格参数 =====
GRID_COLS   = 2          # 固定 2 栏
MAX_ROWS    = 3          # 最多 3 行 → 每页最多 6 张
COL_W       = CONTENT_W / GRID_COLS         # 7.96 cm
IMG_MAX_W   = COL_W - 0.96                  # 7.00 cm
RESERVE_H   = 2.2        # 标题 + 说明 + 行距预留
PAGE_BUDGET = CONTENT_H - RESERVE_H         # 22.42 cm
CAPTION_H   = 0.9
CELL_PAD    = 0.5
SAFETY      = 0.15
ABS_MAX_IMG_H = 9.0      # 单张相绝对高度上限
MIN_IMG_H  = 3.0

# ===== 工具：按比例缩放至上限 =====
from PIL import Image
def rendered_size(path, max_w, max_h):
    img = Image.open(path); ar = img.width / img.height
    w = max_w; h = w / ar
    if h > max_h:
        h = max_h; w = h * ar
    return w, h

def add_picture_fit(paragraph, path, max_w_cm, max_h_cm):
    w, h = rendered_size(path, max_w_cm, max_h_cm)
    run = paragraph.add_run()
    run.add_picture(path, width=Cm(w), height=Cm(h))
    return run

# ===== 表格版面固定（重要）=====
def fix_table_layout(table, col_w):
    table.autofit = False
    tblPr = table._element.tblPr
    layout = tblPr.makeelement(qn('w:tblLayout'), {qn('w:type'): 'fixed'})
    tblPr.append(layout)
    for row in table.rows:
        for cell in row.cells:
            cell.width = Cm(col_w)

# ===== 核心：自适应相片排版 =====
def add_photo_grid(doc, items, caption_size=9):
    """items = [(image_path, caption_text), ...]
    每页 1 个 2 栏网格表，按行数动态缩放相片。"""
    per_page = MAX_ROWS * GRID_COLS
    pages = [items[i:i+per_page] for i in range(0, len(items), per_page)]
    counts = []
    for pi, page in enumerate(pages):
        rows = (len(page) + GRID_COLS - 1) // GRID_COLS
        # 该页相片高度上限
        h_allow = (PAGE_BUDGET / rows) - CAPTION_H - CELL_PAD - SAFETY
        h_allow = max(min(h_allow, ABS_MAX_IMG_H), MIN_IMG_H)

        table = doc.add_table(rows=rows, cols=GRID_COLS)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        fix_table_layout(table, COL_W)

        for idx, (img_path, caption) in enumerate(page):
            ri, ci = divmod(idx, GRID_COLS)
            cell = table.cell(ri, ci)
            # 文字描述在相片上方
            p_cap = cell.paragraphs[0]
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_run(p_cap.add_run(caption), size=caption_size,
                    bold=True, color="595959")
            # 相片（自动缩放）
            p_img = cell.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if img_path and os.path.exists(img_path):
                add_picture_fit(p_img, img_path, IMG_MAX_W, h_allow)
            else:
                set_run(p_img.add_run("[相片缺失]"), size=9, color="C00000")

        # 空白储存格补齐
        for idx in range(len(page), rows * GRID_COLS):
            ri, ci = divmod(idx, GRID_COLS)
            table.cell(ri, ci).text = ""

        set_table_borders(table)
        counts.append(len(page))
        if pi < len(pages) - 1:
            doc.add_page_break()
    return counts
```

> 大量相（如 Lux 29 张）用上边自适应网格；少量相（风扇 2-4 张）可用附录 §八的 `add_photo_grid()`（2 列无框线版，相放大到上限）。

### 5.4 成效对比（真实数据）

| 报告 | 相片数 | 固定 2×2 | 自适应 |
|---|---|---|---|
| 灯具安装 18 张 | 18 | 5 页 | **3 页** [6,6,6] |
| Lux 照度 29 张 | 29 | 8 页 | **5 页** [6,6,6,6,5] |
| 风扇 3 张 | 3 | 1 页 | 1 页（2 行，相放大到 9 cm） |

---

## 6. ⭐ Caption 生成（避开 `_01` 后缀 Bug）

### 6.1 错的做法

```python
# 假设档名 "132室_465lux.jpg"
base = os.path.basename(ph)          # "132室_465lux.jpg"
lux_val = base.split("_")[-1].replace("lux.jpg", "")
# → "465"  OK

# 但若档名是 "132室_465lux_01.jpg"（重复相，_01 后缀）
base = "132室_465lux_01.jpg"
lux_val = base.split("_")[-1].replace("lux.jpg", "")
# → "01.jpg"  ❌  caption 变成 "132室 — 01.jpg lux"
```

**Bug 确认**：用户在 Lux 报告中手动修正了这个 bug（将 caption 改为 "132室 — 465.6 lux" 用平均值）。

### 6.2 正确做法：用数据表的平均值/编号

```python
# LUX_DATA = {"132室": [465, 502, 505, 606]}
loc_avg = round(sum(LUX_DATA[loc]) / len(LUX_DATA[loc]), 1)  # 519.5
caption = f"{loc} — {loc_avg} lux"   # "132室 — 519.5 lux"（交付件实测格式）

# 或者用测点编号
caption = f"{loc} 测点 {idx+1}"       # "132室 测点 1"
```

**原则：caption 内容必须来自结构化数据，不要由档名解析。** 档名只是 ID，不是数据。

---

## 7. MD5 去重（必做）

微信/相机重复汇出的相，MD5 会完全相同。不去重会导致：
- python-docx 自动重用同一张图，导致相片数对不上 glob 数量
- 报告出现重复相

```python
import hashlib, glob
seen_md5 = set()
for loc in LOCATIONS:
    pat = os.path.join(PHOTO_DIR, f"{loc}_*lux*.jpg")
    for ph in sorted(glob.glob(pat)):
        with open(ph, 'rb') as fp:
            m = hashlib.md5(fp.read()).hexdigest()
        if m in seen_md5:
            continue
        seen_md5.add(m)
        # ... 加入 items
```

---

## 8. 压缩相片（控文件大小）

```python
def compress_image(src, dst, max_width_cm=12, dpi=200):
    img = Image.open(src)
    max_w = int(max_width_cm * dpi / 2.54)
    if img.width > max_w:
        ratio = max_w / img.width
        img = img.resize((max_w, int(img.height * ratio)), Image.LANCZOS)
    img.save(dst, "JPEG", quality=85, dpi=(dpi, dpi))
```

`max_width_cm=12, dpi=200, quality=85` 是经验值，文件大小与质量的最佳平衡。
3 份报告（18+29+3=50 张相）总文件 ~3 MB，合理。

---

## 9. 验证清单（每次生成后必做）

### 9.1 相片网格验证（python-docx 读回）

```python
# scripts/verify_docx.py
from docx import Document
from docx.oxml.ns import qn
import os

EMU = 360000
PAGE_BUDGET = 22.42
CONTENT_H   = 24.62

def verify(path, expect_imgs):
    d = Document(path)
    tables = d.tables
    imgs = [r for r in d.part.rels.values() if "image" in r.reltype]

    # 1. 数据表格 vs 相片表格
    photo_pages = []
    for t in tables:
        grid = []
        for row in t.rows:
            hs = []
            for c in row.cells:
                got = None
                for para in c.paragraphs:
                    for ext in para._p.iter(qn('wp:extent')):
                        got = int(ext.get('cy')) / EMU
                hs.append(got)
            grid.append(hs)
        if any(any(h is not None for h in hs) for hs in grid):
            n_img = sum(1 for hs in grid for h in hs if h is not None)
            est = sum(
                (max([h for h in hs if h is not None]) if any(h is not None for h in hs) else 0)
                + 0.9 + 0.5  # CAPTION_H + CELL_PAD
                for hs in grid
            )
            photo_pages.append({
                "rows": len(grid), "cols": len(t.columns),
                "n_img": n_img, "est_h": round(est, 2)
            })

    total_img = sum(p["n_img"] for p in photo_pages)
    shape_ok = all(p["rows"] <= 3 and p["cols"] == 2 for p in photo_pages)
    over = [p for p in photo_pages if p["est_h"] > PAGE_BUDGET + 0.01]

    print(f"{os.path.basename(path)}")
    print(f"  表格={len(tables)} 相片页={len(photo_pages)} 嵌入相={len(imgs)} (预期{expect_imgs}) 表格内相={total_img}")
    print(f"  每页: {photo_pages}")
    print(f"  形状OK={shape_ok}  超出预算={len(over)}")
    assert imgs and total_img == len(imgs), "嵌入相 vs 表格内相不一致"
    assert shape_ok, "相片表格形状错误"
    assert not over, "有页超出 A4 预算"
    print("  ✅ PASS")
```

### 9.2 工程竣工报告完整性检查（加埋这些）

- 页眉有承建商 logo（header 图片存在，宽 ~7.5cm）
- 最后一章是「签署」（16pt bold），含「测试/安装人员：工程承建商」+「业主/用户：____」两行
- 报告开头有「承建商：工程承建商」
- 章节标题顺序符合标准模板（见 §3.6）
- 纸型正确（竣工报告 A4 / 勘察备忘 Letter）
- 页脚有 `PAGE / NUMPAGES` 域

---

## 10. 坑位大全（读一次受用一世）

| # | 坑 | 症状 | 解决 |
|---|---|---|---|
| 1 | managed venv 路径写 `bin/python` | 找不到 python | Windows venv 是 **`Scripts\python.exe`** |
| 2 | 喂 POSIX `/d/...` 路径给 python-docx | PackageNotFoundError | python 跑在 Windows 下，要用 `D:\...` |
| 3 | caption 由档名解析（`_01` 后缀） | "132室 — 01.jpg lux" 垃圾文字 | caption 内容用结构化数据，不要解析档名 |
| 4 | 相片重复汇出无去重 | 嵌入相数 ≠ glob 数 | MD5 set 去重 |
| 5 | 用 `run.add_picture(width=, height=)` 两个都设但 aspect 错 | 相片变形 | 计算好 aspect 后只设限制维度，或两个都按 aspect 算 |
| 6 | TableGrid 预设 `autofit=True` | Word 自动调栏宽，2 栏变 1 栏 | `fix_table_layout()`：tblLayout fixed + autofit=False + 锁 cell.width |
| 7 | 「贪婪塞行」配直向相 | 一页只 2 行 4 张，无改善 | 改用「行数固定、相片按行数缩放」（核心算法） |
| 8 | Word 保存报「权限错误」 | 文件被设唯读 / 锁住 | `attrib -R file.docx` + `os.chmod(file, 0o777)` |
| 9 | Read tool 读图报「Content filtered」 | 模型话看不到 | 靠 cache 机制逐张确认；或用 MD5 交叉验证 |
| 10 | 文件首次存在 工作目录 | Sandbox 静默拒绝，改放到工作目录 |
| 11 | 页眉图太大（15.92×4.32）压正文 | 正文被 logo 覆盖 | 缩到 ~7.5×2.0 cm，或上边距调到 ~4.6 cm |
| 12 | `add_body()` 传两个 text 参数 | "文字2" 被当 size 传入 `Pt()` 报错 | `+` 拼接或分两次调用 |
| 13 | `set_run()` 只设 `run.font.name` | 中文字体不生效 | `w:ascii` + `w:hAnsi` + `w:eastAsia` 三属性一齐设 |
| 14 | `set_cell()` 直接 `cell.text=""` | 残留 empty run 产生 bug | 先清 runs 再 `add_run()` |
| 15 | `set_col_widths()` 只设 `cell.width` | 栏宽飘 | 同时改 `tblGrid`（twips，1cm≈567）+ `cell.width`（Cm） |
| 16 | Python heredoc `\U` 截断 | `SyntaxError: truncated \UXXXXXXXX` | 一律用 Write 写 `.py` 档再执行 |
| 17 | 签署区/表头误填彩色 | 用户要求白底黑字 | 签署章节白底黑字不填色；ME 测试表白底黑字 |

---

## 11. 与其他技能的关系

```
任务：生成一份工程竣工报告（安装/测试/验收类型 / 勘察备忘）
   ↓
路由规则
   ↓
electrical-test-report-generator  ← 测试表格风格（ME 白底黑字、页码页脚）
   +
python-docx-photo-grid（本技能）← 统一排版（蓝色系/抬头/签署）+ 自适应相片网格
   +
site-inspection-memo-generator  ← 勘察备忘录专用流程（Letter/蓝色系/动作栏）
   ↓
生成 .docx
   ↓
用户手改（officecli-workflow 做局部改 OR 用户直接 Word 改）
   ↓
correction-diff-capture  ← 读用户改完的版本，diff 反馈到本技能
```

| 技能 | 关系 |
|---|---|
| `electrical-test-report-generator` | 测试表格风格基底，本技能加相片排版 + 统一签署 |
| `site-inspection-memo-generator` | 勘察备忘录专用（Letter/动作栏色码）；相片网格技术照用本技能 |
| `officecli-workflow` | **生成后**用户手改时，用 OfficeCLI 局部改（快），不要 regen |
| `correction-diff-capture` | 用户改完交付，读取+diff 反馈到本技能的坑位/常数 |
| `word-typography-guide` | 纯文字文档用这个；本技能用于「文字+相片」混合 |
| `gantt-chart-pro` / `faithful-xlsx-template` | Excel 场景用这些 |

---

## 12. 完整最小可运行范例

```python
# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from PIL import Image

# === 环境 ===
PHOTO_DIR = r"D:\path\to\photos"
LOGO      = r"D:\path\to\company_logo.jpg"     # 承建商公司抬头图
OUT       = r"D:\path\to\output\report.docx"
TEMP      = r"<WORKSPACE>\temp\_photos"
os.makedirs(TEMP, exist_ok=True)

# === A4 几何（竣工报告；勘察备忘改 Letter 21.59×27.94）===
PAGE_W, PAGE_H = 21.0, 29.7
MARGIN = 2.54
CONTENT_W = PAGE_W - 2 * MARGIN
CONTENT_H = PAGE_H - 2 * MARGIN
GRID_COLS, MAX_ROWS = 2, 3
COL_W = CONTENT_W / GRID_COLS
IMG_MAX_W = COL_W - 0.96
PAGE_BUDGET = CONTENT_H - 2.2
CAPTION_H, CELL_PAD, SAFETY = 0.9, 0.5, 0.15
ABS_MAX_IMG_H, MIN_IMG_H = 9.0, 3.0

# === 工具函数（§5.3 + 附录 helper 全集）===
# ... rendered_size, add_picture_fit, fix_table_layout, add_photo_grid,
#     set_run, set_cell, add_page_number, add_header_logo, add_signing_block ...

# === 生成 ===
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(PAGE_W), Cm(PAGE_H)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Cm(MARGIN)

add_header_logo(doc, LOGO)                  # 页眉承建商抬头图
# ... 标题、项目信息（承建商：工程承建商）、数据表格 ...
items = []  # [(temp_image_path, "128室 — 484 lux"), ...]
add_photo_grid(doc, items)                  # 自适应相片网格
add_signing_block(doc)                      # 统一签署区
add_page_number(doc)                        # 页脚 PAGE / NUMPAGES
doc.save(OUT)
print("saved:", OUT)
```

---

## 13. 自我进化记录

| 日期 | 变更 | 触发 |
|---|---|---|
| 2026-08-28 | 初版建立 | 学校灯具 Lux 报告任务（v1~v4 四次迭代） |
| 2026-08-28 | 加入 `_01` 后缀 caption bug 教训 | 用户手改 Lux 报告修正 caption |
| 2026-08-28 | 确认自适应网格公式 | 解决「直向相塞不落 3 行」问题 |
| 2026-09-01 | **统一升级**：合并知识库 SOP + 蓝色系模板 + python-docx 踩坑 + 三份交付件实测（承建商抬头图 7.5×2.0、统一签署区、A4/Letter 两套纸型、三件套章节结构），新增 `references/竣工报告三件套_模板与SOP.md` | 用户要求把竣工报告三件套能力统一封装入本技能 |
