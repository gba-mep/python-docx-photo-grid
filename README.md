# python-docx-photo-grid · Word 自适应相片排版

<div align="center">

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://python.org)
[![Part of the MEP Automation Toolkit](https://img.shields.io/badge/Toolkit-MEP%20automation-1565C0?logo=github&logoColor=white)](https://github.com/David-CB666)

![Adaptive Photo Grid Comparison](assets/photo-grid-comparison.jpg)

**Word report generation with adaptive 2-column photo grid algorithm**

A4 geometry · MD5 dedup · Caption generation · Full helper library

[快速开始](#快速开始) · [文件结构](#文件结构) · [技术栈](#技术栈)

</div>

---

> Python-based Word (.docx) 报告生成，配**自适应 2 栏相片网格**算法。解决「固定 2×2 布局浪费空间」的经典问题，按每页实际行数动态缩放相片。

## 解决什么问题

生成有大量相片的 Word 报告时：
- 固定 2×2 排版，行数少时相片太细、浪费空间
- 手动拖相对位，几十张相搞几个钟
- 不同报告（灯具/Lux/风扇）格式不统一
- caption 编号、页码、签署区每次重复做

**python-docx-photo-grid** 将这些全部自动化。

## 核心特性

### 🖼️ 自适应相片网格算法
- 最多 3 行 × 2 栏 = 6 张/页
- 按每页实际相片数量动态缩放
- A4 / Letter 页面几何常数精准计算
- 长宽比自动适配，不变形

### 📸 相片处理管线
- MD5 相片去重（重复相自动跳过）
- 图片压缩（控制文件大小）
- caption 自动生成（避开 `_01` 后缀 bug）

### 📋 完整 Helper 库
- 字体设定（中英文混排）
- 表格边框样式
- 页码 + 页眉抬头图
- 统一签署区块

### 🧪 验证工具
- 结构完整性检查脚本
- 17 个 documented pitfalls

### 🏗️ 工程竣工报告模板
适用于工程各类图文并茂测试竣工报告
（统一蓝色系 + 承建商抬头图 + 章节结构）

> 🔒 **工程竣工报告模板与SOP** 为非公开内容，不在此公开 repo 中。
> 包含完整模板、排版规格、SOP 流程、实机交付件参数。
> 如需商业使用，请邮件联络：**david_1999cn@hotmail.com**

## 适用场景

| 场景 | 例子 |
|:---|:---|
| 竣工/验收/测试报告 | 灯具/Lux/风扇/电箱/设备验收 |
| 巡查/勘察记录 | 现场勘察备忘录、质量检验报告 |
| 任何「文字+大量相片」报告 | 施工日志、售后报告、工作总结 |
| 需要相片排版好看 | 标书附件、客户报告、年报 |

**不适用**：纯文字文档；修改现有 docx（用 officecli-workflow）；表格数据为主。

## 文件结构

```
python-docx-photo-grid/
├── README.md                          # 本文件
├── DOCUMENTATION.md                   # 完整技能文档
├── assets/
│   └── photo-grid-comparison.jpg      # 效果对比图
├── references/
│   └── 竣工报告三件套_模板与SOP.md 🔒 # 模板 & SOP 参考（非公开，需邮件授权）
└── scripts/
    └── adaptive_photo_grid.py         # 可重用程式码模组
```

## 技术栈

- **Python** + **python-docx** — Word 文档生成
- **Pillow** — 图片处理与尺寸计算

## 快速开始

```python
from adaptive_photo_grid import PhotoGridBuilder

builder = PhotoGridBuilder(page_size="A4", columns=2, max_rows=3)
builder.add_photos(["photo1.jpg", "photo2.jpg", "photo3.jpg"])
builder.save("report.docx")
```

详细用法请参阅 [DOCUMENTATION.md](DOCUMENTATION.md)。

---

## License

MIT License — feel free to use, modify, and share.

---

## Related repositories

Part of the **[MEP & construction document automation toolkit](https://github.com/David-CB666)** — open-source tools built from real jobsite workflows.

- **Handbook** — [ai-agent-manual](https://github.com/gba-mep/ai-agent-manual) (8-level AI cultivation for engineers)
- **Document generation** — [material-approval-pipeline](https://github.com/gba-mep/material-approval-pipeline) · [material-submittal-generator](https://github.com/gba-mep/material-submittal-generator) · [excel-template-filler](https://github.com/gba-mep/excel-template-filler) · [daily-construction-log](https://github.com/gba-mep/daily-construction-log) · [officecli-workflow](https://github.com/gba-mep/officecli-workflow)
- **Engineering calculation** — [lighting-lux-calculator](https://github.com/gba-mep/lighting-lux-calculator) · [ups-discharge-time-calculator](https://github.com/gba-mep/ups-discharge-time-calculator) · [gantt-chart-pro](https://github.com/gba-mep/gantt-chart-pro) · [electrical-test-report-generator](https://github.com/gba-mep/electrical-test-report-generator)
- **CAD & drawings** — [electrical-panel-label-plates](https://github.com/gba-mep/electrical-panel-label-plates)
- **Data & OCR** — [ocr-skill](https://github.com/gba-mep/ocr-skill) · [VBA-Macro-Reader-v2.0.0](https://github.com/gba-mep/VBA-Macro-Reader-v2.0.0)
- **Compliance & AI ops** — [confined-space-planner](https://github.com/gba-mep/confined-space-planner) · [路由规则](https://github.com/gba-mep/路由规则) · [consulting-services](https://github.com/gba-mep/consulting-services)
