# 设计规范来源说明

本目录（`design_system/`）集成了 **ACKS Studio 文档设计规范 v2.1**，使 office_suite
产出的 Word / PowerPoint / Excel 文档直接符合规范。

## 来源

- **来源路径**：Thom & 银月共享知识库 `05-Resources/文档设计规范-v2/`
- **版本**：v2.1（v2.0 的 Token 单一信源工程化 + v2.1 的标题溢出/章节编号/页脚回填等修复）
- **复制日期**：2026-08-23
- **授权**：Thom 明确授权将设计规范内容复制一份到 office 库中

## 复制内容

| 内容 | 说明 |
|------|------|
| `tokens.json` | Token 单一信源（颜色/字体/间距/页面/slide/xlsx/glyph/content_limit） |
| `tokens.py` / `slides.py` / `xlsx.py` | Word / PPT / Excel 代码生成模块 |
| `tokens_constants.py` / `slides_constants.py` / `xlsx_constants.py` | 由 tokens.json 自动生成的常量 |
| `acks/` | 设计系统 CLI 包（`python -m acks build/validate/tokens/fonts`） |
| `scripts/` | build_tokens.py（生成常量）/ validate.py（校验）/ download_fonts.py（字体下载） |
| `tests/` | 设计系统测试（29 条用例） |

## 未复制内容（保留在知识库）

- HTML 规范文档（`ACKS Studio 文档设计规范.html` 等）
- `css/`、`js/`、`examples/`（HTML 示例模板）
- `assets/logo-*.png`（品牌 logo，未被 office_suite 代码引用，保留在知识库）

以上属于设计规范的知识库参考资料（给人查阅），非 office_suite 运行依赖，故不复制。

## 本地改动

- 三个代码生成模块的 constants import 由绝对 import 改为**相对 import**
  （`from tokens_constants import *` → `from .tokens_constants import *`），
  以适配 Python 包结构 `office_suite.design_system`。

## 同步策略

本目录与知识库源**会分叉**。若 v2 更新，需手动同步，遵循 v2 README 的 Token 修改流程：

1. 编辑 `tokens.json`（唯一手动改的地方）
2. 运行 `python scripts/build_tokens.py` 重新生成 `*_constants.py`
3. 同步更新 HTML 规范文档（在知识库侧）
4. 运行 `python scripts/validate.py` 确认一致性

> 品牌参数化说明：v2 原代码硬编码「ACKS STUDIO · 爱驰科驶」品牌名、logo 与页脚 label。
> office_suite 集成时将这些抽为可配置参数（默认保留 ACKS 品牌，支持去品牌化/换品牌），
> 具体见 `office_suite/core.py` 与各格式模块的 `theme` / `brand_*` 参数。
