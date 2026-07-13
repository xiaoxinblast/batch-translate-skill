---
name: batch-translate
description: 批量翻译工作流。支持 mqxliff/docx/xlsx/txt → 分批翻译 → 逐批校对 → 写回，全自动循环。 触发条件：用户说"开始批量翻译""批量翻译这个文件""继续翻译"等。
---

# batch-translate — 批量翻译工作流

## 触发条件

用户表达要批量翻译文件时触发本 skill。如：
- "开始批量翻译"
- "批量翻译这个文件"
- "继续翻译下一批"

## 阶段〇：安装/更新工具包

> ⚠️ **强制步骤**：每次触发本 skill 必须首先执行，不可跳过。

### 情况 A：`batch_translate/` 不存在

```bash
git clone https://github.com/xiaoxinblast/batch-translate.git batch_translate
pip install lxml openpyxl python-docx pdfplumber
mkdir -p batch_translate/data batch_translate/exports
```

目录结构（按源文件 stem 分组）：
- `data/<stem>/_working_*, batch_state.json` ← 每个源文件独立
- `exports/<stem>/_batch_NNN_*.json` ← 每个源文件独立
- `data/tm_memory.json, term_base.xlsx, style_guide.txt` ← 共享

### 情况 B：`batch_translate/` 已存在

**必须检查并更新到最新版本**：

```bash
cd batch_translate
# 若不是 git 仓库（旧版安装），删除后重新克隆
if [ ! -d .git ]; then
  cd .. && rm -rf batch_translate
  git clone https://github.com/xiaoxinblast/batch-translate.git batch_translate
  pip install lxml openpyxl python-docx pdfplumber
else
  # 已是 git 仓库，拉取最新
  git fetch origin && git merge origin/main --ff-only 2>/dev/null || git reset --hard origin/main
fi
mkdir -p data exports
cd ..
```

如果远端有更新，合并完成后展示更新日志：

```bash
cd batch_translate && git log --oneline -3 && cd ..
```

合并冲突导致失败时，硬重置到远端最新：

```bash
cd batch_translate && git reset --hard origin/main && cd ..
```

确认 `batch.py` 可执行且依赖已安装后，进入阶段〇.五。

## 阶段〇.五：项目文件扫描

> ⚠️ **强制步骤**：必须在阶段一之前执行。目的是**先摸清整个项目有什么，再决定用什么**。

### 1. 列出根目录
```bash
ls -la
```

### 2. 多维度 Glob 扫描

用 glob 工具做以下搜索（并行执行）：

- 术语类：`**/*术语*` `**/*用語*` `**/*term*` `**/*glossary*` `**/*词汇*`
- 风格指南类：`**/*翻译指南*` `**/*翻訳*` `**/*方針*` `**/*style*guide*` `**/*ローカライズ*` `**/*本地化*`
- 翻译记忆类：`**/*tm*` `**/*memory*` `**/*翻译记忆*`
- 前作/参考类：`**/*前作*` `**/*master*` `**/*マスター*`

### 3. 评估找到的文件

列出所有找到的文件，标注大小（文件大小/行数）。

**优先级判断原则**（通用，不绑定具体文件名）：
- 根目录的文件优先于子目录文件（根目录往往是汇总后的完整版）
- 行数/大小大的优先于小的（大的往往是完整版，小的是摘录）
- `.xlsx` 多 sheet 的优先于单 sheet
- `.docx` 正文长的优先于短的
- 文件修改时间新的优先于旧的

### 4. 确认参考文件来源

> ⚠️ **强制步骤**：扫描后、生成任何文件前，必须询问用户。

向用户展示扫描结果，然后用 `ask` 工具（multiSelect）依次询问：

1. **风格指南来源**：列出找到的候选文件，让用户勾选。推荐项标记"(Recommended)"
2. **术语库来源**：列出找到的候选 xlsx 文件，让用户勾选。推荐项标记"(Recommended)"
3. **翻译记忆来源**：
   - 选项 A："从零开始（空 TM）" (Recommended)
   - 选项 B："从源文件导入已有译文"
   - 选项 C："从指定 TM 文件导入"

> ⚠️ **重要警告**：若用户选择"从源文件导入已有译文"，必须提醒：
> "源文件中已有的译文可能已过时——客户端可能修改了源文但未同步更新译文。导入后需逐条核对。"

用户确认后，进入阶段〇.七。

## 阶段〇.七：子 task 模型分配

> 本 skill 通过 Reasonix 全局 subagent profiles 调度子任务：
> - `/translator` — 翻译（deepseek-pro, effort max）
> - `/trans-reviewer` — 校对（deepseek-pro, effort max）
> - `/context-analyzer` — 语境分析（deepseek-flash）
>
> 模型和 effort 在 `reasonix.toml` 的 `subagent_models` 中配置，无需在本 skill 内指定。

## 阶段一：项目初始化

若 `batch_translate/data/<stem>/batch_state.json` 不存在，则尚未初始化。`<stem>` 为源文件名（不含扩展名）。

### 1. 生成风格指南 → `batch_translate/data/style_guide.txt`

> ⚠️ **强制步骤**：必须在 init 之前完成，不可跳过。

使用用户在阶段〇.五步骤 4 中确认的风格指南来源文件。

#### 提取方法（避免 Shell 和编码问题）
- 用 python heredoc：`python << 'PYEOF' ... PYEOF`
- 输出重定向到文件后用 read_file 读取，避免 GBK 终端编码错误
- docx 用 python-docx，xlsx 用 openpyxl，pdf 用 pdfplumber

#### 编译要求
- 多来源时去重、整理为统一结构
- **源文档要求的语言风格必须在风格指南中占据显著位置**，附正反对比示例（从源文档提取，不自行编造）
- **必须包含以下硬性格式规则**：
  - **引号**：中文译文只能用弯引号 `""`（U+201C / U+201D），**绝对禁止**使用日式角引号 `「」`（U+300C / U+300D）
  - **破折号**：简体字使用 `——`（U+2014）
  - **省略号**：使用 `……`（两个三点省略号）
  - JSON 输出中的引号必须与上述规则一致，禁止用 ASCII `"`（U+0022）冒充中文引号
- **必须包含以下日中翻译注意事项**（正反示例各至少一个）：
  1. **主语/人称代词补充**：日语常省略主语，中文须视上下文补出，避免读者不知所指
  2. **汉字词转换**：日文中的汉字词不能直接照搬（如「検討」→"研究"而非"检讨"），须换为中文固有说法
  3. **语序调整**：日语 SOV → 中文 SVO，宾语、补语位置须重组
  4. **被动转主动**：日语惯用被动表达，中文多数情况须转为主动句
  5. **长句拆分**：日语的连用形可串起多个从句，中文须切成短句、层次分明
  6. **长定语后置**：日语的连体修饰语可极长，中文的前置定语须精简或后置
  7. **暧昧表达去/留**：日语好「～と思われる」「～ようだ」，中文视文本类型决定是否保留推测语气
  8. **指示词具体化**：日语的「これ/それ/あれ」尽量还原为具体名词，避免中文满篇"这个""那个"
  9. **敬语落差弥合**：日语敬语体系（尊敬/謙譲/丁寧）在中文中无对等物，须靠措辞（"您""在下""请"）和语气弥合身份差
- 不标注来源文件路径
- **必须用 write_file 工具将风格指南写入文件**
- 确认文件存在后才进入步骤 2

### 2. 生成术语库 → `batch_translate/data/term_base.xlsx`

使用用户在阶段〇.五步骤 4 中确认的术语库来源文件。

#### 强制规则
- ⚠️ **读取任何参考文件时，必须读取全部行。禁止使用 `max_row` 或 `min()` 截断行数**
- 自适应识别列结构：查找「日文/日本語/ja/JP」「中文/中国語/zh/CN/SC」列
- 注释列如有分类信息，保留到注释字段
- 多来源时以「日文」列为 key 去重
- 输出时报告：总条数、来源文件数、各分类分布
- 以 xlsx 三列格式写入：`原文(ja) | 译文(zh) | 注释`
- 如用户选择不导入任何术语文件 → 运行 `term_base.py --create` 生成空白模板

### 3. 创建翻译记忆

根据用户在阶段〇.五步骤 4 中的选择：

**选项 A（从零开始）**：
```python
import json, os
os.makedirs('batch_translate/data', exist_ok=True)
with open('batch_translate/data/tm_memory.json', 'w') as f:
    json.dump({"entries": []}, f)
```

**选项 B（从源文件导入）**：
> ⚠️ 仅导入**源文与译文匹配的条目**。如发现译文可能过时（如源文含修订标记），跳过该条目并在报告中标注。

**选项 C（从指定 TM 文件导入）**：复制用户指定的 JSON 文件到 `batch_translate/data/tm_memory.json`。

### 4. 询问用户源文件并运行 init

```bash
python batch_translate/batch.py init <源文件> \
  --batch-chars 3000 --context-size 5 \
  --terms batch_translate/data/term_base.xlsx \
  --tm batch_translate/data/tm_memory.json \
  --style-guide batch_translate/data/style_guide.txt
```

xlsx 格式需额外指定 `--source-col A --target-col B`。

### 5. 全量语境分析

> ⚠️ **强制步骤**：init 后必须立即执行，不可跳过。

**必须**调用 `/context-analyzer`，arguments 为：

> 读取 `batch_translate/exports/<stem>/_working.json`，做全量语境分析。
>

task 返回报告后，**必须**将内容写入 `batch_state.json` 的 `document_summary` 字段。方法：

```python
import json
state = json.load(open('batch_translate/data/<stem>/batch_state.json'))
state['document_summary'] += "\n\n" + agent_report  # agent_report 是 task 返回的文本
json.dump(state, open('batch_translate/data/<stem>/batch_state.json', 'w'), ensure_ascii=False, indent=2)
```

### 6. 术语缺口确认

> ⚠️ **强制步骤**：语境分析报告若含「疑似术语库未覆盖的专名」清单（内容非"无"），必须在开译前处理。否则这些专名会被翻译 task 就近臆测（如特定角色名被 task 凭上下文臆测出一个看似合理但错误的译名），且校对 task 只对"是否符合术语库"负责、术语库缺失时无法纠正。

1. 从步骤 5 的报告中提取「疑似术语库未覆盖的专名」清单。若为"无"，跳过本步骤直接进入阶段二。
2. 用 `ask` 工具把清单呈现给用户，请其对每个专名**定名**。每项列出：原文、建议译名、出现频次/上下文；选项给"采纳建议译名 / 自定义（Other）/ 跳过（暂不入库）"。专名较多时分组多次询问。
3. 用户定名后，把确认的 `原文 | 译名 | 注释` 追加写入术语库（去重，跳过已存在的原文）——把 `new_terms` 填成用户的定名结果：

```bash
python << 'PYEOF'
import openpyxl
# new_terms：填入用户定名结果，格式 [(ja, zh, note), ...]
new_terms = [
    # ("原文", "译名", "注释"),
]
p = "batch_translate/data/term_base.xlsx"
wb = openpyxl.load_workbook(p)
ws = wb.active
existing = {str(r[0].value).strip() for r in ws.iter_rows(min_row=2) if r[0].value}
added = 0
for ja, zh, note in new_terms:
    if ja.strip() in existing:
        continue
    ws.append([ja, zh, note])
    added += 1
wb.save(p)
print(f"已新增 {added} 条术语（跳过 {len(new_terms) - added} 条已存在）")
PYEOF
```

4. 写入术语库后**必须重新注入术语**到待翻译 JSON（否则新术语不会出现在各批的 `terms` 字段中）。复用 batch.py 现有的 enrichment（不重跑 init、不影响进度）；把 `STEM` 替换为实际 stem：

```bash
python << 'PYEOF' > batch_translate/exports/enrich_out.txt 2>&1
import sys, json
from pathlib import Path
sys.path.insert(0, "batch_translate")
import batch
STEM = "<stem>"  # ← 替换为实际 stem
state = json.load(open(f"batch_translate/data/{STEM}/batch_state.json", encoding="utf-8"))
batch._enrich_working_json(Path(state["export_file"]), state)
print("✅ 新术语已重新注入 _working.json")
PYEOF
```

用 read_file 读取 `enrich_out.txt` 确认输出 `✅` 后，进入阶段二。

## 校对模式（已有译文）

如果文件已有译文，只需校对：

```bash
python batch_translate/batch.py next --review
```

直接生成校对 JSON（`_batch_NNN_to_review.json`），跳过翻译 task，直接走校对→submit。

## 阶段二：全自动翻译循环

反复执行以下步骤直到全部批次完成：

### Step 1: 获取当前批

```bash
python batch_translate/batch.py next
```

输出格式：`_batch_NNN_to_translate.json`

### Step 2: 翻译

调用 `/translator`，arguments 为：

> 读取 `_batch_NNN_to_translate.json`，翻译所有 entries 的 source 字段，按输出契约将结果写入 `_batch_NNN_translated.json`。


task 返回后，**必须确认**文件存在，否则重试。

### Step 3: 生成校对文件

```bash
python batch_translate/batch.py review _batch_NNN_translated.json
```

输出格式：`_batch_NNN_to_review.json`（source + translated 对照）

### Step 4: 校对

调用 `/trans-reviewer`，arguments 为：

> 读取 `_batch_NNN_to_review.json`，逐条校对所有 entries，按输出契约将结果写入 `_batch_NNN_reviewed.json`。


task 返回后，**必须确认** `_batch_NNN_reviewed.json` 文件存在且包含全部条目，否则重试。

### Step 4.5: 机制化验证校对 JSON

> ⚠️ 校对 task 返回后、提交前，**必须**跑以下验证。它复用 P0 同款逻辑（读 `batch_state.json` 取本批预期 id 与 source），是提交前的流程层前置拦截，与 `submit` 内的工具层兜底构成双保险。**只读、不改文件**。

把 `STEM` 替换为实际源文件 stem 后运行（批号自动从 state 读取，无需手填）：

```bash
python << 'PYEOF' > batch_translate/exports/verify_out.txt 2>&1
import json, re
from pathlib import Path

STEM = "<stem>"  # ← 替换为实际源文件 stem
base = Path("batch_translate")
state = json.load(open(base / "data" / STEM / "batch_state.json", encoding="utf-8"))
export_data = json.load(open(state["export_file"], encoding="utf-8"))
bi = state["current_batch"]
start, end = state["batches"][bi]
batch_entries = export_data["entries"][start:end]
expected = [str(e["id"]) for e in batch_entries]
src_by_id = {str(e["id"]): e.get("source", "") for e in batch_entries}

reviewed = base / "exports" / STEM / f"_batch_{bi + 1:03d}_reviewed.json"
data = json.load(open(reviewed, encoding="utf-8"))

# 容错：自动解包 {"entries":[...]}
if isinstance(data, dict) and "entries" in data:
    print("⚠️ 检测到包装对象 {entries:[...]}，已自动解包（Agent 应输出扁平数组）")
    data = data["entries"]
if not isinstance(data, list):
    print("❌ 不是 JSON 数组 → 退回 Step 4 重跑校对")
    raise SystemExit(1)

sub = [str(r.get("id")) for r in data]
sub_set, exp_set = set(sub), set(expected)
missing = exp_set - sub_set
extra = sub_set - exp_set

# 致命：条数不符 / id 未全覆盖 → 阻止提交
if len(data) != len(expected) or missing:
    print(f"❌ 条数不符：预期 {len(expected)} 条，实际 {len(data)} 条")
    if missing:
        print(f"   缺失 id（前 20）：{sorted(missing)[:20]}")
    print("   → 校对 task 可能只输出了改动条。退回 Step 4，明确要求'输出本批全部条目'后重跑。")
    raise SystemExit(1)

# 警告：标签数不一致 / 多余 id → 列出但不阻止（与 P0 分级一致）
TAG = re.compile(r"<tag\s+id=['\"][^'\"]+['\"].*?/>")
tag_bad = [str(r["id"]) for r in data
           if "<tag" in src_by_id.get(str(r["id"]), "")
           and len(TAG.findall(src_by_id.get(str(r["id"]), ""))) != len(TAG.findall(r.get("target") or ""))]
if tag_bad:
    print(f"⚠️ {len(tag_bad)} 条标签数与 source 不一致：{tag_bad[:20]}（建议退回 Step 4 修正后再 submit）")
if extra:
    print(f"⚠️ {len(extra)} 条 id 不属于本批：{sorted(extra)[:20]}")

print(f"✅ 校验通过：{len(data)} 条，本批 id 全覆盖" + ("（有警告见上）" if (tag_bad or extra) else ""))
PYEOF
```

用 read_file 读取 `batch_translate/exports/verify_out.txt` 查看结果（避免终端 GBK 编码错误）。

**分级处理**：
- 打印 `✅ 校验通过` → 进入 Step 5 提交。
- 打印 `❌ 条数不符 / 不是 JSON 数组`（脚本 exit 1）→ **退回 Step 4 重跑校对**，prompt 中明确强调"必须输出本批全部 N 条，不得只输出改动条"。
- 仅有 `⚠️ 标签数不一致` → 退回 Step 4 修正对应条目的标签后重跑；确认无误再进 Step 5。

### Step 5: 提交并推进

```bash
python batch_translate/batch.py submit _batch_NNN_reviewed.json
```

submit 内部自动执行 write + TM 积累 + 重新 parse + 生成下一批 JSON。提交后回到 Step 1。

### 完成

全部批次完成后，`batch.py submit` 自动清理状态文件。告知用户完成，输出文件位置。


## 错误处理

- task 返回格式错误 → 重试当前 batch
- `batch.py` 命令报错 → 暂停，展示错误信息
- 用户随时可说"进度"查看 `batch.py status`

## 注意事项

- 翻译和校对 task 各自独立上下文
- 所有中间文件按源文件保留在 `batch_translate/exports/<stem>/`
- 源文件不受影响，init 会复制到 `batch_translate/data/<stem>/_working_*`
- 支持格式：mqxliff / docx / xlsx / xlsm / txt / csv / tsv

## Shell 命令规范

- **多行 Python**：使用 heredoc 语法 `python << 'PYEOF' ... PYEOF`（单引号阻止 shell 变量展开）
- **含中文输出时**：将 stdout 重定向到文件后用 read_file 读取，避免 GBK/终端编码错误
  ```bash
  python -c "..." > output.txt 2>&1
  ```
- **路径含特殊字符**（空格、单引号等）：优先用 python `os.path.join()` 拼接路径，而非 bash 字符串拼接
- **检查文件存在**：优先用 glob 工具，而非 `ls path/with/special'chars`
- **pip**：用 `python -m pip install` 而非 `pip install`，避免 PATH 问题
