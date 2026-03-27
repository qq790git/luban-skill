---
name: luban-skill
description: 基于 DataBale/Luban 文档规则生成并校验真实 Excel（.xlsx）文件。凡是涉及 Excel 自动生成、按字段规则出表、模板化制表、批量填充、字段映射、容器/自定义类型、多态结构、校验器（ref/range/path/size/set 等）、或要求参考 datable.cn / Luban 文档规范的任务，都应优先使用本 skill；即使用户未明确点名该 skill，只要目标是“把规则落地成可交付 xlsx + 可验证校验结果”，也必须触发。
---

# DataBale Excel Generator

## 1. 目标
把"规则描述/文档约束/原始数据"稳定落地为：
1) 真实 `.xlsx` 文件（可被后续 Luban 流程消费）   
2) 对生成 Excel 的完整校验报告
3) 可执行的数据生成指引（`gen.bat` / `gen.sh`）

---

## 2. 触发条件（命中任一即触发）
- 生成、改造、批量生产 Excel。
- 将文本规则/数据字典/业务约束转成表格。
- 要求字段类型、必填、枚举、范围、唯一性等严格校验。
- 提到 DataBale / datable.cn / Luban 规范。
- 要求“先按规范出表，再生成配置数据/代码”。

---

## 3. 输入最小集（不足时最小化追问）
- 仅追问"字段定义 + 输出路径"两项关键缺口。
- 复合数据类型（如数组、容器、自定义类型），必须在 `__beans__.xlsx` 中声明。并且优先采用流式格式（紧凑格式），多个单元格，按顺序读取
- 规则来源
  - **`references/type-system.md`** - 完整类型系统和语法
  - **`references/schema.md`** - Schema 定义详解（XML/Excel）
  - **`references/luban-conf.md`** - luban.conf 完整配置参考
  - **`references/validators.md`** - 所有校验器类型和用法
  - **`references/excel-format.md`** - Excel 数据格式详解
  - **`references/command-line.md`** - 命令行参数完整列表
  - **`references/runtime.md`** - 运行时加载和类型映射


若信息不完整：

---

## 4. 执行流程（必须按顺序）
1. **读取规则源**：先读用户给的 DataBale/Luban 规则与补充约束。  
2. **需求归一化**：整理字段清单（列名、类型、必填、约束、示例值、分组）。  
3. **工程检查**：复用用户工程；无则按 `MiniTemplate` 风格补齐骨架。  
4. **补全表结构**: 在 Defines/ 目录或 __tables__.xlsx 声明表结构
4. **生成 Excel**：写入 `Datas/`，确保元信息行与类型声明正确。  
5. **多层校验**：结构/类型/必填/约束/复合类型/格式规则全量校验。  
6. **生成指引**：给出 `gen.bat`/`gen.sh` 的执行方式与产物目录说明。  
7. **结果输出**：按统一模板输出“生成结果 + 校验清单 + 假设与风险”。

---

## 5. 工程与目录规则

### 5.1 标准工程（优先复用）
- 优先 `MiniTemplate` 风格。
- 关键目录：
  - `MiniTemplate/Datas/`
  - `MiniTemplate/output/`
  - `MiniTemplate/luban.conf`
  - `MiniTemplate/gen.bat`（Windows）或 `MiniTemplate/gen.sh`（Mac/Linux）
- 若用户已有结构：禁止强制重建，仅补齐缺失项。

---

## 6. Excel 基础结构规则（必选）

### 6.1 元信息行语义
数据表至少应包含（顺序可变但语义完整）：
- 字段名行：首列 `##var`（或同语义字段行）
- 类型行：首列 `##type`
- 分组行（可选）：首列常见 `##group`，值为 `c` / `s` / `c,s` / 空
- 注释行（可选）：`##comment` 或其他 `##<name>`
- 数据行：紧随表头定义后

### 6.2 文件与 sheet 识别
- 支持 `xls/xlsx/xlm/xlmx/csv`，优先输出 `.xlsx`。
- 若未指定 sheet，默认读取全部 sheet。
- 若 `A1` 不以 `##` 开头，该 sheet 视为非数据 sheet（忽略）。

### 6.3 注释规则
- 字段名为空或以 `#` 开头的列：注释列（忽略）。
- 数据行第一列以 `##` 开头：注释行（忽略）。

### 6.4 声明文件
- 若不存在 `Datas/__tables__.xlsx`，需创建并登记新增表。
- 若不存在 `Datas/__beans__.xlsx` 或 `Datas/__enums__.xlsx`，需创建并保持与模板一致。

### 6.5 三张系统表固定模板（强约束）
- `.claude/skills/luban-skill/template` 目录下的 `__beans__.xlsx`、`__enums__.xlsx`、`__tables__.xlsx` 是**唯一基准模板**。
- 生成或维护这三张系统表时，必须以该目录中的同名模板为蓝本拷贝后再填充。
- **禁止修改模板表头结构**：包括但不限于列名、列顺序、元信息行（如 `##var`、`##type`、`##group`、`##comment`）及其语义。
- 仅允许在模板既有表头之下填入或更新数据行；不得新增、删除、重命名、重排表头列。
- 若业务数据与模板表头不兼容，必须先在输出中报告冲突并停止直接写入，不得“自动改表头”以适配数据。

### 6.6 新生成表格样式与固定（必做）
为保证交付一致性，所有新生成的 `.xlsx` 数据表（至少包含 `Datas/__beans__.xlsx`、`Datas/__enums__.xlsx`、`Datas/__tables__.xlsx` 及业务表）必须执行以下样式规则：

1. 表头着色（元信息头）
- `##var` 行：填充色 `#1F4E78`，字体白色、加粗。
- `##type` 行：填充色 `#2E75B6`，字体白色、加粗。
- `##group` 行：填充色 `#BDD7EE`，字体黑色、加粗。
- 其他 `##` 注释头行（如 `##comment`）：填充色 `#D9E1F2`，字体黑色。

2. 冻结与筛选
- 冻结窗格必须设置到“最后一行表头的下一行”（即数据首行），保证滚动时表头固定可见。
- 必须开启自动筛选，范围覆盖完整表头与数据区域。

3. 列宽与可读性
- 按列内容自动估算列宽，建议最小宽度 `12`，避免标题与数据被截断。
- 表头文字水平居中、垂直居中；数据区默认左对齐。

4. 一致性要求
- 同一批次生成的所有表必须使用同一套颜色与冻结策略，不允许单表特例漂移。
- 若用户未显式要求关闭样式，默认必须启用上述样式与固定策略。

---


## 7. 生成与命令规则
- 按照`references/command-reference.md` 生成 `gen.bat`/`gen.sh` 脚本。
- `luban.conf` 存在性检查，不存在则需要按照`references/luban-conf.md`创建配置文件。

### 7.1 `__tables__.xlsx` 列对齐规则（新增，强约束）
- `__tables__.xlsx` 模板首列为 tag/注释列，**数据行必须保留 A 列占位为空**，业务字段从 B 列开始写。
- 字段顺序必须严格对应模板：`full_name`(B) / `value_type`(C) / `read_schema_from_file`(D) / `input`(E) / `index`(F) / `mode`(G) / `group`(H) / `comment`(I) / `tags`(J) / `output`(K)。
- 禁止将 `input` 写入 `read_schema_from_file` 列；若出现 `xxx.xlsx 不是 bool 类型`，应优先检查是否发生列偏移。

### 7.2 Schema 来源去重规则（新增，强约束）
- 当 `value_type` 已在 `Datas/__beans__.xlsx` 中显式定义时，`__tables__.xlsx` 中对应记录的 `read_schema_from_file` 必须为 `false`。
- 仅在确有需要从业务数据表头推断结构时，才允许将 `read_schema_from_file` 设为 `true`。
- 若出现 `type:'xxx' duplicate`，应优先检查是否同时启用了 bean 显式定义与 `read_schema_from_file=true`。

### 7.3 幂等生成规则（新增，强约束）
- 每次生成前必须先用 `.claude/skills/luban-skill/template` 下三张模板重置 `Datas/__beans__.xlsx`、`Datas/__enums__.xlsx`、`Datas/__tables__.xlsx`，再写入数据。
- 禁止在旧文件上直接追加而不清理，避免重复定义累积导致二次执行失败。

---

## 8. 统一校验框架（必须输出统计）
- 校验器：`references/validators.md`

- 校验结果：`references/excel-format.md`

### 8.1 `gen.bat` 自测闭环（新增，必做）
- 修改任意 Excel/配置/脚本后，必须实际执行一次 `gen.bat`（或 `gen.sh`）做端到端校验。
- 若失败，必须进入“修复 -> 重新生成 -> 再次运行 `gen.bat`”循环，直到校验与导出均通过。
- 输出结果中必须包含：是否通过、自测命令、关键日志摘要、产物路径。
---



