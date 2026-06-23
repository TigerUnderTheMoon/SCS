# ESPR投稿改造计划

## TL;DR

> **目标**: 将现有MDPI Sustainability格式论文改造为ESPR(Springer)格式并完成投稿。
>
> **核心改造**: 引用格式(数字→作者-年份)、删除MDPI元素、scope适配(弱化green economy)、作者元数据填写、硬约束检查(自引≤5/无et al.)。
>
> **交付物**: 改造后的docx手稿、ESPR版cover letter、投稿系统提交确认。
>
> **预计工作量**: 中等(2-3天)
> **并行执行**: YES - 3个执行波次
> **关键路径**: 作者元数据 → 引用改造 → scope适配 → 最终QA → 投稿

---

## Context

### 原始请求
用户决定将现有为MDPI Sustainability准备的论文改造并投稿至Environmental Science and Pollution Research (ESPR)。

### 论文现状
- **当前手稿**: `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`
- **当前格式**: MDPI Sustainability专属格式(结构化摘要、Highlights、Graphical Abstract、数字引用)
- **实证内容**: 完整TWFE+IV+稳健性+中介+调节+异质性，已通过结构QA
- **阻塞项**: 作者元数据未填写(姓名/单位/资助/利益冲突等)
- **数据**: 284城市2006-2024面板，主样本为2006-2021观测值

### 目标期刊状态
- **ESPR**: Springer hybrid期刊，当前EI收录(SCIE 2024.10已剔除)
- **预警状态**: 2024版中科院预警(引用操纵/论文工厂)，2025版已移出
- **2025年起**: 停收"Green Economy"主题
- **费用**: 订阅制0元(投稿时不选Open Access)
- **投稿系统**: Editorial Manager (http://www.editorialmanager.com/espr/)

### 已确认约束
- 必须是EI/SCI/SSCI/SCIE之一 → EI符合
- 难度容易优先 → ESPR接收率~75%
- 速度越快越好 → 首决16天
- 价格≤1万 → 订阅制0元
- 单位认定: 需用户确认北信科技大学是否接受EI

---

## Work Objectives

### Core Objective
将Sustainability格式docx改造为符合ESPR投稿要求的Word文档，填写所有作者元数据，完成Editorial Manager系统投稿。

### Concrete Deliverables
1. 改造后的ESPR格式docx手稿
2. ESPR版cover letter (.docx)
3. 投稿系统提交确认截图/记录
4. `change_log.md`更新记录

### Definition of Done
- [ ] 手稿通过ESPR格式自检清单
- [ ] 引用格式为作者-年份(非数字)
- [ ] 自引ESPR文章≤5篇
- [ ] 参考文献无et al.(全部作者列出)
- [ ] 作者元数据完整
- [ ] Cover letter含preprint声明
- [ ] Editorial Manager系统提交成功

### Must Have
- 引用格式数字→作者-年份(33条全部改造)
- 删除MDPI特有元素(Highlights/Graphical Abstract/Article标签/结构化摘要)
- Abstract改为非结构化，扩至10-15行(~180-220词)
- 作者元数据完整填写
- Statements & Declarations放参考文献前
- Cover letter含preprint声明
- 硬约束: 自引ESPR≤5篇，无et al.

### Must NOT Have (Guardrails)
- 不修改实证结果/系数/表格数据
- 不重新跑Stata
- 不修改.dta数据文件
- 不选Open Access(避免APC)
- 不出现"green economy"措辞
- 不引用被剔除的ESPR文章作为SCIE

---

## Verification Strategy

### Test Decision
- **基础设施**: 无单元测试框架(论文改造任务)
- **验证方式**: Agent-Executed QA为主
  - 格式检查: Python脚本扫描docx结构
  - 引用检查: 统计ESPR自引数量、et al.使用情况
  - 渲染检查: PDF生成+视觉QA
  - 投稿验证: EM系统截图确认

### QA Policy
每任务必须包含Agent-Executed QA场景，证据保存至`.omo/evidence/`。

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 0 (Pre-scan - 5分钟预扫描， unblock Wave 1):
└── 任务0: 当前手稿预扫描(green economy / ESPR自引 / et al. / 字段函数)

Wave 1 (Start Immediately - 基础改造，可并行):
├── 任务1: 作者元数据填写 [quick]
├── 任务2: 引用格式全改造(数字→作者-年份) [unspecified-high]
├── 任务3: 删除MDPI元素 + Abstract改造 [quick]
└── 任务4: 硬约束检查(ESPR自引≤5 + 总自引≤5 + 无et al.) [quick]

Wave 2 (After Wave 1 - 内容适配):
├── 任务5: Scope适配(弱化green economy，强化environmental stressor) [deep]
├── 任务6: Statements & Declarations重组 [quick]
└── 任务7: Cover Letter改写(ESPR版 + preprint声明) [quick]

Wave 3 (After Wave 2 - 最终整合):
├── 任务8: 最终QA + PDF渲染 + 格式自检 [unspecified-high]
└── 任务9: Editorial Manager系统投稿 [quick]

Wave FINAL (After ALL tasks - 4并行审查):
├── F1: Plan compliance audit (oracle)
├── F2: 文档质量审查 (unspecified-high)
├── F3: 投稿前实机QA (unspecified-high)
└── F4: Scope fidelity check (deep)
-> 呈现结果 -> 获取用户确认

Critical Path: 任务0 → 任务1 → 任务2 → 任务5 → 任务8 → 任务9 → F1-F4 → user okay
Parallel Speedup: ~40% faster than sequential
Max Concurrent: 4 (Wave 1)
```

### Dependency Matrix

| Task | Depends On | Blocks |
|---|---|---|
| 0 (预扫描) | - | 2, 4, 5 |
| 1 (元数据) | - | 6, 7, 9 |
| 2 (引用改造) | 0 | 5, 8 |
| 3 (删MDPI元素) | - | 8 |
| 4 (硬约束检查) | 0 | 2, 8 |
| 5 (Scope适配) | 0, 2 | 8 |
| 6 (声明重组) | 1 | 8 |
| 7 (Cover Letter) | 1 | 9 |
| 8 (最终QA) | 2, 3, 4, 5, 6 | 9 |
| 9 (投稿) | 1, 7, 8 | F1-F4 |

### Agent Dispatch Summary

- **Wave 1**: 4 agents - T1→quick, T2→unspecified-high, T3→quick, T4→quick
- **Wave 2**: 3 agents - T5→deep, T6→quick, T7→quick
- **Wave 3**: 2 agents - T8→unspecified-high, T9→quick
- **FINAL**: 4 agents - F1→oracle, F2→unspecified-high, F3→unspecified-high, F4→deep

---

## TODOs

> Implementation + Test = ONE Task. Never separate.
> EVERY task MUST have: Recommended Agent Profile + Parallelization info + QA Scenarios.
> **A task WITHOUT QA Scenarios is INCOMPLETE.**


- [x] 0. 当前手稿预扫描(green economy / ESPR自引 / et al. / 字段函数)

  **What to do**:
  - [ ] 用python扫描当前docx全文，统计"green economy"(含大小写变体)出现次数及位置
  - [ ] 扫描33条引用，统计引用ESPR期刊的文章数量
  - [ ] 扫描33条引用，统计所有作者的总自引数(同一作者引用自己其他论文)
  - [ ] 扫描参考文献列表，确认是否已有"et al."或"..."
  - [ ] 扫描docx是否使用了Word字段函数(如交叉引用、自动目录等)
  - [ ] 将预扫描结果写入 `espr_prescan_report.json`

  **Must NOT do**:
  - 不修改任何内容，仅扫描报告

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: NO (必须在Wave 1前完成)
  - **Parallel Group**: Wave 0
  - **Blocks**: Tasks 2, 4, 5
  - **Blocked By**: None

  **References**:
  - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`

  **Acceptance Criteria**:
  - [ ] `espr_prescan_report.json`已生成
  - [ ] 报告含: green economy计数+位置, ESPR自引计数, 总自引计数, et al.计数, 字段函数计数

  **QA Scenarios**:
  ```
  Scenario: 预扫描完整性验证
    Tool: Bash (python)
    Preconditions: docx存在
    Steps:
      1. 运行预扫描脚本
      2. 确认espr_prescan_report.json生成
      3. 确认所有5个检查项均有结果
    Expected Result: 报告完整，零修改
    Evidence: .omo/evidence/task0-prescan-report.json
  ```

  **Commit**: NO

- [x] 1. 作者元数据填写与确认

  **What to do**:
  - 填写 `author_metadata_intake.md`：作者姓名、单位、邮箱、ORCID(如有)
  - 确认通讯作者(name/affiliation/email/phone/address)
  - 填写CRediT贡献声明(使用CRediT taxonomy)
  - 填写资助声明(funder名称+grant号，或"no external funding")
  - 填写致谢(或确认"Not applicable")
  - 填写利益冲突声明
  - 确认数据可用性声明措辞(代码公开+数据受限)
  - 确认AI使用声明(对齐Springer政策，非MDPI式)

  **Must NOT do**:
  - 不得虚构作者信息
  - 不得假设无利益冲突而不声明

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4)
  - **Blocks**: Tasks 6, 7, 9
  - **Blocked By**: None

  **References**:
  - `outputs/sustainability_restructure_20260603_deep/author_metadata_intake.md`
  - `AGENTS.md`
  - ESPR submission guidelines - Statements & Declarations要求

  **Acceptance Criteria**:
  - [ ] `author_metadata_intake.md`所有字段已填写
  - [ ] 所有声明措辞符合Springer/ESPR政策(非MDPI式)
  - [ ] 资助声明含具体grant号或明确"no external funding"

  **QA Scenarios**:
  ```
  Scenario: 元数据完整性检查
    Tool: Bash (python docx检查)
    Preconditions: author_metadata_intake.md已填写
    Steps:
      1. 读取author_metadata_intake.md，确认所有[Name]占位符已替换
      2. 确认至少1名通讯作者含完整联系信息
      3. 确认CRediT声明覆盖所有作者
    Expected Result: 零占位符残留，所有必填字段完整
    Evidence: .omo/evidence/task1-metadata-check.txt
  ```

  **Commit**: NO (待全部完成后统一提交)

- [x] 2. 引用格式全改造(数字编号→作者-年份)

  **What to do**:
  - 将正文所有 `[数字]` 引用改为 `(Author Year)` 格式
  - 将参考文献列表从编号格式改为字母序作者-年份格式(SPBASIC样式)
  - 示例改造：
    - 原：`[5] Yang et al. find that...` → 改后：`(Yang et al. 2018)`
    - 原：`1. Seto, K.C.; Guneralp, B.; ...` → 改后：`Seto KC, Guneralp B, Hutyra LR (2012) Global forecasts... Proc Natl Acad Sci 109:... https://doi.org/...`
  - 33条引用全部改造
  - 注意：ESPR要求作者全列，无et al.(见任务4硬约束)

  **Must NOT do**:
  - 不得遗漏任何引用
  - 不得改变引用内容/结论
  - 不得引入新引用

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 3, 4)
  - **Blocks**: Tasks 5, 8
  - **Blocked By**: None

  **References**:
  - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`
  - ESPR submission guidelines - References格式要求(SPBASIC)
  - `outputs/sustainability_restructure_20260603_deep/reference_doi_audit.csv`

  **Acceptance Criteria**:
  - [ ] 正文中零个 `[数字]` 引用残留
  - [ ] 参考文献列表按首作者姓氏字母序排列
  - [ ] 每条引用含完整DOI链接
  - [ ] 格式符合SPBASIC

  **QA Scenarios**:
  ```
  Scenario: 引用格式完整性验证
    Tool: Bash (python正则扫描)
    Preconditions: docx已改造
    Steps:
      1. 扫描docx全文，确认无数字引用残留
      2. 统计作者-年份引用数量 = 33
      3. 确认参考文献列表按字母序排列
      4. 抽样验证5条引用格式符合SPBASIC
    Expected Result: 零数字引用，33条作者-年份引用，字母序排列
    Evidence: .omo/evidence/task2-citation-check.txt
  ```

  **Commit**: NO


- [x] 3. 删除MDPI特有元素 + Abstract改造

  **What to do**:
  - **删除MDPI元素**：
    - [ ] 删除首段`Article`标签
    - [ ] 删除`Highlights`部分(5条bullet)
    - [ ] 删除或降级`Graphical Abstract`(ESPR不强制，可删)
    - [ ] 删除`Keywords`后的数字编号(如果有)
  - **Abstract改造**：
    - [ ] 将结构化摘要(Background/Methods/Results/Conclusions)改为非结构化一段式
    - [ ] 扩写至10-15行(约180-220词)，当前130词偏短
    - [ ] 内容保持：研究背景、方法(SCCD+284城市+TWFE)、核心发现(倒U型+拐点0.522)、结论
  - **其他格式**：
    - [ ] 确认使用10pt Times Roman字体
    - [ ] 确认自动页码(非手动)
    - [ ] 移除Word字段函数(如果有)
    - [ ] 确认使用Word表格功能(非Excel截图)

  **Must NOT do**:
  - 不得删除实证内容/数据
  - 不得改变研究发现

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 2, 4)
  - **Blocks**: Task 8
  - **Blocked By**: None

  **References**:
  - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`
  - ESPR submission guidelines - Text Formatting要求

  **Acceptance Criteria**:
  - [ ] 零个MDPI特有元素残留
  - [ ] Abstract为非结构化一段式，180-220词
  - [ ] 字体为10pt Times Roman
  - [ ] 使用自动页码

  **QA Scenarios**:
  ```
  Scenario: MDPI元素清除验证
    Tool: Bash (python docx扫描)
    Preconditions: 改造完成
    Steps:
      1. 全文搜索"Article"(首段标签) → 应不存在
      2. 搜索"Highlights" → 应不存在
      3. 搜索"Graphical Abstract" → 应不存在或已删除
      4. 搜索"Background:""Methods:""Results:""Conclusions:"(摘要结构化标记) → 应不存在
    Expected Result: 零MDPI元素残留
    Evidence: .omo/evidence/task3-mdpi-cleanup.txt

  Scenario: Abstract长度验证
    Tool: Bash (python)
    Preconditions: Abstract已改写
    Steps:
      1. 提取Abstract文本
      2. 统计词数(word count)
      3. 统计行数(按10pt Times Roman估算)
    Expected Result: 180-220词，10-15行
    Evidence: .omo/evidence/task3-abstract-length.txt
  ```

  **Commit**: NO

- [x] 4. 硬约束检查(自引≤5 + 无et al. + 其他)

  **What to do**:
  - **ESPR自引检查**：
    - [ ] 扫描33条引用，统计引用ESPR期刊的文章数量
    - [ ] 若>5篇，必须删减至≤5篇
    - [ ] 注意：引用Sustainability(MDPI)不算ESPR自引
  - **作者全列检查**：
    - [ ] 扫描参考文献列表，确认无"et al."
    - [ ] 确认无"..."省略号替代作者
    - [ ] 长作者列表(>3人)需全列(ESPR新规则)
  - **其他硬约束**：
    - [ ] 确认标题无缩写(SCCD如存在需评估)
    - [ ] 确认关键词6-8个(当前6个，可保持)
    - [ ] 确认标题含地域标注("Chinese Cities"已有)
    - [ ] 确认无Wikipedia引用

  **Must NOT do**:
  - 不得为凑自引数而添加ESPR引用
  - 不得用et al.或省略号规避长作者列表

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 2, 3)
  - **Blocks**: Tasks 2, 8
  - **Blocked By**: None (但任务4结果会影响任务2的引用改造)

  **References**:
  - ESPR submission guidelines - References具体要求
  - `outputs/sustainability_restructure_20260603_deep/reference_doi_audit.csv`

  **Acceptance Criteria**:
  - [ ] ESPR期刊自引数量 ≤ 5
  - [ ] 所有作者总自引数量 ≤ 5(同一作者引用自己其他论文)
  - [ ] 零个"et al."或"..."在参考文献中
  - [ ] 零个Wikipedia引用

  **QA Scenarios**:
  ```
  Scenario: 自引与et al.硬约束验证
    Tool: Bash (python正则扫描)
    Preconditions: 引用改造完成
    Steps:
      1. 扫描参考文献，统计含"Environ Sci Pollut Res"或"ESPR"或ISSN 0944-1344的条目
      2. 确认计数 ≤ 5
      3. 扫描全文"et al"和"..." → 应为零
      4. 扫描Wikipedia引用 → 应为零
    Expected Result: ESPR自引≤5，零et al.，零Wikipedia
    Evidence: .omo/evidence/task4-hard-constraints.txt
  ```

  **Commit**: NO


- [x] 5. Scope适配(弱化green economy，强化environmental stressor)

  **What to do**:
  - **引言改造**：
    - [ ] 弱化"smart city policy evaluation"叙事
    - [ ] 强化"urban environmental stressor"和"carbon emissions as pollution"定位
    - [ ] 明确将碳排放定位为环境压力因子(environmental stressor)，契合ESPR核心scope
    - [ ] 在2.1节首段增加ESPR scope呼应语：将smart-city construction与urban environmental management衔接
  - **讨论改造**：
    - [ ] Discussion 5.1-5.3中替换或弱化"green economy"措辞(ESPR 2025年起禁收)
    - [ ] 用"sustainable urban transition""low-carbon governance""environmental management"替代
    - [x] 5.2节"Governance Sequencing for Sustainability"可保留但需避免green economy字眼
  - **结论改造**：
    - [ ] Conclusions中确保无"green economy"，用"sustainable urban management""environmental governance"替代
  - **关键词检查**：
    - [ ] 确认6个关键词无"green economy"

  **Must NOT do**:
  - 不得改变实证结果/系数
  - 不得添加新数据
  - 不得虚构新发现

  **Recommended Agent Profile**:
  - **Category**: `deep`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: NO (依赖任务2引用改造完成，避免引用编号混乱)
  - **Parallel Group**: Wave 2 (with Tasks 6, 7)
  - **Blocks**: Task 8
  - **Blocked By**: Task 2

  **References**:
  - `outputs/sustainability_restructure_20260603_deep/manuscript_sustainability_deep.docx`
  - ESPR aims and scope - 官方scope定义
  - `issues_remaining.md` - 历史改造记录

  **Acceptance Criteria**:
  - [ ] 全文零个"green economy"(含大小写变体)
  - [ ] 引言明确呼应ESPR的"environmental stressor"和"sustainable urban landscape"scope
  - [ ] 实证结论不变

  **QA Scenarios**:
  ```
  Scenario: Green economy清零验证
    Tool: Bash (python全文扫描，不区分大小写)
    Preconditions: scope适配完成
    Steps:
      1. 全文搜索"green economy" → 计数应为0
      2. 搜索"green-economy" → 0
      3. 搜索相关变体 → 0
    Expected Result: 绝对零残留
    Evidence: .omo/evidence/task5-green-economy-zero.txt

  Scenario: Scope呼应验证
    Tool: Bash (python)
    Preconditions: 改造完成
    Steps:
      1. 提取Introduction首段和末段
      2. 确认含"environmental"或"pollution"或"emissions"关键词
      3. 确认提及"urban landscape"或"sustainable urban"与ESPR scope对齐
    Expected Result: 引言明确将研究定位为环境科学问题
    Evidence: .omo/evidence/task5-scope-alignment.txt
  ```

  **Commit**: NO

- [x] 6. Statements & Declarations重组

  **What to do**:
  - **位置调整**：
    - [ ] 将所有声明(Author Contributions/Funding/Acknowledgments/Conflicts/Data Availability/AI/Ethics)统一置于**References之前**
    - [ ] 添加大标题"Statements and Declarations"
    - [ ] 当前MDPI式分散声明改为ESPR要求的统一节
  - **内容调整**：
    - [ ] Author Contributions: 使用CRediT taxonomy，含具体作者姓名
    - [ ] Funding: 含funder全称+grant号，或明确"no funds..."
    - [ ] Competing Interests: 财务+非财务利益披露，或"no relevant..."
    - [ ] Data Availability: 代码公开+原始数据受限(因源许可)
    - [ ] AI声明: 对齐Springer政策(非MDPI式)，声明AI仅用于语言润色/结构编辑
    - [ ] Ethics: "Not applicable"(城市统计数据，无人/动物)
  - **删除元素**：
    - [ ] 删除MDPI式的"Institutional Review Board Statement"和"Informed Consent Statement"独立标题(合并到Ethics一句)

  **Must NOT do**:
  - 不得虚构资助信息
  - 不得隐瞒利益冲突

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 5, 7)
  - **Blocks**: Task 8
  - **Blocked By**: Task 1 (需作者元数据)

  **References**:
  - ESPR submission guidelines - Statements & Declarations
  - `outputs/sustainability_restructure_20260603_deep/author_metadata_intake.md`

  **Acceptance Criteria**:
  - [ ] 所有声明在References之前，统一在"Statements and Declarations"标题下
  - [ ] 无MDPI式分散声明标题残留
  - [ ] 所有声明含具体作者信息(非占位符)

  **QA Scenarios**:
  ```
  Scenario: 声明位置验证
    Tool: Bash (python docx结构扫描)
    Preconditions: 重组完成
    Steps:
      1. 定位"References"标题位置
      2. 定位"Statements and Declarations"标题位置
      3. 确认Statements在References之前
      4. 确认无"Institutional Review Board Statement"独立标题
    Expected Result: Statements在References前，结构符合ESPR要求
    Evidence: .omo/evidence/task6-statements-position.txt
  ```

  **Commit**: NO

- [x] 7. Cover Letter改写(ESPR版 + preprint声明)

  **What to do**:
  - 基于现有 `submission_cover_letter_draft.md` 改写
  - **收件人**：改为 `Editors, Environmental Science and Pollution Research`
  - **内容调整**：
    - [ ] 强调论文与ESPR scope的匹配(Sustainable Urban and Rural Landscape)
    - [ ] 说明不涉及green economy(规避2025禁令)
    - [ ] 填入作者元数据(通讯作者信息)
    - [ ] 填入CRediT贡献简述
    - [ ] 填入资助/利益冲突声明
  - **preprint声明**(ESPR强制)：
    - [ ] 添加一句：`I have [not] submitted my manuscript to a preprint server before submitting it to Environmental Science and Pollution Research`
    - [ ] 二选一，根据实际情况选择
  - **原始性声明**：
    - [ ] 声明稿件原创、未在他处考虑发表、所有作者已批准
  - **保存为docx**：
    - [ ] 生成 `cover_letter_espr.docx`

  **Must NOT do**:
  - 不得使用MDPI式措辞
  - 不得遗漏preprint声明

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 5, 6)
  - **Blocks**: Task 9
  - **Blocked By**: Task 1 (需作者元数据)

  **References**:
  - `outputs/sustainability_restructure_20260603_deep/submission_cover_letter_draft.md`
  - ESPR submission guidelines - General Information(preprint声明要求)

  **Acceptance Criteria**:
  - [ ] Cover letter收件人为ESPR编辑
  - [ ] 含preprint声明(二选一)
  - [ ] 含原始性/未在他处发表声明
  - [ ] 通讯作者信息完整
  - [ ] 保存为docx格式

  **QA Scenarios**:
  ```
  Scenario: Cover letter完整性验证
    Tool: Bash (python docx扫描)
    Preconditions: cover_letter_espr.docx已生成
    Steps:
      1. 确认含"Environmental Science and Pollution Research"(收件人)
      2. 确认含"preprint server"(preprint声明)
      3. 确认含"not under consideration"或类似原始性声明
      4. 确认通讯作者姓名非占位符
    Expected Result: 所有要素齐全
    Evidence: .omo/evidence/task7-cover-letter-check.txt
  ```

  **Commit**: NO


- [x] 8. 最终QA + PDF渲染 + 格式自检

  **What to do**:
  - **整合检查**：
    - [ ] 确认任务1-7全部完成
    - [ ] docx所有改造点已落实
  - **格式自检**：
    - [ ] 字体：10pt Times Roman
    - [ ] 页码：自动页码
    - [ ] 表格：Word表格功能
    - [ ] 公式：Word公式编辑器
    - [ ] 标题层级：≤3级
    - [ ] 脚注(非尾注)
  - **引用终检**：
    - [ ] 零数字引用
    - [ ] 33条作者-年份引用
    - [ ] 字母序排列
    - [ ] ESPR自引≤5
    - [ ] 无et al.
  - **内容终检**：
    - [ ] 无"green economy"
    - [ ] Abstract 180-220词非结构化
    - [ ] Statements在References前
    - [ ] 无MDPI元素残留
  - **PDF生成**：
    - [ ] 用Word导出PDF
    - [ ] 渲染为PNG contact sheet
    - [ ] 视觉检查无溢出/截断
  - **生成检查报告**：
    - [ ] 创建 `outputs/espr_submission/` 目录(如不存在)
    - [ ] 写入 `outputs/espr_submission/espr_final_checks.json`
    - [ ] JSON格式参考 `sustainability_deep_checks.json`，字段包括:
      - `passed`: bool (全部通过=true)
      - `submission_ready`: bool (作者元数据完整=true)
      - `submission_status`: string (如 "ESPR_FORMAT_QA_PASSED")
      - `word_count_main_text_no_references`: int
      - `abstract_word_count`: int
      - `abstract_is_structured`: bool (应为false)
      - `keywords_count`: int
      - `citation_format`: string (应为 "author-year")
      - `numeric_citations_found`: int (应为0)
      - `espr_self_citations`: int (应≤5)
      - `total_self_citations`: int (应≤5)
      - `et_al_found`: int (应为0)
      - `green_economy_found`: int (应为0)
      - `mdpi_elements_found`: list (应为空)
      - `statements_before_references`: bool (应为true)
      - `long_sentences_ge_35_words`: int
      - `generator_trace_terms_found`: list (应为空)
      - `internal_process_terms_found_main_text`: list (应为空)

  **Must NOT do**:
  - 不得修改实证结果
  - 不得跳过任何检查项

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high`
  - **Skills**: [`safe-edit`]

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 3 (with Task 9, 但9依赖8)
  - **Blocks**: Task 9, F1-F4
  - **Blocked By**: Tasks 2, 3, 4, 5, 6

  **References**:
  - 所有前置任务产出
  - ESPR submission guidelines - 全部要求
  - `outputs/sustainability_restructure_20260603_deep/sustainability_deep_checks.json`(参考格式)

  **Acceptance Criteria**:
  - [ ] `espr_final_checks.json`报告 `passed: true`
  - [ ] PDF成功生成，32+页
  - [ ] 视觉QA通过(contact sheet检查)
  - [ ] 所有ESPR硬约束通过

  **QA Scenarios**:
  ```
  Scenario: 最终格式全面验证
    Tool: Bash (python docx全面扫描)
    Preconditions: 所有改造完成
    Steps:
      1. 执行格式自检脚本(字体/页码/表格/公式/标题层级)
      2. 执行引用终检(数字引用/字母序/自引/et al.)
      3. 执行内容终检(green economy/abstract结构/声明位置/MDPI元素)
      4. 生成espr_final_checks.json
    Expected Result: 所有检查项passed: true
    Evidence: .omo/evidence/task8-final-checks.json

  Scenario: PDF渲染验证
    Tool: Bash (Word导出PDF + pdftoppm渲染)
    Preconditions: docx最终版完成
    Steps:
      1. 用Word COM导出PDF
      2. 用pdftoppm渲染为PNG contact sheet
      3. 视觉检查首页/摘要页/表格页/参考文献页/声明页
    Expected Result: PDF正常生成，无溢出/截断
    Evidence: .omo/evidence/task8-pdf-render/
  ```

  **Commit**: YES
  - Message: `docs(espr): adapt manuscript for ESPR submission`
  - Files: `outputs/espr_submission/manuscript_espr.docx`, `outputs/espr_submission/manuscript_espr.pdf`, `outputs/espr_submission/espr_final_checks.json`, `cover_letter_espr.docx`
  - Pre-commit: python格式自检脚本

- [~] 9. Editorial Manager系统投稿

  **What to do**:
  - **注册/登录**：
    - [ ] 访问 http://www.editorialmanager.com/espr/
    - [ ] 注册账号(如未注册)或登录
  - **新建投稿**：
    - [ ] 选择文章类型(Research Article)
    - [ ] 填入标题(避免缩写)
    - [ ] 填入摘要(非结构化)
    - [ ] 填入关键词(6-8个)
  - **上传文件**：
    - [ ] 上传主手稿docx
    - [ ] 上传cover letter docx
    - [ ] 上传图表(如需单独上传)
    - [ ] 上传补充材料(ESPR要求命名为ESM_1, ESM_2等):
      - `ESM_1.pdf`: 主分析表格(Table S1-S2: VIF诊断+固定效应联合检验)
      - `ESM_2.pdf`: 敏感性分析表格(全样本2006-2024结果)
      - `ESM_3.pdf`: 补充图表(如需要)
      - `README.txt`: 补充材料说明(复制supplementary_materials/README.md内容)
      - 注意: 每个文件需含文章标题、期刊名、作者名、通讯作者邮箱
  - **选择出版模式**：
    - [ ] **选择Subscription(订阅制)**，不选Open Access(避免APC)
  - **填写声明**：
    - [ ] 确认preprint声明
    - [ ] 确认利益冲突声明
    - [ ] 确认资助声明
  - **提交**：
    - [ ] 确认生成的PDF预览
    - [ ] 提交
    - [ ] 记录投稿号(Manuscript ID)
    - [ ] 截图保存投稿确认

  **Must NOT do**:
  - 不得选Open Access(会产生APC ~2万元)
  - 不得遗漏preprint声明
  - 不得上传未最终版的文件

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 3 (after Task 8)
  - **Blocks**: F1-F4
  - **Blocked By**: Tasks 1, 7, 8

  **References**:
  - ESPR Editorial Manager: http://www.editorialmanager.com/espr/
  - ESPR submission guidelines

  **Acceptance Criteria**:
  - [ ] 投稿成功，获得Manuscript ID
  - [ ] 投稿确认截图保存
  - [ ] 选择Subscription模式(非Open Access)
  - [ ] 所有文件已上传

  **QA Scenarios**:
  ```
  Scenario: 投稿确认验证
    Tool: 人工+截图
    Preconditions: EM系统投稿完成
    Steps:
      1. 确认收到投稿确认邮件
      2. 截图EM系统中的投稿状态页面
      3. 记录Manuscript ID
      4. 确认状态为"Submitted to Journal"或"With Editor"
    Expected Result: 获得Manuscript ID，状态正常
    Evidence: .omo/evidence/task9-submission-confirmation.png
  ```

  **Commit**: YES
  - Message: `docs(espr): record submission confirmation`
  - Files: `.omo/evidence/task9-submission-confirmation.png`, `change_log.md`
  - Pre-commit: none


## Final Verification Wave (MANDATORY — after ALL implementation tasks)

> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.

- [x] F1. **Plan Compliance Audit** — `oracle`
  Read the plan end-to-end. For each "Must Have": verify implementation exists (read docx, run python check, inspect file). For each "Must NOT Have": search docx for forbidden patterns — reject with location if found. Check evidence files exist in .omo/evidence/. Compare deliverables against plan.
  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [x] F2. **文档质量审查** — `unspecified-high`
  Open the final docx and PDF. Review for: format consistency, citation format correctness, abstract quality, statement completeness, green economy residue, MDPI element residue, table/figure integrity. Check AI slop: excessive hedging, generic phrasing, tortured phrases.
  Output: `Format [PASS/FAIL] | Citations [PASS/FAIL] | Content [PASS/FAIL] | Statements [N/N] | VERDICT`

- [x] F3. **投稿前实机QA** — `unspecified-high`
  Start from the final docx. Execute EVERY QA scenario from EVERY task — follow exact steps, capture evidence. Test cross-task integration (引用格式+scope适配+声明位置同时正确). Test edge cases: long author lists, self-citation count, DOI completeness. Save to `.omo/evidence/final-qa/`.
  Output: `Scenarios [N/N pass] | Integration [N/N] | Edge Cases [N tested] | VERDICT`

- [x] F4. **Scope Fidelity Check** — `deep`
  For each task: read "What to do", read actual diff (git log/diff or docx comparison). Verify 1:1 — everything in spec was built (no missing), nothing beyond spec was built (no creep). Check "Must NOT do" compliance. Detect cross-task contamination: Task N touching Task M's files. Flag unaccounted changes.
  Output: `Tasks [N/N compliant] | Contamination [CLEAN/N issues] | Unaccounted [CLEAN/N files] | VERDICT`

---

## Commit Strategy

- **Task 8**: `docs(espr): adapt manuscript for ESPR submission` - manuscript_espr.docx, manuscript_espr.pdf, espr_final_checks.json, cover_letter_espr.docx
- **Task 9**: `docs(espr): record submission confirmation` - task9-submission-confirmation.png, change_log.md

---

## Success Criteria

### Verification Commands
```bash
# 格式自检
python espr_format_check.py  # Expected: all checks passed

# 引用检查
python espr_citation_check.py  # Expected: 0 numeric citations, 33 author-year, ESPR self-citation ≤ 5

# 内容检查
python espr_content_check.py  # Expected: 0 "green economy", abstract 180-220 words, statements before references
```

### Final Checklist
- [ ] 所有"Must Have"项已落实
- [ ] 所有"Must NOT Have"项已排除
- [ ] 引用格式为作者-年份(零数字引用)
- [ ] ESPR自引≤5篇
- [ ] 参考文献无et al.(全部作者列出)
- [ ] 全文无"green economy"
- [ ] Abstract为非结构化，180-220词
- [ ] Statements & Declarations在References前
- [ ] 无MDPI元素残留
- [ ] Cover letter含preprint声明
- [ ] 作者元数据完整填写
- [ ] EM系统投稿成功，获得Manuscript ID
- [ ] 选择Subscription模式(非Open Access)

---

## 投稿后跟踪建议

1. **审稿状态**：登录EM系统定期查看状态变化(Submitted → With Editor → Under Review → Decision)
2. **预期时间线**：首决约16天，整体3.9个月
3. **revision准备**：若收到修改意见，重复Wave 2-3流程
4. **风险监控**：关注ESPR是否被EI剔除(每季度查Engineering Village)
5. **备选方案**：若ESPR投稿期间被EI剔除，立即转投EBE或Urban Climate




