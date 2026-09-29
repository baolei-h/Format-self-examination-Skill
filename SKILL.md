---
name: bioactive-surface-book-auditor
description: >-
  Trigger on requests such as 请调用 bioactive-surface-book-auditor 完成第X章自查, 自查第X章, or 检查第X章是否符合书稿规范, and equivalent requests. Supports explicit safe repair and focused reference, figure/table, format, recent-progress, monograph-logic, and delivery checks. If a chapter number or title is given without a path, search the current workspace and use the most likely manuscript. Default is audit-only; never edit the original unless repair is explicitly requested.
---

# 《生物活性表面可控构建技术》章节自查

## Simple Invocation / 一句话调用

普通作者只需输入以下任一请求，无需知道脚本、Python 命令、文件路径、参数或输出目录：

- `请调用 bioactive-surface-book-auditor 完成第5章自查`
- `请调用 bioactive-surface-book-auditor 自查第5章`
- `调用 bioactive-surface-book-auditor 检查第5章`
- `请用 bioactive-surface-book-auditor 检查第5章是否符合书稿规范`
- `请调用 bioactive-surface-book-auditor 完成“角膜与眼表修复材料的生物活性界面”章节自查`

以上均默认为 **audit-only**。收到完整自查请求后，不再询问是否检查格式、图片、文献、近期进展或报告格式。自动执行：

1. 在当前 workspace 中寻找章节文件；若用户上传了 Word/Markdown/TXT，直接使用该文件。
2. 读取本 Skill 的 `references/book-writing-standard.md` 与 `references/audit-rubric.md`。
3. 运行适用的确定性 DOCX 或文本检查；检查图表、公式、单位、统计、标题、文献和年份数据。
4. 模型必须继续通读全文（长文分段但覆盖全文），完成专著叙事、学位论文整合、重复内容、机制证据、科学表述强度、近年主题和版权/科研诚信风险审核。脚本运行完成不等于自查完成。
5. 对 DOCX 图片进行视觉检查：比例尺、A/B/C 面板标记与位置、坐标轴/单位、字体、组别标签、颜色是否为唯一编码、亮度/对比度一致性、非等比拉伸、网页/聊天截图风险及图像完整性线索。若当前环境无法查看嵌入图片或原始图像，报告必须写 `△ 需要视觉检查`，不能标记为通过。
6. 将确定性发现和 AI 全文审核结果合并，生成稿件同目录下 `自查结果/<章节>_自查报告.md`、`自查结果/<章节>_自查报告.docx`，并可保留自动检查 JSON。Word 生成失败时至少保留 Markdown 并说明原因。
7. 聊天中只返回章节、结论、P0/P1/P2/P3 数量、最多五项优先问题、报告位置及“原始稿件未修改”。不要粘贴整篇报告。

自然语言发现与选择规则：

- 搜索 workspace，而非一开始询问路径。DOCX 优先，其次 Markdown，再 TXT。支持 `第5章`、`Ch05`、`Ch5`、`Chapter05`、`Chapter5`、`5_`、`05_`及章节标题匹配。
- 多个候选按明确版本号最高、文件名日期最新、修改时间最新依次选择；过滤自查报告、自动检查结果、规范修订版、Word 临时锁文件、备份/旧版/归档/临时文件。
- 仍有两个无法判定的同优先级当前版本时，才询问用户选择。完全搜索不到后才说：“未在当前工作区找到第X章稿件，请上传Word或告诉我文件位置。”
- 用户提供完整交稿目录时，一并检查 `01_Chapter`、`02_Figures`、`03_FigureSources`、`04_Tables_Data`、`05_References`、`06_Permissions`、`07_Checklist`。

## Invocation modes

### Full audit (default)

任何完整自查表达都执行全部适用模块。DOCX 执行确定性检查与 AI 全文审核；Markdown/TXT 执行结构、标题、术语、正文逻辑、引文、参考文献、近期进展和科学表述检查，并把页面、字体、图像与 DPI 标为 `未自动测量`。

### Audit and safe repair

只有用户明确写出“自查并修复”“修复”“修改”“规范化”或“按规范调整”时进入 `audit-and-fix`：先对原稿完成全部确定性检查与 AI 全文审核；将原稿审核记录写为内部 `completed: true` JSON 后，调用内部修复流程。修复流程会再次确定性检查原稿、在新文件中安全规范格式，然后对新文件重新运行全部确定性检查和 AI 全文审核。输出 `第X章_规范修订版.docx`（冲突时用递增版本名），绝不覆盖原稿。

允许的机械修复仅限于：A4 与页边距、段落对齐/行距/缩进、明确的标题/图题/表题/参考文献样式、页码字段、简单编号格式及明确可统一的字体字号。修复前后核对源文件 SHA256 不变、正文 OOXML 文本签名不变、内嵌图片媒体字节哈希不变。禁止改变科学数据、数值、P 值、统计方法、实验结果、结论、机制、真实引文内容、DOI、生物学标签、图像像素、图源和版权信息。不得自动删文献或替换科学术语。

### Focused checks

以下请求重点检查指定模块，仍扫描并报告发现的 P0：

- `检查第5章参考文献`：引文映射、GB/T 7714-2025 字段、年份和重复项。
- `检查第5章图表`：图片对象、图题/表题编号与位置、DPI、公式及视觉复核边界。
- `检查第5章格式`：页面、section、页眉页脚、页码、样式继承、标题及段落格式。
- `检查第5章近5–10年研究进展`：前沿专题和文献年代；未联网时不得声称已验证最新研究。
- `检查第5章专著写作逻辑`：科学问题到应用的技术叙事、学位论文整合与重复内容。
- `完成第5章交稿前验收`：包括完整交稿包目录与来源/许可材料。

专项模式仍输出同一报告结构，并明确指出哪些模块未作为本轮重点审核；不能声称未执行的专项已经通过。

## Running the internal workflow

Python 命令只供 Agent/开发者内部使用。**Do not ask the user to manually run Python commands when the task can be executed by the agent.**

以 Skill 仓库目录作为脚本根目录，自动调用（将用户原句传给内部入口，由 Agent 自动发现文件）：

```text
python scripts/run_audit.py --workspace <当前workspace> --request "请调用 bioactive-surface-book-auditor 完成第X章自查"
```

若稿件是对话中已上传的附件，则将该文件路径作为内部 `--input`；不得要求用户重复提供路径。根据返回的 JSON，Agent 必须阅读全文并写出结构化语义复核结果：`completed: true`、`findings` 列表、`recent_progress.summary`，以及每个必需 `sections` 项的 `status` 和 `summary`。必需项为 `monograph_logic`、`thesis_integration`、`team_research_continuity`、`content_repetition`、`mechanism_evidence`、`clinical_extrapolation`、`terminology_consistency`、`statistics_review`、`figure_visual_review`、`copyright_integrity`。原稿与修订版的审核结果均需提供对应稿件的 `source_sha256`；报告生成器核对复核哈希与确定性检查稿件哈希一致。然后内部调用 `scripts/generate_report_docx.py` 生成报告。只有完整 AI 审核完成后才可将报告称为“完整自查报告”。

修复请求的顺序必须是：原稿 `audit-only` 确定性检查 → 原稿 AI 全文审核 → Agent 写入内部审核 JSON（包括 `completed: true` 与原稿 `source_sha256`）→ 调用 `run_audit.py --mode audit-and-fix --pre-repair-review <内部审核JSON>` → 修订版 AI 全文审核 → 生成以修订版检查结果为准的最终报告。由于自然语言原句包含修复意图，第一步内部运行时必须显式覆盖为 `audit-only`（使用已解析的章节号/已上传附件路径），不能把修复意图直接交给会执行修复的阶段。`audit-and-fix` 缺少完整标记或审核记录哈希与当前原稿不一致时必须拒绝修复。

运行依赖来自 `requirements.txt`。优先使用调用环境中已有 Python 和依赖；依赖缺失时由 Agent 在隔离环境安装依赖或说明明确限制，不把冗长 traceback 或手工命令交给普通作者。

## Standards and deterministic checks

`references/book-writing-standard.md` 是内部规范基准，出版社模板或用户更新要求优先。`references/audit-rubric.md` 定义 P0–P3 与结论。脚本只作可重复测量，不替代语义或专家判断。

适用时至少覆盖：

- 每个 Word section 的 A4 尺寸、方向、四边距、页眉页脚和页码字段；不可靠项目标记人工确认。
- run 直接格式、段落样式、样式继承链、Normal 与文档默认格式的有效字体字号解析；未知格式不得误报为异常。
- 第X章、X.1、X.1.1、X.1.1.1 的连续性、跳级、重号、章号、深度和一级小节数量；篇幅仅作规划提示。
- 第2–10章的前沿专题识别；近5年为当前年份减4至当前年，近10年为当前年份减9至当前年。当前年份默认 `datetime.now().year`，开发者可用 `--current-year` 覆盖。
- 参考文献总数、已识别/未知年份、近5/10年数量与比例、经典/未来年份；前沿专题单独统计。比例先报事实，不机械定合格线。
- GB/T 7714-2025 顺序编码、正文映射、按文献类型的核心字段、DOI 与重复候选。旧文献缺 DOI 不判错，不自动补 DOI，不自动删条目；不确定字段标 `△ 人工核验`。
- 图对象/图题数量、编号/章号/连续性、正文先引、矢量或位图、有效 DPI、低分辨率/网页截图候选、比例尺与图像完整性人工检查。
- 表格编号、表题在上、可编辑性、图片表格风险、边框/三线表、表头单位、跨页表头和表注；复杂表提醒保存 Excel 源文件。
- 公式编号/引用/编辑性候选，截图风险，以及化学结构的 ChemDraw/SVG/PDF 源文件核验。
- 公式变量斜体、单位与数学函数正体、上下标、说明性下标、符号首次解释和单位；化学结构优先核对 ChemDraw/CDX/CDXML 与 SVG/PDF 等可编辑源文件。图片截图只列为风险线索，不自动改写公式/结构。
- 统计 n、重复类型、mean±SD/SEM、median、误差棒、检验、P 值/显著性符号及图轴单位等启发式；检验适当性必须标记作者/统计人员确认，除非用户另供原始数据并明确委托统计复核。
- 数值单位空格、单位写法、rpm 与 ×g 候选、范围、上下标、符号、有效数字；不确定时只给候选，不自动改科学含义。
- 缩写首次定义、术语统一候选；只能提示疑似项，禁止自动替换术语。
- 若目录存在，检查 TOC 字段/手工目录、层级和页码；单章无目录不判错。
- 若提供交稿包，盘点 01–07 目录及文件数。

## Required AI full-text review

完整自查必须读全文，包括正文、表格、图题/图注、脚注（可提取时）和参考文献。长稿分段时保持章节定位和上下文，不得只看脚本结果或抽样。按以下逻辑判断科学问题、原理、构建方法、代表性研究、机制、性能/应用、近年进展和发展方向是否连贯；检查“某某研究了……”式逐篇堆叠、实验方法重复、摘要拼接、过渡缺失和团队技术脉络。

同时检查概念/背景/方法/结果重复，因果语言与证据是否匹配，体外→动物→临床外推，绝对化结论，缩写术语一致，统计报告可读性，前沿主题遗漏与综述/原始研究搭配。涉及“首次”“证实”“直接调控”“完全”“临床有效”等只提出证据核验和降调建议，不擅改科学结论。

版权、图源和科研诚信只给有出处的风险线索：来源不明/转载改编图、团队已发表图重用、Permission、显微比例尺、疑似重复或不当图像处理。不得无证据指控不端，也不得声称能自动确定版权许可。需要原图、许可凭证或原始数据才能判断时标 `△ 作者/专家人工确认`。

只有在确实查询并核对了来源后才能提出近年文献具体建议。联网研究应使用可核验来源，不得虚构文献、DOI、统计证据或研究进展。

## Report contract

报告固定为 `references/audit-rubric.md` 所规定的 11 个部分：总体结果、必须修改的问题、格式检查、图表公式统计、参考文献、内容和专著叙事、近5–10年进展、版权溯源科研诚信、交稿包、推荐顺序、最终勾选清单。总体结果包括稿件名、章节、时间、规模、小节数、图表公式、文献年限及 P0/P1/P2/P3 数量。只用以下结论：`可进入下一轮编辑`、`修改后可进入下一轮编辑`、`暂不建议交稿`，并写明：

> 该结论仅代表本书内部撰写规范自查，不代表出版社最终审核结果。

输出 Markdown 与可编辑 Word；Word 生成失败时保留 Markdown，并简短说明。若 AI 复核未完成，报告必须标记“草稿，不可作为完整自查结论”，不能称为已完成完整自查。

## Platform and responsibility limits

本 Skill filesystem-first、connector-free，路径使用相对路径与 `pathlib`，支持 Windows/macOS/Linux。原始稿件默认只读。AI 无法替代作者科学审核、版权负责人、统计专家、出版方审读，也不能在无图像查看能力时声称图片视觉检查通过。
