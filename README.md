# 最简单的使用方法

把章节 Word 放入工作区，然后在支持 Skills 的 Agent 中输入：

> 请调用 bioactive-surface-book-auditor 完成第5章自查

Skill 自动定位章节，并依次完成格式、图表、参考文献、统计信息、正文逻辑、近年进展、科学表述与版权/科研诚信风险检查，生成 Markdown 和 Word 报告。普通作者无需手动运行 Python、填写路径或指定输出目录。默认只检查，不修改原稿。

其他常用说法：

- `请调用 bioactive-surface-book-auditor 自查第5章`
- `请用 bioactive-surface-book-auditor 检查第5章是否符合书稿规范`
- `请调用 bioactive-surface-book-auditor 完成“角膜与眼表修复材料的生物活性界面”章节自查`
- `请调用 bioactive-surface-book-auditor 完成第5章自查并修复`（新建规范修订版，不覆盖原稿）
- `请调用 bioactive-surface-book-auditor 检查第5章参考文献`
- `请调用 bioactive-surface-book-auditor 检查第5章图表`
- `请调用 bioactive-surface-book-auditor 检查第5章格式`
- `请调用 bioactive-surface-book-auditor 检查第5章近5–10年研究进展`
- `请调用 bioactive-surface-book-auditor 检查第5章专著写作逻辑`
- `请调用 bioactive-surface-book-auditor 完成第5章交稿前验收`

若只提供章节号/章节标题，Skill 会搜索当前 workspace；DOCX 优先，其次 Markdown、TXT。用户已上传稿件时直接使用上传文件。只有没有匹配稿件或存在无法判定的同优先级版本时才会追问。

报告默认写入稿件所在目录的 `自查结果/`：

- `第5章_自查报告.docx`：普通作者优先阅读
- `第5章_自查报告.md`
- `第5章_自动检查结果.json`：机器可读的确定性检查结果
- 明确要求修复时，另生成 `<稿件名>_规范修订版.docx`

## 安装

将本仓库作为 `bioactive-surface-book-auditor` Skill 安装到你使用的 Agent 所支持的 Skills 目录中，并保持 `SKILL.md`、`references/` 与 `scripts/` 的相对位置不变。依赖见 `requirements.txt`。日常使用不需要手动安装 Python 包；Agent 应优先使用现有运行环境，需要补依赖时由 Agent 在隔离环境处理。

## Codex 使用

将仓库安装为 Codex 可发现的 Skill 后，在包含章节文件的 workspace 中输入上方的一句话调用即可。Skill 会读取仓库规范、自动寻找稿件、运行内部检查并完成全文语义复核。对于已上传附件，直接选用该附件。

## ChatGPT Work

在支持本地文件访问及 Skills 的 Work 环境中，将 Skill 加入可用 Skills，并把章节文件放入本次任务可访问的工作区。使用相同的一句话调用。若 Work 环境不提供本地文件访问，先将稿件作为附件提供；不要求作者填写脚本参数。

## WorkBuddy

在能加载本地 Skill 文件和执行 Python 的 WorkBuddy 环境中导入整个 Skill 目录，并将稿件置于 Agent 可访问的工作区。用自然语言调用；Agent 按 `SKILL.md` 调用内置工作流并完成人工语义复核。

## 豆包工作模式

在能够读取 Skill 目录、workspace 文件并运行本地脚本的豆包工作模式中加载整个 Skill。若平台本身不支持自动发现 Skill，则将 `SKILL.md` 作为任务说明提供，并保持 `references/`、`scripts/` 与它相邻。对话式模型必须阅读全文并完成 AI 语义部分，不能仅运行 Python 后结束。

## Developer / 高级用法

下面命令仅供开发者、测试人员或 Agent 内部调用，普通作者不需要使用。先在本仓库根目录安装 `requirements.txt` 中依赖。

```bash
python scripts/run_audit.py --workspace /path/to/workspace --chapter 5
python scripts/run_audit.py --workspace /path/to/workspace --request "请调用 bioactive-surface-book-auditor 完成第5章自查"
python scripts/run_audit.py --input /path/to/Ch05.docx --chapter 5
python scripts/audit_docx.py /path/to/Ch05.docx --current-year 2026
python scripts/audit_text.py /path/to/Ch05.md --current-year 2026
python scripts/run_audit.py --input /path/to/Ch05.docx --chapter 5 --mode audit-and-fix --pre-repair-review /path/to/自查结果/原稿_AI审核.json
python scripts/normalize_docx.py /path/to/Ch05.docx --output /path/to/Ch05_规范修订版.docx
python scripts/generate_report_docx.py --audit-json /path/to/audit.json --review-json /path/to/ai-review.json --output-md /path/to/自查结果/第5章_自查报告.md --output-docx /path/to/自查结果/第5章_自查报告.docx
```

`run_audit.py` 负责发现文件、确定性检查和可选修复；DOCX 修复模式必须提供原稿完整 AI 全文审核记录（每个必需语义模块均有状态/摘要、近年进展摘要和 `completed: true`），并验证审核记录中的源文件 SHA256。缺项或哈希不符会拒绝修复。低层级 `normalize_docx.py` 示例只供开发者使用，调用前仍须完整检查原稿，调用后还要复查新文件。脚本不能代替 Agent 对全文的语义、视觉和版权/科研诚信风险复核。`generate_report_docx.py` 只有在传入完整覆盖必需模块的复核 JSON 后才将报告视为完整报告。Markdown/TXT 的页面、字体和图片 DPI 标为未自动测量。

## 安全修复

只有请求明确包含“修复、修改、规范化、按规范调整”等意思时才修复。先对原稿完成确定性与 AI 全文审核，之后才允许创建新副本；对修订版再次完成确定性与 AI 审核。在新文件中仅规范可机械判断的排版，不碰科学内容，并验证原稿 SHA256、正文文本签名和嵌入图片字节保持不变。输出存在时不会覆盖，会选择递增版本名。禁止自动改数据、P 值、统计方法、实验结论、机制、真实引文、DOI、术语、图片像素、来源或版权声明。

## 检查范围与限制

DOCX 可做页面、样式继承、编号、图像有效 DPI、图片/图题、Word 表格/公式对象、参考文献和单位等可测检查；Markdown/TXT 做内容与结构检查。完整自查还要求模型通读全文，按专著逻辑审核学位论文素材整合、内容重复、机制和临床外推、统计表达、术语以及近年研究主题，并把发现写入报告。

即使自动检查通过，以下事项仍需相应人员确认：图片比例尺、面板与标签、图像处理和科研诚信；原始图源与第三方版权许可；统计检验是否恰当；科学主张与原始证据是否一致；前沿主题和最新文献是否完整；复杂表格/公式/化学结构的可编辑源文件。AI 不得编造文献或 DOI，也不能代替作者科学审核、版权人员、统计专家或出版社终审。
