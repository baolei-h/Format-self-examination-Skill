# 推荐调用方式

## 完整自查（默认只读）

`请调用 bioactive-surface-book-auditor 完成第5章自查`

## 自查并修复（生成副本，不覆盖原稿）

`请调用 bioactive-surface-book-auditor 完成第5章自查并修复`

## 专项检查

- 参考文献：`请调用 bioactive-surface-book-auditor 检查第5章参考文献`
- 图表：`请调用 bioactive-surface-book-auditor 检查第5章图表`
- 格式：`请调用 bioactive-surface-book-auditor 检查第5章格式`
- 近年研究：`请调用 bioactive-surface-book-auditor 检查第5章近5–10年研究进展`
- 专著逻辑：`请调用 bioactive-surface-book-auditor 检查第5章专著写作逻辑`
- 交稿验收：`请调用 bioactive-surface-book-auditor 完成第5章交稿前验收`

普通作者无需手动运行 Python，无需填写文件路径或输出参数。若只给章节号/标题，Skill 自动搜索 workspace；若稿件已上传，直接使用上传文件。默认完整自查且不修改原稿；仅明确要求修复时生成新修订副本。

# 其他调用示例

## Codex / WorkBuddy / 豆包工作模式

- `使用 bioactive-surface-book-auditor 自查 Ch05_骨与硬组织植入材料生物活性表面.docx，只报告问题，不修改文件。`
- `按《生物活性表面可控构建技术》撰写规范检查这个章节，重点检查字体字号、图表、参考文献和第15节近5-10年进展。`
- `自查这个章节交稿包，检查正文、FigureSources、References、Permissions 是否齐全。`
- `先自查，再安全修复所有可自动修复的格式问题；不要改科学内容。修复后重新跑一次自查。`
- `只检查第6章的引文与参考文献：正文引用是否一一对应、编号是否连续、是否符合GB/T 7714-2025框架。`

## 纯文本模式

- `下面是第4章正文纯文本。使用该Skill做内容、标题层级、术语、引文和参考文献自查；页面、字体和图片DPI标为无法测量。`
