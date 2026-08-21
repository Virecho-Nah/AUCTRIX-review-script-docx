# AUCTRIX Review Script DOCX

版本：2.0.0

创作者：**VIRÉCHO**

AUCTRIX 系列剧本审稿 Skill。用于审阅影视剧本、竖屏短剧、分集剧本、AI 漫剧、场景大纲等 Word 文档，诊断结构、单集发动、因果、人物动机与主体性、情绪递进、钩子、对白、视觉执行、连续性和制作可行性，并把意见作为真实 Word 批注锚定在相关原文段落。

本 Skill 把“文本证据与问题根因”置于“批注数量”之前：不为显得细致而堆满意见，不用泛泛的编剧理论替代具体判断，也不在用户未授权时直接改写剧本正文。每条批注都应说明问题如何影响理解、情绪、人物可信度、追更动力或制作执行，并给出最小可执行的修改方向。

## 作者与 AUCTRIX 定位

本 Skill 由 **VIRÉCHO** 创作，可独立安装、单独调用，作为完整的 Word 剧本全文诊断与原位批注工具使用。

同时，本 Skill 也是 VIRÉCHO 未来剧本 Agent **AUCTRIX** 的基础能力模块之一，计划在 AUCTRIX 工作流中承担全文诊断、问题聚类、编辑优先级判断、原位批注和 Word 保真交付。无论独立使用还是由 AUCTRIX 调用，都遵守同一原则：**审稿不是改戏，批注不是代写；具体问题必须有文本证据，修改建议必须能落到人物选择、情绪推进、观众追更或制作执行。**

## 主要能力

- 区分通用剧本与商业短剧／AI 漫剧模式，不把平台硬节拍套到所有项目
- 在商业短剧模式中优先检查开篇高光、主角行动线、单集／阶段钩子、情绪回合与信息差期待
- 完整阅读 Word 剧本后再判断，不边读开头边密集下结论
- 从全局结构、分集发动、人物、情绪、钩子、对白、视觉与制作等维度诊断
- 使用 P0—P3 优先级聚类问题，减少同一根因的重复批注
- 将意见写成真实 Word 批注，准确锚定相关原文段落
- 保留原文、格式、表格、已有批注和修订痕迹
- 对新增批注、OOXML 关系、锚点和可见正文指纹进行结构校验
- 配合文档渲染流程检查分页、字体、表格及版式

## 适用任务

- “帮我审一下这个 Word 剧本，并返回带批注版本”
- “按主编标准找出结构、人物动机和单集钩子的问题”
- “给这份短剧／AI 漫剧做全文诊断，不要直接改正文”
- “保留原格式和已有修订，在对应段落写专业 Word 批注”

本 Skill 不是自动改稿器。除非用户明确授权，它不会重写剧情、接受或拒绝修订、删除已有批注，也不会把脱离原文的长篇审稿报告替代原位批注。

## 安装

将以下内层文件夹复制到 Codex Skills 目录：

```text
auctrix-review-script-docx/
```

常见的个人 Skills 目录：

```text
~/.codex/skills/
```

不要把最外层发布仓库整体当作 Skill 安装，否则 `SKILL.md` 的层级可能不正确。

## 调用示例

```text
使用 $auctrix-review-script-docx 全文审阅这份 Word 剧本，在相关原文处加入主编式批注，并返回保留正文、格式、已有批注与修订痕迹的新 DOCX。
```

## 运行要求

- Python 3.10+
- Python 标准库与可用的 `lxml`
- 可处理 `.docx`／OOXML 的运行环境
- 推荐同时具备 Word 或 LibreOffice，用于最终视觉渲染检查

如果环境缺少渲染器，Skill 可以在结构审计通过后交付，但必须明确说明视觉 QA 未完成。

## 结构

```text
AUCTRIX-review-script-docx/
├── README.md
├── LICENSE
└── auctrix-review-script-docx/
    ├── SKILL.md
    ├── LICENSE
    ├── agents/
    │   └── openai.yaml
    ├── references/
    │   ├── commercial-short-drama-rubric.md
    │   ├── review-plan-format.md
    │   └── review-rubric.md
    └── scripts/
        ├── extract_docx.py
        ├── apply_review_comments.py
        ├── audit_review_comments.py
        └── test_roundtrip.py
```

## License

MIT
