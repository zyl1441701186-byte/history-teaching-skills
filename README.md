# History Teaching Skills

> 面向高中历史教师的轻量AI备课工作流：资料筛选、教学设计、原生PPTX课件。

![Version](https://img.shields.io/badge/version-H5-1E869F)
![License](https://img.shields.io/badge/license-MIT-C69A4B)

H5重新回到“少量强规则”的思路。三个Skill以各自的 `SKILL.md` 为核心，不再要求启动时读取大量规则文件，也不再生成多套矩阵和过程清单。

## 工作流

```text
教师指定资料文件夹
        │
        ▼
curate-history-sources
        │
        ▼
资料摘要.md
        │
        ▼
design-history-lesson
        │
        ▼
教学设计.md
（末尾包含逐页蓝图与素材建议）
        │
        ▼
教师确认一次教学结构与素材选择
        │
        ▼
build-history-classroom-pptx
        │
        ├── 课堂课件.pptx
        └── 教师备注.md
```

资料筛选启动时，在教师指定文件夹内创建 `yyyyMMdd-HHmmss_Trae_课标题` 任务目录。最终只保留四个教学成果：

```text
任务目录/
├── 资料摘要.md
├── 教学设计.md
├── 课堂课件.pptx
└── 教师备注.md
```

中间文件统一放入任务目录内的 `工作文件/`，不作为交付物。

## 三个Skill

### curate-history-sources

- 遍历教师已经筛选的本地资料；
- 定向联网核对课标、出处、冲突和核心缺口；
- 并列呈现2025年与2020年修订课标原文及变化；
- 整合知识、史料、试题、教学思路和多类型素材建议；
- 推荐一条有史学依据的教学主线，必要时提供一个备选；
- 只交付 `资料摘要.md`。

### design-history-lesson

- 使用“环节”组织真实课堂；
- 区分导入锚点、教学主线和核心问题；
- 保证基础知识、材料处理和学生笔记；
- 组织文物、地图、史料、表格、示意图等多类型材料；
- 形成材料—任务—思考—答案—结论的课堂推进；
- 把逐页课件蓝图和素材建议合并进 `教学设计.md`。

### build-history-classroom-pptx

- 直接生成原生可编辑PPTX，不做HTML转PPT；
- 优先参考本地私有云校深色模板；
- 保持统一顶栏、底栏和完整L1/L2/L3标题层级；
- 保证知识密度、课堂链条和学生笔记；
- 允许教师后补图片、地图、漫画和视频占位符；
- 表格答案与结论在同一页动画揭示；
- 禁止页码、“笔记”提示、AI标记和制作残留；
- 只交付 `课堂课件.pptx` 与 `教师备注.md`。

## 公开版与个人版

公开仓库不包含学校数字资产和课程标准PDF。

个人使用时，可以在课件Skill中放入：

```text
build-history-classroom-pptx/
└── private-assets/
    └── templates/
        ├── cloud-school-dark.pptx
        └── cloud-school-light.pptx
```

未指定主题时优先使用深色模板。模板只作为本地视觉参考，不进入开源仓库和最终课件交付。

2025年修订课程标准PDF可由教师在任务中直接提供，或放入个人Skill的 `assets/` 目录。

## 安装

将 `skills/` 下三个目录复制到TRAE Skills目录，例如：

```text
C:\Users\<用户名>\.trae-cn\skills\
```

H5公开版结构：

```text
skills/
├── curate-history-sources/
│   └── SKILL.md
├── design-history-lesson/
│   └── SKILL.md
└── build-history-classroom-pptx/
    ├── SKILL.md
    ├── agents/
    └── scripts/
```

PPT脚本用于私有模板探测、统一主题、原生动画和文字溢出检查。它们可以直接执行，不要求模型读取脚本全文。

## 使用

第一步：

```text
请使用 curate-history-sources，遍历“D:\备课资料\第7课”，
完成双版本课标对照、资料核验、材料整理和教学主线建议。
```

确认资料与主线后：

```text
请使用 design-history-lesson，根据资料摘要撰写完整教学设计，
并在文末给出逐页课件蓝图和素材建议。
```

确认教学设计与素材选择后：

```text
请使用 build-history-classroom-pptx，生成原生课堂课件和教师备注。
```

## 边界

- AI不得凭记忆补写课标、史料原文、数据和学术观点；
- 网络搜索必须针对具体缺口，并打开完整来源；
- 外部图片和视频由教师最终选择；
- 教学主线不能替代知识教学；
- 课件优先服务真实课堂，而不是商务展示；
- 无法可靠完成原生动画或PPTX检查时应直接说明，不伪造通过。

## 版本

- `V1`：最初公开版本；
- `H2`：完整工作流重构；
- `H3`：区分公开版与私有模板；
- `H4`：强化课堂可用性检查；
- `H5`：删除重复规则、矩阵和过程文件，回归轻量单文件Skill。

具体变化见 [CHANGELOG.md](CHANGELOG.md)。

## License

MIT
