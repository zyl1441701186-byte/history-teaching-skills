---
name: build-history-classroom-pptx
description: 根据已确认的高中历史教学设计、课件制作清单、素材建议表和资料摘要，先生成逐页素材选择表并等待教师确认，再制作原生可编辑、知识覆盖完整、带规范图片视频占位符和逐页教师备注的PowerPoint课件。用于高中历史PPTX制作、问题与答案分步揭示、学生笔记落点、原生动画、文字防溢出和逐页验收；不联网选图，不自动插入外部图片视频，禁止HTML转PPT和商务汇报化。
---

# 目标

交付一套服务真实课堂、便于教师后续手动补入图片或视频的历史课件。课件必须同时承担：问题推进、材料处理、知识落实、学生笔记和课堂检测，不能只展示文学化主线。

最终只交付：

- `课堂课件.pptx`
- `教师备注.md`

逐页蓝图、素材选择确认表、构建脚本、动画计划、渲染图和检查记录只保留在任务工作区。

# 必读规则

开始前完整读取：

- [references/input-and-boundaries.md](references/input-and-boundaries.md)
- [references/classroom-narrative.md](references/classroom-narrative.md)
- [references/content-capacity-and-knowledge.md](references/content-capacity-and-knowledge.md)
- [references/media-confirmation-and-placeholders.md](references/media-confirmation-and-placeholders.md)
- [references/visual-system.md](references/visual-system.md)
- [references/evidence-visuals.md](references/evidence-visuals.md)
- [references/animation-and-qa.md](references/animation-and-qa.md)
- [references/teacher-notes-template.md](references/teacher-notes-template.md)

若环境提供通用PPTX Skill，读取其生成、渲染和溢出检查方法，但以下规则优先：

- 主色固定为 `#1E869F`；
- 画布固定为33.867 cm × 19.05 cm；
- 正文不小于24 pt，默认28 pt以上；
- 标题不小于28 pt，默认32 pt以上；
- 每页显示完整标题路径和统一顶栏、底栏；
- 教学必要时允许高密度页；
- 结论条使用高饱和橙色，放大居中；
- 不自动搜索或插入外部图片、地图、漫画和视频；
- 不采用商务卡片墙、仪表盘和HTML转PPT；
- 不交付源码、素材表和检查文件。

# 启动条件和优先级

必须有已经确认的：

1. `教学设计.md`
2. `课件制作清单.md`
3. `课件素材建议表.md`

还应读取 `资料核心摘要.md` 以约束史实、引文、概念和来源。

优先级：教学设计决定目标、主线、环节和知识；课件制作清单决定课堂动作和揭示顺序；素材建议表决定需要教师确认的外部视觉意图；资料摘要决定证据边界；用户模板只控制视觉参考。

冲突影响史实、知识或教学逻辑时停止，不得自行裁决。

# 强制技术路线

- 直接生成原生 `.pptx`，优先使用PptxGenJS和环境PPTX工具。
- 禁止HTML、网页截图、PDF或整页图片作为版式中间产物。
- 学生可见文字、标题、结论、表格、时间轴和示意图保持原生可编辑。
- 使用 [scripts/history_theme.mjs](scripts/history_theme.mjs) 或严格复现其常量与占位符函数。
- 需要动画的对象设置稳定对象名。
- 先运行 [scripts/Test-PowerPointAnimation.ps1](scripts/Test-PowerPointAnimation.ps1)；失败时停止。
- 动画后运行 [scripts/Inspect-PptxTiming.ps1](scripts/Inspect-PptxTiming.ps1)。
- 每次生成后运行 [scripts/Test-PowerPointTextOverflow.ps1](scripts/Test-PowerPointTextOverflow.ps1)，任何溢出或自动缩字都必须修正。

# 工作流

## 1. 建立任务工作区

```text
work/
├── source/
├── build/
├── rendered/
├── qa/
├── 素材选择确认表.md
└── animation-plan.json
```

不修改用户原文件。

## 2. 执行知识覆盖审计

从教学设计的知识覆盖矩阵提取全部必备知识，建立内部映射：

```text
知识ID → 教学环节 → 课堂动作 → 学生可见页面 → 笔记落点
```

以下内容只出现在教师备注中不算覆盖：基础史实、关键概念、制度内容、主要过程、因果关系、历史影响和要求学生记录的结论。

教学主线只作为组织线索，不得占用页面主体却挤掉知识内容。发现课件清单漏项时停止并回到教学设计修正，不得由PPT Skill擅自补写史实。

## 3. 建立逐页蓝图

每页登记：

- 页码和完整标题路径；
- 所属教学环节和知识ID；
- 唯一主要课堂任务；
- 学生初始可见内容；
- 文字、数据和原生图形；
- 外部视觉素材意图；
- 学生任务和预期产出；
- 笔记落点与结论条；
- 动画对象、顺序和揭示内容；
- 与前后页关系；
- 建议用时及必讲、可压缩、可跳过标记。

按学生认知和页面容量拆页，不按教案章节机械转换。

## 4. 执行教师素材确认关口

根据逐页蓝图生成临时 `素材选择确认表.md`，列出每一页的外部图片、地图、漫画和视频建议。

每页必须由教师选择：

- 不使用素材；
- 保留占位符；
- 已有指定素材但仍保留占位符；
- 暂不确定，保留可删除占位符。

必须暂停并等待教师确认。用户已提供完整确认表时可以直接读取。未确认不得制作整套PPT。

确认“不使用素材”的页面重新设计为纯文字、表格、时间轴或原生关系图，不得留下空洞区域。其他选择只生成占位符，不插入实际外部文件。

## 5. 执行动画能力预检

运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-PowerPointAnimation.ps1 -WorkDir <任务工作区>
```

只有真实创建、保存、重开并读回动画序列后才能继续。

## 6. 原生制作PPTX

- 使用固定画布、页框、字体和配色；
- 每页显示一级、二级和三级标题（存在时）；
- 先安排知识、材料和学生任务，再安排答案和结论；
- 外部视觉素材统一使用 `MEDIA-PXX-XX` 占位符；
- 表格、时间轴、制度图、因果图和简单统计图可以原生生成，但只能使用已核验知识或数据；
- 对全部文字显式设置字体、字号、边距和行距；
- 禁止自动缩小文字和整页扁平化。

## 7. 执行容量与溢出检查

先做版式容量检查，再运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Test-PowerPointTextOverflow.ps1 -PptxPath <课堂课件.pptx> -OutputJson <qa/text-overflow.json>
```

发现溢出时按顺序处理：删除重复修饰、改成结构化表达、拆为同层级续页、将教师背景说明移入备注。不得删除必备知识，不得突破字号下限，不得启用自动缩字。

## 8. 写入并验证原生动画

默认只使用出现、淡入和擦除，以单击触发为主。问题先出现，学生处理后再显示证据强调、答案和结论条。

动画写入后重新打开并检查时间轴、对象顺序、文件修复提示和备注一致性。再次运行文字溢出检查，防止动画保存过程改变版式。

## 9. 生成教师备注

逐页记录：完整预期回答、追问、纠偏、过渡、时间、动画单击、学生笔记和外部素材手动插入说明。

每个占位符必须在备注中说明：建议内容、教学作用、比例、裁切、图注、来源核验、动画时机以及不插入时的替代方案。

## 10. 逐页验收

必须完成：

1. 渲染全部页面为PNG；
2. 逐页全尺寸检查；
3. 运行文字溢出与自动缩字检查；
4. 检查知识覆盖矩阵全部落到学生可见页面或笔记结论；
5. 检查完整标题路径、顶栏、底栏和主色一致；
6. 检查正文、材料、表格、人工标注均不低于24 pt；
7. 检查标题不低于28 pt，结论不低于32 pt；
8. 检查问题先于答案、结论最后出现；
9. 检查占位符ID、比例和教师备注一致；
10. 检查PowerPoint可正常打开、保存和放映；
11. 检查所有计划动画均为原生时间轴。

任何一项未通过都必须修改并重新检查。

# 失败与降级

- 缺少已确认教学设计、课件清单或素材建议表：停止。
- 教师尚未完成逐页素材选择：暂停，不得生成整套课件。
- 知识覆盖矩阵与课件清单冲突：返回教学设计修正。
- 无法生成原生PPTX、原生动画、逐页渲染或文字溢出检查：停止。
- 不得静默改为HTML、静态PPT、图片化页面或复制页伪动画。

只有用户明确同意后才可采用复制页模拟动画，并在备注中标明。

# 交付

最终目录只包含：

```text
课堂课件.pptx
教师备注.md
```

交付说明只报告总页数、动画页数、占位符数量、知识覆盖是否通过、文字溢出是否为零、是否采用模板及是否完成逐页渲染和打开测试。
