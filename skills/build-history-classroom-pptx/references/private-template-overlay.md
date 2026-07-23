# 私有模板覆盖层

## 目标

公开Skill不发布学校数字资产；授权教师个人安装后，仍能稳定优先使用真实云校模板。私有模板是运行时输入，不是开源仓库内容。

## 目录契约

在课件Skill根目录下使用固定路径：

```text
private-assets/
└── templates/
    ├── cloud-school-dark.pptx
    └── cloud-school-light.pptx
```

公开仓库的 `.gitignore` 必须排除整个 `private-assets/`。禁止使用强制添加、改名规避或压缩包嵌套等方式提交私有文件。

## 选择优先级

1. 用户在当前任务明确指定的模板路径；
2. 私有覆盖层中与主题对应的模板；
3. 文字化视觉规范。

未指定主题时为 `dark`。要求 `light` 但私有浅色模板缺失时，使用文字化浅色规范，不得擅自改用私有深色模板。

## 探测

运行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Resolve-PrivateTemplate.ps1 -Theme dark
```

只接受脚本返回的规范化绝对路径。返回 `private-template` 后验证文件可打开、可渲染且页面数大于零；返回 `written-fallback` 时不得声称使用了真实模板。

## 使用与保密

- 只读取模板的视觉语言、空间结构、品牌元素和版式节奏；
- 允许把模板复制到任务临时目录进行渲染，任务结束后不纳入交付；
- 不把模板、模板截图、提取图片、校徽或其他派生资产写入Git；
- 不在公开日志中输出模板二进制、截图或可还原的资产内容；
- 最终只交付课堂课件和教师备注；
- 教师备注记录 `用户指定模板`、`private-template` 或 `written-fallback`，不记录私有绝对路径。

## 失败处理

私有模板不存在、打不开或渲染失败时，明确降级到同主题文字规范并记录原因。不得因曾在其他任务看过模板而凭记忆宣称完成模板审阅。
