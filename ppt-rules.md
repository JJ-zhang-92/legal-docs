# PPT 生成技能 — 技术与设计规范

> 本规范基于真实项目中生成单页 PPT 的经验总结。

## 一、工作纪律

| 规则 | 说明 |
|------|------|
| 固定输出路径 | 用户提供的工作文件夹是唯一输出位置，不跳到其他目录 |
| 不新建无关文件夹 | 需要子目录或外部工具（如 git clone）时，先确认位置 |
| 不改动用户原始文件 | 所有输出使用新文件名，不覆盖 |

## 二、技术禁忌

| 禁忌 | 原因 | 正确做法 |
|------|------|---------|
| 从零用 python-pptx 新建 PPTX | 空模板 + 自定义尺寸 + XML 操作，PowerPoint 可能拒绝打开 | **在已有可打开文件上做减法再加法** |
| 用 `python -c` 内联含中文路径的命令 | PowerShell GBK 编码截断、Unicode 解析失败 | 写成 `.py` 脚本文件再执行 |
| 保存到 PowerPoint 正打开的文件 | Permission Denied | 输出到新文件名 |
| 假设 python-pptx 能读写 = PowerPoint 能打开 | python-pptx 解析宽松，PowerPoint 校验严格 | 用 COM `PowerPoint.Presentations.Open()` 真实验证 |
| pip install 含 Unicode 的 requirements.txt | GBK 解码失败 | 设置 `PYTHONUTF8=1 PYTHONIOENCODING=utf-8` |
| git clone 大仓库未用 shallow | 超时 | `git clone --depth 1` |

## 三、设计参数（本次项目特例，不可外推）

以下数值基于本模板约束：5.35" 栏宽 / 网格纸背景 / 0.76-0.80" 卡片高 / 11pt 基准正文

| 参数 | 本次安全值 | 失效条件 |
|------|-----------|---------|
| 左栏正文 | 11pt / 0.34" 间距 | 栏宽 < 5" 或行长 > 30 字符需缩小 |
| KPI 数字 | 13-14pt | 数字超过 4 个汉字需缩小或减字 |
| 案例描述 | 10pt | 描述 > 20 字单行需缩小到 9pt 或拆行 |
| 卡片高度 | ≥0.76" | 仅限单行标题+双行描述；单行可降到 0.60" |
| 标签 | 8.5pt | 标签数 > 11 个或栏宽 < 5" 需缩至 8pt |
| 标签宽度公式 | `len×0.14+0.28"` | 仅限微软雅黑 8.5pt；换字体或字号须重测 |
| 页脚 | 9-10pt | 单行文本 |

**通用规则：每换一个模板，字号和间距必须重新验证，不能照搬。**

## 四、融合方法论

```
最终版 = 用户提供的模板（背景/配色/装饰/页脚）
       + 保留原模板装饰层（网格纸、纸张纤维、线条）
       + 删除原模板内容区形状
       + 新建内容层：卡片式案例 + 圆角徽章标签
       + 从其他版本择优取文字内容
```

**关键原则**：格式和内容分离。在已知可用的基底文件上做减法再加法，比从零新建风险低两个数量级。

## 五、ppt-master 使用边界

| 适用场景 | 原因 |
|---------|------|
| 从零生成多页 Deck（10-15 页） | 完整管线：文档→SVG→PPTX，支持 General / Consultant / Consultant Top 三种风格 |
| 大批量文档转 Markdown | `pdf/doc/ppt_to_md.py` 提取稳定可靠 |
| 学习设计规范 | `executor-consultant.md`、`shared-standards.md` 色彩/字号/KPI/阴影规范可参考 |

| 不适用场景 | 原因 |
|-----------|------|
| 单页 PPT 美化/微调 | 管线太重 |
| 中文单页 SVG→PPTX 导出 | 兼容性差，打不开 |
| 需要保留原模板背景/装饰 | ppt-master 是完整替换，不保留原设计 |

| 可独立复用的组件 | 路径 |
|-----------------|------|
| PPT 文本提取 | `skills/ppt-master/scripts/source_to_md/ppt_to_md.py` |
| PDF 文本提取 | `skills/ppt-master/scripts/source_to_md/pdf_to_md.py` |
| 设计规范参考 | `skills/ppt-master/references/executor-consultant.md` |
| SVG/PPTX 技术约束 | `skills/ppt-master/references/shared-standards.md` |

**最佳实践路径**：
```
ppt-master 提取脚本（取内容）
      +
ppt-master 设计参考（取规范）
      +
python-pptx 执行（生成/修改）
      +
COM 验证（PowerPoint.Presentations.Open）
```

## 六、验证检查清单

每次生成 PPT 后逐项确认：

- [ ] COM `Presentations.Open()` 可打开
- [ ] 所有文字不超出形状边界
- [ ] 字号/间距/配色符合规格表
- [ ] 输出文件在用户指定工作文件夹内
- [ ] 不覆盖任何原有文件
