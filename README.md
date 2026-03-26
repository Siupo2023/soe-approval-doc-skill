# SOE Approval Doc Skill - 使用指南

## 🎯 核心功能

### 1. 智能文档生成
- ✅ 基于系统思维的深度问题分析
- ✅ 符合国企行文规范的内容生成
- ✅ 逻辑严密的论证结构
- ✅ 简明扼要的语言风格

### 2. 格式自动化 ⭐ 新增
- ✅ 页面设置：A4 纸、标准页边距（上25、下25、左28、右26毫米）
- ✅ 字体规范：方正小标宋简体（标题）、仿宋GB2312（正文）
- ✅ 层次结构：一、→（一）→1.→（1）四级标题
- ✅ 行距控制：标题 30 磅、正文 28 磅
- ✅ 附件和落款格式规范

### 3. 多格式输出
- ✅ Markdown 文本（预览和编辑）
- ✅ Word 文档（.docx，完全符合格式规范）⭐ 新增
- ✅ PDF 文档（通过 Word 转换）⭐ 新增

---

## 🚀 快速调用

### 方式 1：使用 `/请示` 命令（推荐）

```
/请示 采购办公设备
/请示 修订管理制度
/请示 成立专项小组
```

### 方式 2：自然语言触发

```
请示助手，帮我写一份...
写请示，关于...
起草请示：...
我要写一份国企风格的内部请示
```

### 方式 3：Word/PDF 输出 ⭐ 新增

```
生成 Word 版本
导出为 PDF
按国企格式生成
应用格式规范
```

---

## 📄 Word 文档生成 ⭐ 新增

### 自动生成脚本

**脚本位置**：`~/.claude/skills/soe-approval-doc/scripts/generate_approval_doc.py`

### 使用方式

#### 在 Claude 中调用（推荐）

```
用户：请示助手，帮我写一份采购办公设备的请示，生成 Word 版本

Claude：
[收集信息]
1. 主送对象是谁？
2. 需要采购哪些设备？
...

[调用脚本生成 Word 文档]
✅ 文档已生成：/path/to/output.docx
```

#### 命令行直接调用

```bash
# 生成示例文档
python3 ~/.claude/skills/soe-approval-doc/scripts/generate_approval_doc.py output.docx

# 或使用绝对路径
python3 /Users/sibyl/.claude/skills/soe-approval-doc/scripts/generate_approval_doc.py ~/Desktop/请示.docx
```

### 格式规范速查

| 元素 | 规范 |
|------|------|
| 纸张 | A4 纵向 |
| 页边距 | 上25、下25、左28、右26（毫米） |
| 标题 | 方正小标宋简体，二号字，30磅行距，居中 |
| 正文 | 仿宋GB2312，三号字，28磅行距，首行缩进2字符 |
| 层次 | 一、→（一）→1.→（1） |
| 附件 | 正文下空1行，左空2字 |
| 落款 | 部门 + 汉字日期 |

---

## 📁 目录结构

```
~/.claude/skills/soe-approval-doc/
├── skill.md                        # 主技能文件
├── README.md                       # 本文件
├── references/                     # 参考资料
│   ├── format-standards.md        # ⭐ 格式规范详解（新增）
│   ├── style-guide.md              # 行文风格指南
│   ├── systems-thinking.md         # 系统思维参考
│   ├── templates/                  # 文档模板
│   │   ├── template-standard.md    # 标准模板
│   │   ├── template-procurement.md # 采购模板
│   │   └── template-policy.md      # 制度模板
│   └── examples/                   # 历史成功案例（待添加）
├── scripts/                        # 自动化脚本（新增）
│   └── generate_approval_doc.py   # Word 生成脚本
├── assets/                         # 资源文件（新增）
└── sessions/                       # 会话记录
```

### 关键参考文件

#### 1. 格式规范详解 (`format-standards.md`) ⭐ 新增
- 内容质量标准（六大要求）
- 页面格式规范
- 字体格式规范
- 层次结构规范
- 附件和落款格式
- 质量检查清单

#### 2. 行文风格指南 (`style-guide.md`)
- 国企公文语言规范
- 常用术语和表达
- 格式要求说明

---

## 🎯 使用方式对比

| 方式 | 适用场景 | 示例 |
|-----|---------|------|
| **/请示** | 快速启动，日常使用 | `/请示 采购电脑` |
| **自然语言** | 详细说明需求 | `请示助手，我需要写一份关于...` |

两种方式功能完全相同，都会：
1. 引导你提供必要信息
2. 对复杂问题进行系统思维深度分析
3. 生成符合国企规范的请示文档
4. 根据反馈迭代优化

---

## 📝 如何添加参考资料

### 添加成功案例

```bash
# 将已批准的请示（脱敏后）放入
~/.claude/skills/soe-approval-doc/references/examples/你的案例.md
```

案例格式：
```markdown
# 案例名称

**类型**：采购类/制度类/人事类
**特点**：这个案例的成功之处

---

[请示文档内容]
```

### 添加新模板

```bash
~/.claude/skills/soe-approval-doc/references/templates/template-新类型.md
```

---

## 💡 技能特点

### 1. 双重专业能力
- **国企公文专家**：严格遵循格式和语言规范
- **系统思维专家**：深度分析，找到根本解决方案

### 2. 智能场景识别
- **简单场景** → 标准流程
- **复杂场景** → 深度追问

### 3. 动态参考系统
- 所有参考资料可随时更新
- 无需修改技能文件
- 自动使用最新资料

### 4. 多种调用方式
- `/请示` - 快速命令
- 自然语言触发
- 两种方式功能相同

---

## 📌 重要说明

### 关于系统思维
- ✅ 信息收集阶段：深度提问
- ✅ 文档生成阶段：融入论证逻辑
- ❌ 最终文档：不使用术语
- ❌ 交互过程：不"讲课"

### 关于参考资料
- 建议定期添加成功案例
- 案例需先脱敏
- 优秀案例显著提升输出质量

---

## 🔧 依赖要求 ⭐ 新增

### Python 依赖

要使用 Word 文档生成功能，需要安装 `python-docx`：

```bash
pip3 install python-docx
```

或使用：

```bash
pip install python-docx
```

### 系统要求

- Python 3.6+
- python-docx 库
- 支持 Word 文档编辑的软件（Microsoft Word、WPS、LibreOffice 等）

### 检查安装状态

```bash
# 检查是否已安装
python3 -c "import docx; print('✅ python-docx 已安装')"

# 如果提示 "ModuleNotFoundError"，则需要安装
```

---

**文件关系**：
- `请示.md` - 单文件版本，支持 `/请示` 快速调用
- `soe-approval-doc/` - 完整版本，包含所有参考资料
- 两者共享同一套参考资料目录

**更新参考内容**：
只需编辑 `soe-approval-doc/references/` 下的文件，两个版本都会自动使用最新内容。

---

**开始使用**：试试 `/请示 你的需求` 吧！🏢
