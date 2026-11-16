# jn-gongwen-format

济南卷烟厂公文格式 Skill —— 让 AI 一键生成符合厂内标准的 Word 文档。

一个标准的 Agent Skill 包，同时兼容 **WorkBuddy** 与 **豆包（Doubao）**。
生成脚本仅依赖 Python 3 标准库，无需 `python-docx`。

## 格式标准

| 项目 | 规范 |
|---|---|
| 页面 | A4；上 3.7cm，下 3.5cm，左 2.8cm，右 2.6cm |
| 主标题 | 方正小标宋简体，二号（22pt），居中 |
| 副标题 | 楷体_GB2312，三号（16pt），居中 |
| 一级标题 | 黑体，三号，如“一、” |
| 二级标题 | 楷体_GB2312，三号，如“（一）” |
| 三级标题 | 仿宋_GB2312，三号，如“1．” |
| 四级标题 | 仿宋_GB2312，三号，如“（1）” |
| 正文 | 标题下空一行；仿宋_GB2312，三号，两端对齐，固定行距 29 磅 |
| 标题行距 | 主、副标题固定值 29 磅 |

## 安装

### WorkBuddy

将整个 `jn-gongwen-format/` 目录复制到 `~/.workbuddy/skills/` 下，重启会话即可。
或在 WorkBuddy 中直接说“安装这个 skill”并附上本仓库地址。

### 豆包

将本目录打包为 ZIP（确保 `SKILL.md` 位于压缩包根目录），在豆包中上传技能即可。
也可在豆包对话中说明技能名称、触发场景与执行指令，让它协助创建。

> 需要现成 ZIP 时，运行 `python scripts/package_for_sharing.py`，会在技能目录旁生成两个安装包：
> `jn-gongwen-format.zip`（保留顶层目录，适合 WorkBuddy）与
> `jn-gongwen-format-doubao.zip`（`SKILL.md` 在压缩包根目录，适合豆包上传）。
> 两个包内容完全一致，只有 ZIP 内层级不同。

## 使用

安装后，直接用自然语言提出需求，例如：

> 按济南卷烟厂公文格式，帮我写一份关于开展公文格式培训的通知，包含总体要求、检查考核两部分。

或手动调用脚本：

```bash
python scripts/generate_gongwen.py assets/example_content.json -o 输出.docx
python scripts/verify_gongwen.py 输出.docx
```

快速生成覆盖全部四级标题的样例：

```bash
python scripts/generate_gongwen.py --self-test -o sample.docx
```

## 内容 JSON

```json
{
  "main_title": "主标题",
  "subtitle": "可选副标题",
  "paragraphs": [
    {"type": "body", "text": "正文段落"},
    {"type": "h1", "text": "总体要求"},
    {"type": "h2", "text": "统一思想"},
    {"type": "h3", "text": "明确责任"},
    {"type": "h4", "text": "细化节点"}
  ],
  "sign": "（济南卷烟厂）"
}
```

标题编号由脚本自动生成与重置，无需手写。也支持简化写法：
`{"title": "...", "body": ["段落1", "段落2"], "sign": "..."}`。

## 目录结构

```
jn-gongwen-format/
├── SKILL.md                       # 技能元数据与工作流
├── scripts/generate_gongwen.py    # 零依赖 docx 生成器
├── scripts/verify_gongwen.py      # 格式合规校验器
├── references/format_spec.md      # 完整格式规范与核对清单
└── assets/example_content.json    # 内容模板
```

## 字体说明

规范文本中的“仿宋GB2312 / 楷体GB2312”对应 Chinese Windows 实际安装名
`仿宋_GB2312` / `楷体_GB2312`。脚本默认使用带下划线的名称；如目标机器字体名不同，
修改 `scripts/generate_gongwen.py` 顶部的 `FONTS` 常量即可。

## License

MIT © 张栋
