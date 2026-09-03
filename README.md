# zblog2hugo

把 Z-Blog ASP 2.3 的数据库（zblog.mdb）里的文章一键转换成 Hugo 博客格式的 Markdown 文件。

## 功能特性

- 读取 Access 数据库中的文章表（blog_Article）、分类表（blog_Category）、标签表（blog_Tag）
- 每篇文章生成一个 Markdown 文件（带 Hugo 的 YAML front matter）
- 正文由 HTML 自动转换成 Markdown 格式
- 分类、标签一并转换到 front matter 中
- slug 优先使用文章的 log_Url 字段，为空时直接使用文章编号
- 不转换摘要（log_Intro）、评论和图片

## 环境要求

- Python 3.x
- Windows 系统 + Microsoft Access ODBC 驱动（装有 Office 或 Access Database Engine 一般都有）
- 依赖库：

```bash
pip install pyodbc markdownify
```

> 注意：Access ODBC 驱动的位数（32/64 位）需要和 Python 的位数一致，否则会报“未发现数据源名称”。

## 使用方法

```bash
python zblog2hugo.py [数据库文件] [输出目录]
```

两个参数均可省略，省略时使用默认值 `zblog.mdb` 和 `posts`。

参数说明：

| 参数 | 说明 | 默认值 |
| --- | --- | --- |
| 数据库文件 | Z-Blog 的 .mdb 数据库路径（相对、绝对路径均可） | `zblog.mdb` |
| 输出目录 | Hugo 文章的输出文件夹（不存在会自动创建） | `posts` |

完整示例：

```bash
python zblog2hugo.py                  # 使用默认值
python zblog2hugo.py zblog.mdb posts  # 手动指定数据库和输出目录
```

运行完成后会提示转换结果，例如：

```text
转换完成：共 2100 篇文章，成功写出 2100 个文件到 posts
```

## 输出说明

每篇文章输出为一个 `slug.md` 文件，格式如下：

```markdown
---
title: "文章标题"
date: 2012-07-04T23:06:11+08:00
slug: "1201"
draft: false
categories:
  - "分类名称"
tags:
  - "标签一"
  - "标签二"
---

（正文，已转换成 Markdown 格式）
```

front matter 字段说明：

| 字段 | 来源 | 说明 |
| --- | --- | --- |
| title | log_Title | 文章标题 |
| date | log_PostTime | 发布时间（北京时间 +08:00） |
| slug | log_Url，为空则用 log_ID | 文章地址标识，同时用作文件名 |
| draft | log_Level | 等级 >= 2 为 `false`（发布），否则 `true`（草稿） |
| categories | log_CateID -> blog_Category | 分类名称（每篇文章一个分类） |
| tags | log_Tag -> blog_Tag | 标签名称列表 |
| weight | log_IsTop | 置顶文章会额外写入 `weight: 1` |

## 代码结构

| 函数 | 作用 |
| --- | --- |
| `connect_database` | 连接 Access 数据库 |
| `load_categories` / `load_tags` | 读取分类表 / 标签表，转成 {ID: 名称} 字典 |
| `load_articles` | 读取全部文章 |
| `make_slug` | 生成 slug（log_Url 为空时用文章编号） |
| `build_front_matter` | 生成 YAML front matter |
| `build_body` | 把 HTML 正文转换成 Markdown |
| `make_filename` | 生成输出文件名（slug.md） |
| `convert_all` | 主转换流程 |

## Hugo 站点配置建议

转换出的文章 front matter 带 `categories` 字段、少数正文含内嵌 HTML，建议在 Hugo 站点配置（hugo.toml）中加上：

```toml
[taxonomies]
  tag = "tags"
  category = "categories"   # 部分主题的 related 模板依赖 categories taxonomy

[markup]
  [markup.goldmark]
    [markup.goldmark.renderer]
      unsafe = true         # 允许渲染正文残留的内嵌 HTML，消除 Raw HTML omitted 警告
```

## 常见问题

**Q：提示“未发现数据源名称并且未指定默认驱动”**
A：系统缺少 Access ODBC 驱动，或驱动位数与 Python 不一致。可到微软官网下载 [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=13255)（位数要和 Python 一致）。

**Q：转换后的文章图片显示异常**
A：本工具不处理图片链接，如原文章引用了旧站图片，需要自行迁移图片并修改链接。
