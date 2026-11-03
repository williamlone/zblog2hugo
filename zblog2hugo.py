# -*- coding: utf-8 -*-
"""
zblog2hugo.py —— 把 Z-Blog ASP 2.3 数据库（zblog.mdb）里的文章转换成 Hugo 格式

功能说明：
    1. 读取 Access 数据库中的三张表：
        - blog_Article  ：文章表（标题、正文、发布时间等）
        - blog_Category ：分类表（cate_ID -> cate_Name）
        - blog_Tag      ：标签表（tag_ID -> tag_Name）
    2. 把每篇文章生成为一个 Markdown 文件（带 Hugo 的 YAML front matter），
       输出到 posts 目录，文件名格式：slug.md（slug 见第 4 点）。
    3. 文章正文（log_Content）由 HTML 转换成 Markdown 格式；
       不转换摘要（log_Intro）和评论，也不处理图片。
    4. slug 优先使用文章的 log_Url 字段，该字段为空时直接使用文章编号。

使用方法（通过命令行参数指定数据库和输出目录，均可省略，默认 zblog.mdb 和 posts）：
    python zblog2hugo.py [数据库文件] [输出目录]
    例如：
    python zblog2hugo.py                  # 使用默认值 zblog.mdb 和 posts
    python zblog2hugo.py zblog.mdb posts  # 手动指定数据库和输出目录

运行环境：
    - Python 3.x
    - pyodbc 库（pip install pyodbc）
    - markdownify 库（pip install markdownify，用于 HTML 转 Markdown）
    - 系统需安装 Microsoft Access ODBC 驱动（一般 Windows + Office 自带）
"""

import os
import re
import json
import argparse
import urllib.parse
import pyodbc
from markdownify import MarkdownConverter, chomp

# ==================== 配置区（按需修改） ====================

# Access ODBC 驱动名称（32/64 位需与 Python 位数一致）
ODBC_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"

# 时区偏移字符串，用于 Hugo front matter 的 date 字段（北京时间 +08:00）
TIMEZONE_SUFFIX = "+08:00"

# 最低发布等级：Z-Blog 中 log_Level >= 2 视为已发布，
# 小于该值（草稿、待审核）的文章会标记为 draft: true
MIN_PUBLISH_LEVEL = 2

# 正文里的博客地址占位符，输出时会被配置表里的 ZC_BLOG_HOST 替换
HOST_PLACEHOLDER = "<#ZC_BLOG_HOST#>"


# ==================== 数据库读取部分 ====================

def connect_database(db_path):
    """
    连接 Access 数据库，返回数据库连接对象。
    参数 db_path：数据库文件（.mdb）的路径（相对或绝对路径均可）。
    """
    # Access ODBC 驱动不支持相对路径，统一转成绝对路径
    db_path = os.path.abspath(db_path)
    # 拼接 ODBC 连接字符串
    conn_str = f"Driver={{{ODBC_DRIVER}}};DBQ={db_path};"
    # 建立并返回连接
    return pyodbc.connect(conn_str)


def load_categories(conn):
    """
    读取分类表 blog_Category，返回 {分类ID: 分类名称} 字典。
    参数 conn：数据库连接对象。
    """
    # 分类字典，用于后面查表
    categories = {}
    # cate_ID 是分类编号，cate_Name 是分类名称
    for row in conn.cursor().execute("SELECT cate_ID, cate_Name FROM blog_Category"):
        categories[row.cate_ID] = row.cate_Name
    return categories


def load_tags(conn):
    """
    读取标签表 blog_Tag，返回 {标签ID: 标签名称} 字典。
    参数 conn：数据库连接对象。
    """
    # 标签字典，用于后面查表
    tags = {}
    # tag_ID 是标签编号，tag_Name 是标签名称
    for row in conn.cursor().execute("SELECT tag_ID, tag_Name FROM blog_Tag"):
        tags[row.tag_ID] = row.tag_Name
    return tags


def load_blog_host(conn):
    """
    从配置表 blog_Config 的 Blog 记录里读取 ZC_BLOG_HOST（博客地址）。
    Blog 记录的值是一串用 \\x02 分成两段的文本：前一段是字段名，后一段是字段值，
    字段名和字段值内部都用 \\x01 分隔；字段值经过 URL 编码（例如 https%3A//...）。
    参数 conn：数据库连接对象。
    返回：解码后的博客地址，例如 "https://www.williamlong.info/weibo/"；
          找不到时返回空字符串。
    """
    # 用参数化查询取出 Blog 这条配置记录
    row = conn.cursor().execute(
        "SELECT conf_Value FROM blog_Config WHERE conf_Name = ?", "Blog"
    ).fetchone()
    if not row:
        return ""

    # \x02 前面是字段名列表，后面是字段值列表
    parts = (row.conf_Value or "").split("\x02")
    if len(parts) < 2:
        return ""
    names = parts[0].split("\x01")
    values = parts[1].split("\x01")

    # 找到 ZC_BLOG_HOST 在字段名列表中的位置，再取同位置的值
    if "ZC_BLOG_HOST" not in names:
        return ""
    index = names.index("ZC_BLOG_HOST")
    if index >= len(values):
        return ""
    # 字段值是 URL 编码的，解码后才是真实地址
    return urllib.parse.unquote(values[index])


def load_articles(conn):
    """
    读取文章表 blog_Article 中的所有文章，返回文章列表。
    每篇文章是一个字典，只包含转换需要的字段。
    参数 conn：数据库连接对象。
    """
    articles = []
    cursor = conn.cursor()
    # 只查询转换需要的字段（不读取摘要 log_Intro 和评论，正文统一转 Markdown）
    sql = """
        SELECT log_ID, log_CateID, log_Url, log_Title, log_Content,
               log_PostTime, log_Level, log_IsTop, log_Tag
        FROM blog_Article
        ORDER BY log_ID
    """
    for row in cursor.execute(sql):
        # 把一行数据整理成字典，字段含义见注释
        articles.append({
            "id": row.log_ID,            # 文章编号（用于生成 slug 和文件名）
            "cate_id": row.log_CateID,   # 所属分类的编号
            "url": (row.log_Url or "").strip(),  # 自定义 URL 别名（可能为空）
            "title": row.log_Title or "",  # 标题（空值转成空字符串）
            "content": row.log_Content or "",  # 正文（原始 HTML 格式）
            "post_time": row.log_PostTime,     # 发布时间（datetime 类型）
            "level": row.log_Level,      # 发布等级（>=2 为已发布）
            "is_top": bool(row.log_IsTop),     # 是否置顶
            "tag_field": row.log_Tag or "",    # 标签字段，格式如 "{1}{5}{8}"
        })
    return articles


# ==================== 转换处理部分 ====================

class ZBlogConverter(MarkdownConverter):
    """
    自定义 HTML 转 Markdown 的转换器，修正“链接文字与地址相同”的链接。
    markdownify 默认会把 <a href="url">url</a> 转成简写的 <url>，
    这种写法在 Hugo 里会失效，这里改成反引号包裹的代码文本：
        <a href="http://example.com">http://example.com</a>  ->  `http://example.com`
    """

    def convert_a(self, el, text, convert_as_inline):
        # 取出链接地址和标题
        href = el.get("href")
        title = el.get("title")
        # 链接文字与地址完全相同（且没有 title）时，输出反引号包裹的代码文本
        if (href
                and text.strip().replace(r"\_", "_") == href
                and not title
                and not self.options["default_title"]):
            # chomp 用来保留链接前后的空格，避免文字和链接粘连
            prefix, suffix, _ = chomp(text)
            return f"{prefix}`{href}`{suffix}"
        # 其他情况交给父类，按正常的 Markdown 链接处理
        return super().convert_a(el, text, convert_as_inline)


def parse_tag_ids(tag_field):
    """
    解析文章的标签字段，提取里面所有的标签编号。
    Z-Blog 的标签字段格式是花括号包裹的编号，例如 "{1}{5}{8}"。
    参数 tag_field：文章的 log_Tag 字段内容。
    返回：标签编号列表，例如 [1, 5, 8]。
    """
    # 用正则表达式找出所有 {数字} 里的数字
    return [int(num) for num in re.findall(r"\{(\d+)\}", tag_field)]


def make_slug(article):
    """
    生成文章的 slug（用于 Hugo 的文章地址和输出文件名）。
    规则：优先使用文章自带的自定义 URL（log_Url），为空时直接用文章编号。
    参数 article：文章字典。
    返回：slug 字符串，例如 "1201" 或 "my-custom-url"。
    """
    # log_Url 非空时用自定义 URL，否则回退到文章编号
    return article["url"] or str(article["id"])


def make_yaml_string(text):
    """
    把任意文本安全地转成 YAML 字符串（带双引号，支持中文和特殊字符）。
    技巧：JSON 的双引号字符串本身就是合法的 YAML 字符串。
    参数 text：原始文本。
    返回：带双引号的 YAML 安全字符串。
    """
    # ensure_ascii=False 保证中文原样输出，不变成 \uXXXX 转义
    return json.dumps(text, ensure_ascii=False)


def build_front_matter(article, categories, tags):
    """
    根据文章信息生成 Hugo 的 YAML front matter（文件开头 --- 之间的部分）。
    参数：
        article：文章字典（load_articles 返回的格式）
        categories：分类字典 {分类ID: 分类名称}
        tags：标签字典 {标签ID: 标签名称}
    返回：front matter 字符串（以 --- 开头和结尾）。
    """
    # 分类名称：通过文章的分类编号查分类表，查不到则跳过
    category_name = categories.get(article["cate_id"])
    # 标签名称列表：把标签编号逐个换成标签名称
    tag_names = [tags[tag_id] for tag_id in parse_tag_ids(article["tag_field"]) if tag_id in tags]

    # 发布时间格式：2006-07-14T10:39:14+08:00（Hugo 推荐的 ISO 8601 格式）
    date_str = article["post_time"].strftime("%Y-%m-%dT%H:%M:%S") + TIMEZONE_SUFFIX
    # 是否草稿：等级低于 MIN_PUBLISH_LEVEL 的文章标记为草稿，Hugo 不会发布它
    is_draft = article["level"] < MIN_PUBLISH_LEVEL
    # slug：优先用文章自带的自定义 URL（log_Url），为空时直接用文章编号
    slug = make_slug(article)

    # 逐行拼出 YAML 内容
    lines = [
        "---",
        # 文章标题
        f"title: {make_yaml_string(article['title'])}",
        # 发布时间
        f"date: {date_str}",
        # 文章地址标识
        f"slug: {make_yaml_string(slug)}",
        # 草稿标记
        f"draft: {str(is_draft).lower()}",
    ]
    # 置顶文章额外记录一个 weight，数值越小越靠前
    if article["is_top"]:
        lines.append("weight: 1")

    # 写入分类（Z-Blog 每篇文章只有一个分类）
    if category_name:
        lines.append("categories:")
        lines.append(f"  - {make_yaml_string(category_name)}")
    # 写入标签列表（可能没有标签，则不写 tags 字段）
    if tag_names:
        lines.append("tags:")
        for name in tag_names:
            lines.append(f"  - {make_yaml_string(name)}")

    # 收尾的 ---
    lines.append("---")
    # 每行之间用换行符连接
    return "\n".join(lines)


def build_body(article, blog_host):
    """
    生成文章正文部分（Markdown 文件里 front matter 之后的内容）。
    把原始 HTML 正文转换成 Markdown 格式，不包含摘要（log_Intro）。
    参数：
        article：文章字典。
        blog_host：博客地址，用于替换正文里的 <#ZC_BLOG_HOST#> 占位符。
    返回：Markdown 格式的正文字符串。
    """
    # 先替换占位符再转换：HTML 转 Markdown 时会把下划线转义成 \_，
    # 导致占位符变形（<#ZC\_BLOG\_HOST#>）而无法再匹配，所以必须在转换前替换
    content_with_host = article["content"].replace(HOST_PLACEHOLDER, blog_host)
    # 用自定义转换器把 HTML 转成 Markdown；标题用 # 风格（ATX），列表用 - 符号
    markdown_text = ZBlogConverter(
        heading_style="ATX",
        bullets="-",
    ).convert(content_with_host)
    # 转换后可能产生连续多个空行，统一压缩成最多一个空行（两个换行符）
    markdown_text = re.sub(r"\n{3,}", "\n\n", markdown_text).strip()
    return markdown_text


def build_markdown(article, categories, tags, blog_host):
    """
    把一篇 Z-Blog 文章组装成完整的 Hugo Markdown 文件内容。
    结构：front matter + 空行 + 正文。
    参数与 build_front_matter 相同，另外多一个 blog_host（博客地址）。
    返回：完整的文件内容字符串。
    """
    # front matter（标题、时间、分类、标签等元信息）
    front_matter = build_front_matter(article, categories, tags)
    # 正文内容（占位符 <#ZC_BLOG_HOST#> 会在正文生成时替换成博客地址）
    body = build_body(article, blog_host)
    # 组合成完整文件：front matter、空行、正文
    markdown_text = f"{front_matter}\n\n{body}\n"
    # 兜底：front matter 里若也出现占位符，一并替换
    return markdown_text.replace(HOST_PLACEHOLDER, blog_host)


def make_filename(article):
    """
    生成输出文件名，格式：slug.md。
    slug 规则与 make_slug 一致（log_Url 非空用 log_Url，否则用文章编号）。
    参数 article：文章字典。
    返回：文件名字符串，例如 "1201.md"。
    """
    # 文件名直接使用 slug + .md 后缀
    return f"{make_slug(article)}.md"


# ==================== 主流程 ====================

def convert_all(db_path, output_dir):
    """
    主转换流程：
    连接数据库 -> 读取分类/标签/文章 -> 逐篇生成 Markdown 文件到输出目录。
    最后在控制台打印转换统计结果。
    参数 db_path：数据库文件（.mdb）路径。
    参数 output_dir：Hugo 文章输出目录。
    """
    # 连接 Access 数据库
    conn = connect_database(db_path)
    # 预先读取分类表和标签表，转成 {ID: 名称} 字典方便查询
    categories = load_categories(conn)
    tags = load_tags(conn)
    # 读取配置表里的博客地址（用于替换正文中的 <#ZC_BLOG_HOST#> 占位符）
    blog_host = load_blog_host(conn)
    # 读取全部文章
    articles = load_articles(conn)

    # 输出目录不存在则创建
    os.makedirs(output_dir, exist_ok=True)

    # 文章总数，用于进度显示
    total = len(articles)
    # 逐篇文章转换并写出文件（enumerate 从 1 开始计数，方便显示进度）
    for index, article in enumerate(articles, start=1):
        # 生成 Markdown 内容（正文里的 <#ZC_BLOG_HOST#> 已替换为真实地址）
        markdown_text = build_markdown(article, categories, tags, blog_host)
        # 拼出完整输出路径
        output_path = os.path.join(output_dir, make_filename(article))
        # 以 UTF-8 编码写入文件（Hugo 要求 UTF-8）
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(markdown_text)
        # 单行刷新进度：\r 让光标回到行首，覆盖上一行，形成原地更新的效果
        print(f"\r转换进度：{index}/{total} 篇", end="", flush=True)
    # 进度显示结束后换行，避免后续输出接在同一行
    print()

    # 关闭数据库连接
    conn.close()
    # 打印统计信息
    print(f"转换完成：共 {len(articles)} 篇文章，文件已输出到 {output_dir}")


def parse_arguments():
    """
    解析命令行参数。
    两个参数都可以省略，省略时使用默认值：
        第 1 个：数据库文件路径（.mdb 文件），默认 zblog.mdb
        第 2 个：Hugo 文章输出目录，默认 posts
    返回：包含 db_path 和 output_dir 两个属性的对象。
    """
    # 创建参数解析器，description 会显示在 -h 帮助信息里
    parser = argparse.ArgumentParser(
        description="把 Z-Blog ASP 2.3 数据库（.mdb）里的文章转换成 Hugo 格式的 Markdown 文件"
    )
    # 第 1 个位置参数：数据库文件路径（nargs="?" 表示可以省略，省略时用 default）
    parser.add_argument(
        "db_path",
        nargs="?",
        default="zblog.mdb",
        help="Z-Blog 数据库文件路径（默认：zblog.mdb）",
    )
    # 第 2 个位置参数：输出目录（同样可以省略）
    parser.add_argument(
        "output_dir",
        nargs="?",
        default="posts",
        help="Hugo 文章输出目录（默认：posts）",
    )
    # 解析命令行参数
    args = parser.parse_args()
    # 数据库文件不存在时，打印用法（帮助信息）和错误提示后退出（退出码 2）
    if not os.path.exists(args.db_path):
        parser.error(f"找不到数据库文件：{args.db_path}（默认文件为 zblog.mdb，可把它放到脚本目录，或指定其他路径）")
    # 数据库存在，返回解析结果
    return args


# 脚本入口：直接运行本文件时才执行转换
if __name__ == "__main__":
    # 解析命令行参数，拿到数据库路径和输出目录
    args = parse_arguments()
    # 执行转换
    convert_all(args.db_path, args.output_dir)
