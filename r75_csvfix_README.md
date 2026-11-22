# csvfix.py —— CSV 乱码修复 / 批量合并 / 去重

单文件零依赖，Python 3 直接跑。专治「Excel 双击 CSV 中文全是乱码」这一个老毛病，顺带把多份表合并、去重一起解决。

## 它会做什么

Excel 打开 CSV 乱码，90%是因为文件是 GBK 或纯 UTF-8 没带 BOM。Excel 只认「UTF-8 带 BOM」这个格式。
`csvfix.py` 会把所有 CSV 统一转成 Excel 友好的 UTF-8-BOM，同时能：

- 自动识别原编码（UTF-8 / UTF-8-BOM / GBK / UTF-16）
- 自动识别分隔符（逗号、制表符、分号、竖线）
- 批量合并多个 CSV，表头只保留一次
- 按整行去重，可选追加「来源文件」列
- 默认只预览，加 `--apply` 才真写盘
- 支持原地覆盖，自动留 `.bak`，还带一键撤销

## 快速开始

```bash
# 1. 先预览：看看哪些文件要转、原来是什么编码
python3 csvfix.py ./下载的表格

# 2. 真转：结果放到 ./下载的表格/_utf8
python3 csvfix.py ./下载的表格 --apply

# 3. 原地覆盖（自动备份 .bak）
python3 csvfix.py ./下载的表格 --apply --inplace

# 4. 合并成一个文件
python3 csvfix.py ./下载的表格 --merge 汇总.csv

# 5. 合并 + 去重 + 标记来源
python3 csvfix.py ./下载的表格 --merge 汇总.csv --dedupe --source

# 6. 搞错了，撤销
python3 csvfix.py . --undo csvfix_undo.json
```

## 实际输出长这样

```
文件                                          编码          状态
一月.csv                                      gb18030     待转换
二月.csv                                      gb18030     待转换

以上只是预览。加 --apply 才会真写盘。
```

加 `--apply` 之后：

```
一月.csv                                      gb18030     已转换 → _utf8/一月.csv（4 行）
二月.csv                                      gb18030     已转换 → _utf8/二月.csv（2 行）

共处理 2 个文件，撤销记录写入 csvfix_undo.json
```

## 参数

| 参数 | 作用 |
|---|---|
| `--apply` | 真写盘，不加就只预览 |
| `--inplace` | 原地覆盖，自动留 `.bak` |
| `--merge OUT.csv` | 把输入全部合并成 OUT.csv |
| `--dedupe` | 按整行去重（转换和合并都能用） |
| `--source` | 合并时追加「来源文件」列 |
| `--undo UNDO.json` | 按撤销记录回滚 |

## 几个常见问题

**转完了 Excel 还是乱码？**
确认你打开的是 `_utf8` 目录里的新文件，不是原文件。新文件开头有三个字节 `EF BB BF`，这就是 BOM，Excel 看到它才按 UTF-8 解析。

**WPS 需要吗？**
WPS 对编码更宽容，但转完同样更稳，尤其是要给同事传、要上传到各种后台系统的时候。

**表头不一致的文件能合并吗？**
能。脚本会以第一个文件的表头为准，表头对不上的文件整份当数据接在后面，不会丢行。

**数据里带逗号和引号会被拆坏吗？**
不会。走的是标准 csv 解析，带引号的字段会原样保留。

**会改我的原文件吗？**
默认不会，只写到 `_utf8` 子目录。只有你显式加 `--inplace` 才原地改，而且改前先备份。

## 环境

Python 3.7+，纯标准库，不联网、不上传任何文件，Mac / Windows / Linux 都能跑。

买断制，源码可随意修改自用。
