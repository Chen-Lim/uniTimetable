# uniTimetable — Rust Rewrite

> English version below.

📅 **这个版本是对原 Python 版本的 Rust 重写。不再需要导出 Excel 文件——只需提供你的 USyd iCal 订阅链接，程序会自动下载原始日历、解析并重构为干净的 ICS 文件，方便导入任意日历应用。**

## 功能特性

- **直接订阅链接输入**：通过 USyd 个人 iCal 链接自动下载课程表，无需手动导出文件。
- **智能重构事件**：
  - 将原始 ICS 中每节课的独立条目合并为带 `RRULE` 的周重复事件，大幅减少日历条目数量。
  - 自动识别学期中断（如复活节假期、期中假），将断开的课程段分别生成独立的重复规则。
  - 支持悉尼时区（`Australia/Sydney`），兼容全球日历应用。
- **地址标准化**：从原始点分隔格式中提取楼名和房间号，生成可读性更好的的地址形式（如 `Belinda Hutchinson Building, Sydney`）。
- **支持单次事件**：Presentation、考试等一次性课程不会被错误地添加重复规则。
- **高性能**：基于 Rust 构建，为后续部署为 Telegram Bot 服务做准备。

## 快速开始

### 环境要求

- [Rust](https://rustup.rs/) 1.70+（通过 `rustup` 安装即可）

### 编译
```bash
git clone https://github.com/your-repo/uniTimetable.git
cd uniTimetable
git checkout rust-rewrite
cargo build --release
```

### 使用方法

**第一步**：登录 [Sydney Timetable](https://timetable.sydney.edu.au)，在页面中找到你的个人 iCal 订阅链接，格式如下：
```
https://timetable.sydney.edu.au/even/rest/calendar/ical/your-key
```

**第二步**：运行程序
```bash
# 通过订阅链接（推荐）
./target/release/ischeduler --url "https://timetable.sydney.edu.au/even/rest/calendar/ical/your-key"

# 指定输出路径
./target/release/ischeduler --url "..." --output ~/Desktop/timetable.ics

# 使用本地已下载的 ICS 文件（调试用）
./target/release/ischeduler --file raw_timetable.ics

# 显示详细处理信息
./target/release/ischeduler --url "..." --verbose
```

**第三步**：将生成的 `timetable_clean.ics` 导入日历（Outlook、Apple Calendar、Google Calendar 均支持；iOS/iPadOS 建议通过邮件附件添加）。

### 项目结构
```
src/
├── main.rs          # CLI 入口
├── fetcher.rs       # 下载并验证 iCal 链接
├── parser.rs        # 解析原始 ICS（处理 RFC 5545 折行）
├── transformer.rs   # 核心重构逻辑
└── writer.rs        # 输出符合规范的 ICS 文件
```

### 开源协议

本项目采用 MIT 开源协议。欢迎提交 Issue 和 PR~

---

> English version here.

📅 This is a Rust rewrite of the original Python version. No Excel export needed — just provide your USyd iCal subscription URL, and the tool will automatically download, parse, and restructure your timetable into a clean ICS file ready to import into any calendar app.

## Features

- **Direct subscription URL input**: Automatically downloads your timetable via your personal USyd iCal link — no manual file export required.
- **Smart event restructuring**:
  - Collapses individual per-occurrence VEVENTs from the raw ICS into weekly `RRULE` recurring events, significantly reducing calendar clutter.
  - Automatically detects mid-semester breaks (Easter, Study Break, etc.) and splits affected courses into separate recurrence segments instead of incorrectly merging them.
  - Supports the Sydney timezone (`Australia/Sydney`), compatible with all major calendar apps worldwide.
- **Address standardisation**: Extracts building name and room number from the raw dot-separated location format into human-readable addresses (e.g. `Belinda Hutchinson Building, Sydney`).
- **Single-occurrence support**: One-off events such as Presentations are kept as single events without incorrect recurrence rules.
- **High performance**: Built in Rust, designed for future deployment as a Telegram Bot service.

## Quick Start

### Requirements

- [Rust](https://rustup.rs/) 1.70+ (install via `rustup`)

### Build
```bash
git clone https://github.com/your-repo/uniTimetable.git
cd uniTimetable
git checkout rust-rewrite
cargo build --release
```

### How to Use

**Step 1**: Log in to [Sydney Timetable](https://timetable.sydney.edu.au) and find your personal iCal subscription URL:
```
https://timetable.sydney.edu.au/even/rest/calendar/ical/your-personal-key
```

**Step 2**: Run the tool
```bash
# Via subscription URL (recommended)
./target/release/ischeduler --url "https://timetable.sydney.edu.au/even/rest/calendar/ical/your-key"

# Specify output path
./target/release/ischeduler --url "..." --output ~/Desktop/timetable.ics

# Use a locally downloaded ICS file (for debugging)
./target/release/ischeduler --file raw_timetable.ics

# Show detailed processing info
./target/release/ischeduler --url "..." --verbose
```

**Step 3**: Import the generated `timetable_clean.ics` into your calendar app of choice. Outlook, Apple Calendar, and Google Calendar all support ICS import. For iOS/iPadOS, adding via email attachment is recommended.

### Project Structure
```
src/
├── main.rs          # CLI entry point
├── fetcher.rs       # Downloads and validates the iCal URL
├── parser.rs        # Parses raw ICS (handles RFC 5545 line folding)
├── transformer.rs   # Core restructuring logic
└── writer.rs        # Outputs a spec-compliant ICS file
```

### Open Source License

This project is licensed under the MIT License. Feel free to submit Issues and PRs!

⏰ Make timetable management easier!
