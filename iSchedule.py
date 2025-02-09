import os
import glob
import pandas as pd
from datetime import datetime, timedelta
import pytz
import re
from uuid import uuid4

# 配置参数
SKIP_ROWS = 2  # 跳过前两行（空行和表头）| usyd导出的xls文件默认从第三行开始
TIMEZONE = pytz.timezone("Australia/Sydney")
COLUMNS = [
    "Subject Code", "Description", "Group", "Activity",
    "Day", "Time", "Campus", "Location", "Duration", "Dates"
]


def parse_duration(duration_str):
    """解析持续时间（如'1 hr' → 60分钟）"""
    if "hr" in duration_str:
        return int(duration_str.split()[0]) * 60
    return 0


def parse_date(date_str, year):
    """解析日期字符串（如'24/2' → datetime对象）"""
    try:
        return datetime.strptime(f"{date_str}/{year}", "%d/%m/%Y")
    except ValueError:  # 处理跨年情况
        return datetime.strptime(f"{date_str}/{year + 1}", "%d/%m/%Y")


def standardize_address(address):
    """标准化地址"""
    parts = address.split('.')
    if len(parts) > 3:
        """防止因ABS Building改名影响地图定位"""
        if parts[3] == "Belinda Hutchinson Building":
            return "Abercrombie Building\, Sydney"
        else:
            return parts[3] + "\, Sydney"
    else:
        return None

def classroom_location(address):
    """提取教室位置"""
    parts = address.split('.')
    if len(parts) > 4:
        return parts[4]
    else:
        return "Online"


def create_ics_event(subject_code, group, start, end, description, until, location=None):
    """生成符合iOS标准的ICS事件"""
    event = [
        "BEGIN:VEVENT",
        f"UID:{uuid4()}@uni.sydney.edu.au",
        f"DESCRIPTION:{description}",
        f"DTSTART;TZID=Australia/Sydney:{start.strftime('%Y%m%dT%H%M%S')}",
        f"DTEND;TZID=Australia/Sydney:{end.strftime('%Y%m%dT%H%M%S')}",
        f"RRULE:FREQ=WEEKLY;BYDAY={start.strftime('%a').upper()[:2]};UNTIL={until.strftime('%Y%m%dT%H%M%S')}",
        f"SUMMARY:{subject_code}-{group}",
        "TRANSP:OPAQUE"
    ]

    # 如果存在有效地址，添加LOCATION字段
    if location:
        event.insert(3, f"LOCATION:{location}")

    event.append("END:VEVENT")
    return '\n'.join(event) + '\n'  # 确保每行有换行符


# ---------- 核心处理逻辑 ----------
def generate_ics_from_excel(excel_path, ics_path, year):
    """处理单个Excel文件"""
    try:
        df = pd.read_excel(
            excel_path,
            skiprows=SKIP_ROWS,
            names=COLUMNS,
            header=None,
            usecols=range(len(COLUMNS))
        ).dropna(subset=["Subject Code"])
    except Exception as e:
        print(f"文件 {os.path.basename(excel_path)} 读取失败: {str(e)}")
        return False

    ics_content = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Sydney University//Timetable//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:University Timetable",
        "X-WR-TIMEZONE:Australia/Sydney"
    ]

    for idx, row in df.iterrows():
        try:
            subject_code = re.match(r"^([A-Za-z]+\d+)", row["Subject Code"]).group(1)
            group = row["Group"].strip()
            time_str = f"{row['Time']}:00" if len(row['Time']) == 5 else row['Time']
            start_time = datetime.strptime(time_str, "%H:%M:%S").time()
            duration = parse_duration(row["Duration"])

            # 处理Location字段
            location = None
            raw_location = row["Location"].strip()
            if raw_location not in ["-", ""]:
                standardized_location = standardize_address(raw_location)
                if standardized_location:  # 仅当地址有效时赋值
                    location = standardized_location

            # 处理classroom
            classroom = None
            if raw_location not in ["-", ""]:
                classroom = classroom_location(raw_location)

            # 处理Description
            description = row["Description"]
            if classroom:
                description = f"{description} \\n{classroom}"

            # 处理日期范围
            for date_range in row["Dates"].split(","):
                date_range = date_range.strip()
                if "-" not in date_range:
                    continue

                start_str, end_str = date_range.split("-")
                start_date = parse_date(start_str.strip(), year)
                end_date = parse_date(end_str.strip(), year)

                # 生成事件时间
                event_start = TIMEZONE.localize(
                    datetime.combine(start_date.date(), start_time)
                )
                event_end = event_start + timedelta(minutes=duration)
                until = TIMEZONE.localize(
                    datetime.combine(end_date.date(), start_time)
                )

                # 添加到ICS内容
                ics_content.append(
                    create_ics_event(
                        subject_code=subject_code,
                        group=group,
                        start=event_start,
                        end=event_end,
                        description=description,
                        until=until,
                        location=location  # 传递标准化后的地址
                    )
                )

        except Exception as e:
            print(f"文件 {os.path.basename(excel_path)} 第{idx + SKIP_ROWS + 1}行错误: {str(e)}")
            continue

    ics_content.append("END:VCALENDAR")

    with open(ics_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(ics_content))
    return True


# ---------- 主程序 ----------
def main():
    year = datetime.now().year

    # 遍历所有Excel文件
    for excel_path in glob.glob("*.xls*"):
        if not (excel_path.endswith(".xls") or excel_path.endswith(".xlsx")):
            continue

        ics_path = f"{os.path.splitext(excel_path)[0]}.ics"
        print(f"\n{'=' * 40}")
        print(f"正在处理: {os.path.basename(excel_path)}")

        if generate_ics_from_excel(excel_path, ics_path, year):
            print(f"成功生成: {os.path.basename(ics_path)}")
        else:
            print("生成失败，请检查文件格式")


if __name__ == "__main__":
    main()