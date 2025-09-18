import re
from io import StringIO
from pathlib import Path

import pandas as pd
import pytz

# 常量定义
INPUT_PATH = Path.home() / "Downloads/(日志)表格视图.xlsx"
OUTPUT_PATH = Path.home() / "Downloads/tqsl.adi"
LOCAL_TIMEZONE = pytz.timezone('Asia/Shanghai')

# 字段映射
FIELD_MAPPING = {
    'CALL': '呼号',
    'MODE': '模式',
    'RST_RCVD': '对方信号',
    'RST_SENT': '己方信号',
    'RIG': '对方设备',
    'ANTENNA': '对方天馈',
    'RX_PWR': '对方功率',
    'QTH': '对方QTH',
    'MY_RIG': '己方设备',
    'MY_ANTENNA': '己方天馈',
    'TX_PWR': '己方功率',
    'MY_CITY': '己方QTH',
    'COMMENTS': '补充'
}

# 频率正则
FREQUENCY_PATTERN = re.compile(r"(\d+\.\d+)(-\d+)?")


def write_adif_field(output: StringIO, field: str, data) -> None:
    """写入ADIF格式的字段"""
    if pd.notna(data):
        s = str(data).strip()
        if '\n' in s:
            raise ValueError(f"字段 '{field}' 包含换行符: {s!r}")
        output.write(f"<{field.upper()}:{len(s)}>{s}\n")


def generate_adif(df: pd.DataFrame) -> str:
    """生成ADIF格式的日志内容"""
    output = StringIO()

    for _, row in df.iterrows():
        # 处理时间字段
        local_dt = row['时间'].to_pydatetime()
        if not local_dt.tzinfo:
            local_dt = LOCAL_TIMEZONE.localize(local_dt)
        utc_dt = local_dt.astimezone(pytz.utc)

        write_adif_field(output, 'QSO_DATE', utc_dt.strftime('%Y%m%d'))
        write_adif_field(output, 'TIME_ON', utc_dt.strftime('%H%M'))

        # 处理频率字段
        freq_str = str(row['频率'])
        if match := FREQUENCY_PATTERN.match(freq_str):
            primary = float(match.group(1))
            if offset := match.group(2):
                # 处理频差
                write_adif_field(output, 'FREQ', primary + float(offset))
                write_adif_field(output, 'FREQ_RX', primary)
            else:
                write_adif_field(output, 'FREQ', primary)

        # 写入映射字段
        for adif_field, col_name in FIELD_MAPPING.items():
            write_adif_field(output, adif_field, row.get(col_name))

        # 记录结束标记
        output.write("<EOR>\n\n")

    return output.getvalue()


def main():
    try:
        # 读取Excel文件
        df = pd.read_excel(INPUT_PATH)

        # 生成ADIF内容
        adif_content = generate_adif(df)

        # 写入文件
        OUTPUT_PATH.write_text(adif_content, encoding='utf-8')
        print(f"成功生成ADIF文件: {OUTPUT_PATH}")

    except Exception as e:
        print(f"处理过程中出错: {e}")
        raise


if __name__ == "__main__":
    main()
