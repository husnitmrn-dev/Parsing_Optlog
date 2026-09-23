import csv
import io
import re
from pathlib import Path
import polars as pl

HEADERS = [
    'NE Name',
    'Site ID',
    'Source',
    'Operator',
    'Domain',
    'Workstation',
    'Operate Type',
    'Operate Time',
    'Result',
    'Error Code',
    'Command Level',
    'End Time',
    'Operation Command Information',
]


def extract_site_id(ne_name):
    parts = ne_name.split('#') if ne_name else []
    return parts[1] if len(parts) > 2 else ''


def split_columns(line):
    return [value.strip().strip('"') for value in re.split(r'\t+|,{1}|\s{2,}', line.strip())]


def extract_identity_ne(text):
    plus_match = re.match(r'^\+\+\+\s+([^\s,"\t]+)', text)
    if plus_match and '#' in plus_match.group(1):
        return plus_match.group(1).strip('"')

    identity_match = re.search(
        r'([A-Za-z0-9_-]+#[^\s,"\t]+#[^\s,"\t]+)'
        r'.*?\d{1,2}[-/]\d{1,2}[-/]\d{4}'
        r'.*?\d{1,2}:\d{2}:\d{2}',
        text,
    )
    return identity_match.group(1).strip('"') if identity_match else ''


def parse_optlog_csv(raw_text):
    rows = csv.reader(raw_text.splitlines())
    parsed_rows = []
    current_ne = ''
    pending_command = ''

    for csv_row in rows:
        values = [value.replace('\xa0', ' ').strip() for value in csv_row]
        if not values:
            continue

        ne_value = next((value for value in values if value.count('#') >= 2), '')
        if ne_value:
            current_ne = ne_value

        if values[0] == 'Operation Command Information:':
            pending_command = values[1] if len(values) > 1 else ''
            continue

        if values[0] != 'EMS' or len(values) < 9 or 'MAINTENANCE' not in values:
            continue

        maintenance_index = values.index('MAINTENANCE')
        row = [
            current_ne,
            extract_site_id(current_ne),
            values[0],
            values[1] if len(values) > 1 else '',
            values[2] if len(values) > 2 else '',
            values[3] if len(values) > 3 else '',
            values[maintenance_index],
            values[maintenance_index + 1] if len(values) > maintenance_index + 1 else '',
            values[maintenance_index + 2] if len(values) > maintenance_index + 2 else '',
            values[maintenance_index + 3] if len(values) > maintenance_index + 3 else '',
            values[maintenance_index + 4] if len(values) > maintenance_index + 4 else '',
            values[maintenance_index + 5] if len(values) > maintenance_index + 5 else '',
            pending_command,
        ]
        parsed_rows.append(row)

    return parsed_rows


def parse_optlog_txt(raw_text):
    raw_text = raw_text.replace('\xa0', ' ')
    lines = raw_text.splitlines()
    rows = []
    current_ne = ''
    pending_command = ''

    for line in lines:
        text = line.strip().strip('"').strip()
        if not text:
            continue

        identity_ne = extract_identity_ne(text)
        if identity_ne:
            current_ne = identity_ne
            pending_command = ''
            continue

        command_marker = 'Operation Command Information:'
        if command_marker in text:
            pending_command = text.split(command_marker, 1)[1].strip().strip('"')
            continue

        if not text.startswith('EMS'):
            continue

        columns = split_columns(text)
        if len(columns) < 9 or 'MAINTENANCE' not in columns:
            continue

        maintenance_index = columns.index('MAINTENANCE')
        row = [
            current_ne,
            extract_site_id(current_ne),
            columns[0],
            columns[1] if len(columns) > 1 else '',
            columns[2] if len(columns) > 2 else '',
            columns[3] if len(columns) > 3 else '',
            columns[maintenance_index],
            columns[maintenance_index + 1] if len(columns) > maintenance_index + 1 else '',
            columns[maintenance_index + 2] if len(columns) > maintenance_index + 2 else '',
            columns[maintenance_index + 3] if len(columns) > maintenance_index + 3 else '',
            columns[maintenance_index + 4] if len(columns) > maintenance_index + 4 else '',
            columns[maintenance_index + 5] if len(columns) > maintenance_index + 5 else '',
            pending_command,
        ]
        rows.append(row)

    return rows


def convert_optlog_file(input_path: str, output_path: str = None) -> str:
    """
    Mengonversi file OPTLOG (.txt/.csv) ke file Excel (.xlsx)
    """
    input_file = Path(input_path)
    raw_text = input_file.read_text(encoding='utf-8-sig', errors='ignore')

    if input_file.suffix.lower() == '.csv':
        rows = parse_optlog_csv(raw_text)
    else:
        rows = parse_optlog_txt(raw_text)

    if not rows:
        print(f"Tidak ada data valid ditemukan di {input_file.name}")
        return None

    df = pl.DataFrame(rows, schema=HEADERS, orient='row')

    if not output_path:
        output_dir = input_file.parent / 'Hasil_OPTLOG'
        output_dir.mkdir(exist_ok=True)
        output_path = output_dir / f"{input_file.stem}_OPTLOG.xlsx"

    df.write_excel(
        workbook=str(output_path),
        worksheet='OPTLOG_Data',
        table_style=None,
        header_format={'bold': True, 'text_wrap': False, 'valign': 'vcenter'},
        autofit=True,
        freeze_panes='A2',
    )

    print(f"File berhasil dikonversi ke Excel: {output_path}")
    return str(output_path)
