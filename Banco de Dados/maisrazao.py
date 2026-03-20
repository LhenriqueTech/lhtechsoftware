import os
import pdfplumber
import openpyxl
import re
from datetime import datetime, timedelta
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill


# ---------------------------------------------------------------------------
# Constante: jornada semanal fixa = 08:48:00  →  fração de dia Excel
# ---------------------------------------------------------------------------
JORNADA_SEMANAL_STR = "08:48:00"
_h, _m, _s = 8, 48, 0
JORNADA_SEMANAL_FRACTION = (_h * 3600 + _m * 60 + _s) / 86400.0


# ---------------------------------------------------------------------------
# Funções auxiliares
# ---------------------------------------------------------------------------

def is_within_date_range(date_str, start_date_str, end_date_str):
    date_format = "%d/%m/%Y"
    date = datetime.strptime(date_str, date_format)
    start_date = datetime.strptime(start_date_str, date_format)
    end_date = datetime.strptime(end_date_str, date_format)
    return start_date <= date <= end_date


def time_str_to_timedelta(time_str):
    if time_str:
        h, m, s = map(int, time_str.split(':'))
        return timedelta(hours=h, minutes=m, seconds=s)
    return timedelta(0)


def timedelta_to_time_str(td):
    total_seconds = int(td.total_seconds())
    h, remainder = divmod(total_seconds, 3600)
    m, s = divmod(remainder, 60)
    return f"{h:02}:{m:02}:{s:02}"


def extract_data_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "".join(page.extract_text() for page in pdf.pages)

    pattern = re.compile(
        r'(\d{2}/\d{2}/\d{4})\s*'
        r'(\d{2}:\d{2}:\d{2})?\s*'
        r'(\d{2}:\d{2}:\d{2})?\s*'
        r'(\d{2}:\d{2}:\d{2})?\s*'
        r'(\d{2}:\d{2}:\d{2})?\s*'
        r'(\d{2}:\d{2}:\d{2})?\s*'
        r'(\d{2}:\d{2}:\d{2})?\s*'
        r'(\d{2}:\d{2}:\d{2})?'
    )
    return pattern.findall(text)


def apply_borders(sheet):
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )
    for row in sheet.iter_rows():
        for cell in row:
            cell.border = thin_border


def add_person_name(sheet, name):
    header_cell = sheet.cell(row=1, column=1, value=f"Relatório de Ponto {name}")
    header_cell.font = Font(bold=True, size=14)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
    header_cell.alignment = Alignment(horizontal='center', vertical='center')


def format_time_columns(sheet):
    # Colunas C a J (índices 3..10)
    for col in range(3, 11):
        col_letter = openpyxl.utils.get_column_letter(col)
        for cell in sheet[col_letter]:
            cell.number_format = "[h]:mm:ss"
            cell.alignment = Alignment(horizontal='left', vertical='center')


def highlight_cells(sheet, cells, color="929292"):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for cell in cells:
        sheet[cell].fill = fill
        sheet[cell].alignment = Alignment(horizontal='left', vertical='center')


def highlight_red(sheet, cells):
    gray_fill = PatternFill(start_color="CBCBCB", end_color="CBCBCB", fill_type="solid")
    for cell in cells:
        sheet[cell].fill = gray_fill
        sheet[cell].alignment = Alignment(horizontal='left', vertical='center')


def align_all_left(sheet):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.coordinate != 'A1':
                cell.alignment = Alignment(horizontal='left', vertical='center')


def center_align_A1(sheet):
    sheet['A1'].alignment = Alignment(horizontal='center', vertical='center')


def fill_G42_from_A1(sheet):
    nome_a1 = sheet['A1'].value
    nome = nome_a1.replace("Relatório de Ponto ", "").replace("012025", "").strip().lower()
    sheet['G42'] = nome.upper()
    sheet['G42'].alignment = Alignment(horizontal='left', vertical='center')


def fill_G43_with_cpf(sheet, cpf_dict):
    person_name = sheet['G42'].value.strip().lower()
    cpf = cpf_dict.get(person_name, "CPF: 000.000.000-00")
    sheet['G43'] = cpf
    sheet['G43'].alignment = Alignment(horizontal='left', vertical='center')


day_translation = {
    "Monday": "Segunda-feira",
    "Tuesday": "Terça-feira",
    "Wednesday": "Quarta-feira",
    "Thursday": "Quinta-feira",
    "Friday": "Sexta-feira",
    "Saturday": "Sábado",
    "Sunday": "Domingo",
}


def translate_day_of_week(day_of_week):
    return day_translation.get(day_of_week, day_of_week)


empresa_nome = "MR ORGANIZAÇÃO CONTABIL LTDA"
empresa_cnpj = "CNPJ: 19.331.844/0001-43"


def add_empresa_info(sheet, nome_empresa, cnpj):
    sheet['G47'] = f" {nome_empresa}"
    sheet['G48'] = f" {cnpj}"
    sheet['G47'].alignment = Alignment(horizontal='left', vertical='center')
    sheet['G48'].alignment = Alignment(horizontal='left', vertical='center')


def apply_font_to_cells(sheet):
    font = Font(name='Times New Roman', size=10)
    for cell_ref in ('G47', 'G48', 'G43', 'G42'):
        sheet[cell_ref].font = font


# ---------------------------------------------------------------------------
# Extrato de hora (saldo) — gerado dinamicamente para o ano corrente
# ---------------------------------------------------------------------------

_PT_ABBR = ["jan", "fev", "mar", "abr", "mai", "jun",
            "jul", "ago", "set", "out", "nov", "dez"]


def add_hour_statement(sheet):
    """
    Preenche o bloco de extrato de hora a partir da linha 40.

    Gera automaticamente 12 linhas para jan-YY até dez-YY do ANO CORRENTE
    (sem linha de saldo anterior). Assim, ao virar o ano o relatório
    sempre exibe o intervalo correto.
    """
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )

    ano_corrente = datetime.now().year
    ano_abbr = ano_corrente % 100  # ex.: 2026 → 26

    # ── Cabeçalho (linha 40) ──
    headers = {
        'A40': "Extrato hora",
        'B40': "Horas Positivas - B.H",
        'C40': "Horas Negativas",
        'D40': "PG em Folha",
        'E40': "Saldo",
    }
    for ref, val in headers.items():
        sheet[ref] = val
        sheet[ref].font = Font(bold=True)
        sheet[ref].border = thin_border
        sheet[ref].alignment = Alignment(horizontal='left', vertical='center')

    # ── 12 meses do ano corrente (linhas 41–52) ──
    for i, abbr in enumerate(_PT_ABBR):
        row = 41 + i
        label = f"{abbr}-{ano_abbr:02d}"
        for col_letter, value in zip(('A', 'B', 'C', 'D', 'E'),
                                     (label, "", "", "", "")):
            ref = f"{col_letter}{row}"
            sheet[ref] = value
            sheet[ref].font = Font(bold=True)
            sheet[ref].border = thin_border
            sheet[ref].alignment = Alignment(horizontal='left', vertical='center')

    # ── Saldo final (linha 53) ──
    saldo_row = 53
    for col_letter in ('A', 'B', 'C', 'D', 'E'):
        ref = f"{col_letter}{saldo_row}"
        sheet[ref] = ""
        sheet[ref].font = Font(bold=True)
        sheet[ref].border = thin_border
        sheet[ref].alignment = Alignment(horizontal='left', vertical='center')
    sheet[f'A{saldo_row}'] = "Saldo final"


# ---------------------------------------------------------------------------
# Função principal de processamento (importável pelo app)
# ---------------------------------------------------------------------------

def process_pdfs(pdf_folder, output_folder, start_date, end_date, cpf_dict, progress_callback=None):
    """
    Processa todos os PDFs na pasta pdf_folder e gera planilhas Excel em output_folder.

    Colunas geradas por linha de dia:
      A  Data
      B  Dia da Semana
      C  Entrada (Manhã)
      D  Saída (Manhã)
      E  Entrada (Tarde)
      F  Saída (Tarde)
      G  Horas de Trabalho  = F-C-(E-D)
      H  H. Semanal         = 08:48:00 (fixo)
      I  Horas Extras       = IF(G>H, G-H, 0)
      J  Horas Negativas    = IF(G<H, H-G, 0)
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]
    total = len(pdf_files)
    generated = []

    for idx, pdf_file in enumerate(pdf_files, start=1):
        pdf_path = os.path.join(pdf_folder, pdf_file)
        person_name = os.path.splitext(pdf_file)[0].replace('Relatório de Ponto - ', '')

        data = extract_data_from_pdf(pdf_path)
        filtered_data = [m for m in data if is_within_date_range(m[0], start_date, end_date)]

        wb = openpyxl.Workbook()
        ws = wb.active

        add_person_name(ws, person_name)

        # ── Cabeçalhos (linha 2) ──
        headers = [
            'Data', 'Dia da Semana',
            'Entrada (Manhã)', 'Saída (Manhã)',
            'Entrada (Tarde)', 'Saída (Tarde)',
            'Horas de Trabalho', 'H. Semanal',
            'Horas Extras', 'Horas Negativas',
        ]
        ws.append(headers)

        total_work_hours = timedelta(0)
        total_extra_hours = timedelta(0)
        total_neg_hours = timedelta(0)

        for i, match in enumerate(filtered_data, start=3):
            date_str = match[0]
            date_obj = datetime.strptime(date_str, "%d/%m/%Y")
            day_of_week = translate_day_of_week(date_obj.strftime("%A"))

            work_hours = time_str_to_timedelta(match[5])
            total_work_hours += work_hours

            # Linha de dados
            row_data = [date_str, day_of_week] + [f if f else "" for f in match[1:7]]
            ws.append(row_data)

            # G = Horas de Trabalho  (formula)
            ws[f'G{i}'] = f'=F{i}-C{i}-(E{i}-D{i})'
            ws[f'G{i}'].number_format = "[h]:mm:ss"

            # H = H. Semanal  (valor fixo = 08:48:00 como fração de dia)
            ws[f'H{i}'] = JORNADA_SEMANAL_FRACTION
            ws[f'H{i}'].number_format = "[h]:mm:ss"

            # I = Horas Extras  =IF(G>H, G-H, 0)
            ws[f'I{i}'] = f'=IF(G{i}>H{i},G{i}-H{i},0)'
            ws[f'I{i}'].number_format = "[h]:mm:ss"

            # J = Horas Negativas  =IF(G<H, H-G, 0)
            ws[f'J{i}'] = f'=IF(G{i}<H{i},H{i}-G{i},0)'
            ws[f'J{i}'].number_format = "[h]:mm:ss"

        # ── Linha de totais ──
        ws.append([])
        total_row_idx = len(filtered_data) + 4
        ws.append([
            "Total", "", "", "", "", "",
            timedelta_to_time_str(total_work_hours),
            JORNADA_SEMANAL_STR,
            "",
            "",
        ])

        # ── Larguras de coluna ──
        column_widths = [15, 20, 20, 20, 20, 20, 20, 20, 20, 20]
        for ci, width in enumerate(column_widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = width

        format_time_columns(ws)

        highlight_cells(ws, ['A1'], color="929292")
        highlight_cells(ws, ['K1'], color="929292")

        # Destaca fins de semana
        for row in ws.iter_rows(min_col=2, max_col=2):
            for cell in row:
                if cell.value in ('Sábado', 'Domingo') and any(
                    cell.offset(column=i).value == '' for i in range(1, 10)
                ):
                    highlight_red(ws, [cell.coordinate] + [cell.offset(column=i).coordinate for i in range(1, 10)])

        apply_borders(ws)
        align_all_left(ws)
        center_align_A1(ws)
        add_empresa_info(ws, empresa_nome, empresa_cnpj)
        fill_G42_from_A1(ws)
        fill_G43_with_cpf(ws, cpf_dict)

        # Extrato de hora (saldo dinâmico)
        add_hour_statement(ws)

        apply_font_to_cells(ws)

        excel_filename = f'Relatório de Ponto {person_name}.xlsx'
        output_path = os.path.join(output_folder, excel_filename)
        wb.save(output_path)
        generated.append(output_path)

        if progress_callback:
            progress_callback(int(idx / total * 100), f"Processado: {pdf_file}")

    return generated
