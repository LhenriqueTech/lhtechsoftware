import os
import pdfplumber
import openpyxl
import re
from datetime import datetime, timedelta
from openpyxl.styles import Border, Side, Alignment, Font, PatternFill


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
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    header_cell.alignment = Alignment(horizontal='center', vertical='center')


def format_time_columns(sheet):
    for col in range(3, 10):  # Colunas C a I
        col_letter = openpyxl.utils.get_column_letter(col)
        for cell in sheet[col_letter]:
            cell.number_format = "HH:MM:SS"
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


def add_hour_statement(sheet, saldo_anterior, horas_positivas, horas_negativas, pg_em_folha, saldo_final):
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )

    sheet['A40'] = "Extrato hora"
    sheet['B40'] = "Horas Positivas"
    sheet['C40'] = "Horas Negativas"
    sheet['D40'] = "PG em Folha"
    sheet['E40'] = "Saldo"

    sheet['A41'] = "Saldo Anterior"
    sheet['B41'] = ""
    sheet['C41'] = 67.17
    sheet['D41'] = ""
    sheet['E41'] = 67.17

    for cell in ('A40', 'A41', 'B40', 'C40', 'C41', 'D40', 'E40', 'E41'):
        sheet[cell].font = Font(bold=True)
        sheet[cell].border = thin_border

    sheet['A42'] = "Jan-25"
    sheet['B42'] = horas_positivas
    sheet['C42'] = horas_negativas
    sheet['D42'] = ""
    sheet['E42'] = 2.35

    for cell in ('A42', 'B42', 'C42', 'D42', 'E42'):
        sheet[cell].font = Font(bold=True)
        sheet[cell].border = thin_border

    sheet['A44'] = "Saldo final"
    sheet['B44'] = ""
    sheet['C44'] = ""
    sheet['D44'] = ""
    sheet['E44'] = saldo_final

    for cell in ('A43', 'B43', 'C43', 'D43', 'E43', 'A44', 'B44', 'C44', 'D44', 'E44'):
        sheet[cell].font = Font(bold=True)
        sheet[cell].border = thin_border


# ---------------------------------------------------------------------------
# Função principal de processamento (importável pelo app)
# ---------------------------------------------------------------------------

def process_pdfs(pdf_folder, output_folder, start_date, end_date, cpf_dict, progress_callback=None):
    """
    Processa todos os PDFs na pasta pdf_folder e gera planilhas Excel em output_folder.

    Args:
        pdf_folder (str): Caminho da pasta com os PDFs de ponto.
        output_folder (str): Caminho da pasta onde os .xlsx serão salvos.
        start_date (str): Data inicial no formato DD/MM/AAAA.
        end_date (str): Data final no formato DD/MM/AAAA.
        cpf_dict (dict): Dicionário {nome_minúsculo: "CPF: XXX.XXX.XXX-XX"}.
        progress_callback (callable, optional): Função (pct: int, msg: str) para reportar progresso.

    Returns:
        list[str]: Caminhos dos arquivos Excel gerados.
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

        headers = [
            'Data', 'Dia da Semana',
            'Entrada (Manhã)', 'Saída (Manhã)',
            'Entrada (Tarde)', 'Saída (Tarde)',
            'Horas de Trabalho', 'Horas de Falta', 'Horas Extras',
        ]
        ws.append(headers)
        ws['G2'] = 'Horas de Trabalho'

        total_work_hours = timedelta(0)
        total_absence_hours = timedelta(0)
        total_extra_hours = timedelta(0)

        for i, match in enumerate(filtered_data, start=3):
            date_str = match[0]
            date_obj = datetime.strptime(date_str, "%d/%m/%Y")
            day_of_week = translate_day_of_week(date_obj.strftime("%A"))

            work_hours = time_str_to_timedelta(match[5])
            absence_hours = time_str_to_timedelta(match[6])
            extra_hours = time_str_to_timedelta(match[7])

            total_work_hours += work_hours
            total_absence_hours += absence_hours
            total_extra_hours += extra_hours

            ws.append([date_str, day_of_week] + [f if f else "" for f in match[1:]])
            ws[f'G{i}'] = f'=F{i}-C{i}-(E{i}-D{i})'

        ws.append([])
        ws.append([
            "Total", "", "", "", "", "",
            timedelta_to_time_str(total_work_hours),
            timedelta_to_time_str(total_absence_hours),
            timedelta_to_time_str(total_extra_hours),
        ])

        column_widths = [15, 20, 20, 20, 20, 20, 20, 20, 20]
        for i, width in enumerate(column_widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

        format_time_columns(ws)

        highlight_cells(ws, ['A1'], color="929292")
        highlight_cells(ws, ['J1'], color="929292")
        highlight_cells(ws, ['G37', 'H37', 'I37'], color="FFFFFF")

        for row in ws.iter_rows(min_col=2, max_col=2):
            for cell in row:
                if cell.value in ('Sábado', 'Domingo') and any(
                    cell.offset(column=i).value == '' for i in range(1, 9)
                ):
                    highlight_red(ws, [cell.coordinate] + [cell.offset(column=i).coordinate for i in range(1, 9)])

        apply_borders(ws)
        align_all_left(ws)
        center_align_A1(ws)
        add_empresa_info(ws, empresa_nome, empresa_cnpj)
        fill_G42_from_A1(ws)
        fill_G43_with_cpf(ws, cpf_dict)

        saldo_anterior = 67.17
        horas_positivas = 0.13
        horas_negativas = 2.48
        pg_em_folha = 2.35
        saldo_final = 69.52
        add_hour_statement(ws, saldo_anterior, horas_positivas, horas_negativas, pg_em_folha, saldo_final)

        apply_font_to_cells(ws)

        excel_filename = f'Relatório de Ponto {person_name}.xlsx'
        output_path = os.path.join(output_folder, excel_filename)
        wb.save(output_path)
        generated.append(output_path)

        if progress_callback:
            progress_callback(int(idx / total * 100), f"Processado: {pdf_file}")

    return generated
