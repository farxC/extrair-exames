"""Gera PDF fictício sem bordas e dados esperados para validar o extrator."""
import csv
from pathlib import Path
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4


def generate():
    directory = Path(__file__).resolve().parent / 'mock'
    directory.mkdir(exist_ok=True)
    pdf = canvas.Canvas(str(directory / 'exames_mock.pdf'), pagesize=A4)
    pdf.setTitle('Exames fictícios — teste de extração')
    positions = [30, 110, 270, 390, 510]

    def text(y, value):
        pdf.setFont('Helvetica', 10)
        pdf.drawString(30, y, value)

    def row(y, values):
        pdf.setFont('Helvetica', 9)
        for x, value in zip(positions, values):
            pdf.drawString(x, y, value)

    def header(y):
        row(y, ['DATA EXAME', 'PROCEDIMENTO', 'PRESTADOR', 'TIPO DO EXAME', 'VALOR'])

    text(800, 'MOCK — DADOS FICTÍCIOS, SEM VALIDADE')
    text(760, 'FUNCIONÁRIO: ANA EXEMPLO')
    text(742, 'CPF: 000.000.000-00')
    text(724, 'SETOR: ADMINISTRATIVO')
    text(706, 'FUNÇÃO: ANALISTA')
    header(675)
    row(653, ['01/11/2026', 'EXAME CLÍNICO', 'CLÍNICA ALFA', 'PERIÓDICO', '50,00'])
    row(629, ['02/11/2026', 'AVALIAÇÃO', 'CLÍNICA BETA', 'PERIÓDICO', '120,00'])
    row(615, ['', 'AUDIOLÓGICA COMPLETA', '', '', ''])
    row(588, ['TOTAL', '', '', '', '170,00'])

    text(530, 'FUNCIONÁRIO: BRUNO EXEMPLO')
    text(512, 'CPF: 111.111.111-11')
    text(494, 'SETOR: OPERAÇÕES')
    text(476, 'FUNÇÃO: TÉCNICO')
    header(445)
    row(423, ['03/11/2026', 'HEMOGRAMA', '', 'ADMISSIONAL', '35,50'])
    pdf.showPage()
    header(790)
    row(768, ['04/11/2026', 'EXAME CLÍNICO', 'CLÍNICA GAMA', 'ADMISSIONAL', '60,00'])
    row(744, ['TOTAL', '', '', '', '95,50'])
    pdf.save()

    fields = ['funcionário', 'cpf', 'setor', 'função', 'data exame',
              'procedimento', 'prestador', 'tipo do exame', 'valor']
    ana = ['ANA EXEMPLO', '000.000.000-00', 'ADMINISTRATIVO', 'ANALISTA']
    bruno = ['BRUNO EXEMPLO', '111.111.111-11', 'OPERAÇÕES', 'TÉCNICO']
    expected = [
        ana + ['01/11/2026', 'EXAME CLÍNICO', 'CLÍNICA ALFA', 'PERIÓDICO', '50,00'],
        ana + ['02/11/2026', 'AVALIAÇÃO AUDIOLÓGICA COMPLETA', 'CLÍNICA BETA', 'PERIÓDICO', '120,00'],
        bruno + ['03/11/2026', 'HEMOGRAMA', '', 'ADMISSIONAL', '35,50'],
        bruno + ['04/11/2026', 'EXAME CLÍNICO', 'CLÍNICA GAMA', 'ADMISSIONAL', '60,00'],
    ]
    with (directory / 'esperado.csv').open('w', encoding='utf-8-sig', newline='') as out:
        writer = csv.writer(out, delimiter=';')
        writer.writerows([fields] + expected)
    print(directory / 'exames_mock.pdf')


if __name__ == '__main__':
    generate()
