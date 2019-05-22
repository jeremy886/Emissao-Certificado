# coding=utf-8
import docx
import csv
import os
import subprocess

def make_certificate(filename, atributos):
    doc = docx.Document(filename)
    for p in doc.paragraphs:
        inline = p.runs
        for i in range(len(inline)):
            if 'name_participante' in inline[i].text:
                inline[i].text = inline[i].text.replace('name_participante', '')
                inline[i].text = atributos[0].title()
                inline[i].bold = True
            elif 'tipo_evento' in inline[i].text:
                inline[i].text = inline[i].text.replace('tipo_evento', '')
                inline[i].text = atributos[1].title()
                inline[i].bold = True
            elif 'nome_evento' in inline[i].text:
                inline[i].text = inline[i].text.replace('nome_evento', '')
                inline[i].text = atributos[2].title() + ' '
                inline[i].bold = True
            elif 'data' in inline[i].text:
                inline[i].text = inline[i].text.replace('data', '')
                inline[i].text = ' ' + atributos[3]+ ' '
                inline[i].bold = True
            elif 'localidade' in inline[i].text:
                inline[i].text = inline[i].text.replace('localidade', '')
                inline[i].text = atributos[4].title()
                inline[i].bold = True
            elif 'carga_horaria' in inline[i].text:
                inline[i].text = inline[i].text.replace('carga_horaria', '')
                inline[i].text = atributos[5] + ' horas'
                inline[i].bold = True
            elif 'codigo_certificado' in inline[i].text:
                inline[i].text = inline[i].text.replace('codigo_certificado', '')
                inline[i].text = atributos[6]
                inline[i].bold = True

    doc.save('certificados/{}.docx'.format(atributos[0]))
    convert_toPDF(atributos)


def convert_toPDF(atributos):
    subprocess.call('libreoffice --convert-to pdf "certificados/{}.docx"'.format(atributos[0]), shell=True, stdout=False)
    print ('Certificado emitido: {}.pdf'.format(atributos[0]))
    subprocess.call('mv "{}.pdf" certificados/'.format(atributos[0]), shell=True, stdout=False)
    remove_DOCX()


def remove_DOCX():
    subprocess.call('rm -r certificados/*.docx', shell=True, stdout=False)


def certificate(filename):
    with open(filename, 'r') as csv_file:
        attendents = csv.reader(csv_file, delimiter=',')

        for row in attendents:
            atributos = [row[0], row[1], row[2], row[3], row[4], row[5], row[6]]
            make_certificate('modelo.docx', atributos)


if __name__ == '__main__':
    certificate('participantes.csv')