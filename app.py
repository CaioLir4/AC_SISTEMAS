from flask import Flask, render_template, request, redirect, url_for, send_file
from peewee import Model, SqliteDatabase, CharField, DateTimeField, TextField
from datetime import datetime
import os
import xlsxwriter

app = Flask(__name__)

# Definindo o banco de dados SQLite
db = SqliteDatabase('atendimentos.db')


# Definindo o modelo de dados
class Atendimento(Model):
    empresa = CharField()
    cliente = CharField()
    atendente = CharField()
    observacao = TextField()
    data_hora = DateTimeField(default=datetime.now)

    class Meta:
        database = db


# Criando as tabelas se não existirem
db.connect()
db.create_tables([Atendimento])


# Rota para listar os atendimentos
@app.route('/')
def listar_atendimentos():
    atendimentos = Atendimento.select()
    return render_template('lista_atendimentos.html', atendimentos=atendimentos)


# Rota para adicionar um novo atendimento
@app.route('/novo_atendimento', methods=['GET', 'POST'])
def novo_atendimento():
    colaboradores = ['Colaborador 1', 'Colaborador 2',
                     'Colaborador 3']  # Substitua com os nomes reais dos colaboradores
    if request.method == 'POST':
        empresa = request.form['empresa']
        cliente = request.form['cliente']
        atendente = request.form['atendente']
        observacao = request.form['observacao']
        Atendimento.create(empresa=empresa, cliente=cliente, atendente=atendente, observacao=observacao)
        return redirect(url_for('listar_atendimentos'))
    return render_template('novo_atendimento.html', colaboradores=colaboradores)


# Rota para visualizar a observação completa de um atendimento
@app.route('/visualizar_observacao/<int:id>')
def visualizar_observacao(id):
    atendimento = Atendimento.get_or_none(id=id)
    if atendimento:
        return render_template('visualizar_observacao.html', atendimento=atendimento)
    else:
        return 'Atendimento não encontrado', 404


# Rota para baixar o banco de dados em formato xlsx
@app.route('/baixar_xlsx')
def baixar_xlsx():
    atendimentos = Atendimento.select()
    output_path = 'atendimentos.xlsx'
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet()

    # Escrevendo cabeçalhos
    headers = ['Empresa', 'Cliente', 'Atendente', 'Observação', 'Data e Hora']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Escrevendo os dados
    for row, atendimento in enumerate(atendimentos, start=1):
        worksheet.write(row, 0, atendimento.empresa)
        worksheet.write(row, 1, atendimento.cliente)
        worksheet.write(row, 2, atendimento.atendente)
        worksheet.write(row, 3, atendimento.observacao)
        worksheet.write(row, 4, atendimento.data_hora.strftime('%d/%m/%Y %H:%M:%S'))

    workbook.close()
    return send_file(output_path, as_attachment=True)

# Rota para baixar o banco de dados em formato XML
@app.route('/baixar_xml')
def baixar_xml():
    atendimentos = Atendimento.select()
    xml_data = '<?xml version="1.0" encoding="UTF-8" ?>\n<atendimentos>\n'
    for atendimento in atendimentos:
        xml_data += f'    <atendimento>\n        <empresa>{atendimento.empresa}</empresa>\n        <cliente>{atendimento.cliente}</cliente>\n        <atendente>{atendimento.atendente}</atendente>\n        <observacao>{atendimento.observacao}</observacao>\n        <data_hora>{atendimento.data_hora.strftime("%d/%m/%Y %H:%M:%S")}</data_hora>\n    </atendimento>\n'
    xml_data += '</atendimentos>'
    response = app.response_class(xml_data, mimetype='application/xml')
    response.headers.set('Content-Disposition', 'attachment', filename='atendimentos.xml')
    return response

# Rota para baixar o banco de dados em formato CSV
@app.route('/baixar_csv')
def baixar_csv():
    atendimentos = Atendimento.select()
    csv_data = 'Empresa,Cliente,Atendente,Observação,Data e Hora\n'
    for atendimento in atendimentos:
        csv_data += f'{atendimento.empresa},{atendimento.cliente},{atendimento.atendente},"{atendimento.observacao}",{atendimento.data_hora.strftime("%d/%m/%Y %H:%M:%S")}\n'
    response = app.response_class(csv_data, mimetype='text/csv')
    response.headers.set('Content-Disposition', 'attachment', filename='atendimentos.csv')
    return response

# Rota para baixar o banco de dados em formato TXT
@app.route('/baixar_txt')
def baixar_txt():
    atendimentos = Atendimento.select()
    txt_data = 'Empresa | Cliente | Atendente | Observação | Data e Hora\n'
    for atendimento in atendimentos:
        txt_data += f'{atendimento.empresa} | {atendimento.cliente} | {atendimento.atendente} | {atendimento.observacao} | {atendimento.data_hora.strftime("%d/%m/%Y %H:%M:%S")}\n'
    response = app.response_class(txt_data, mimetype='text/plain')
    response.headers.set('Content-Disposition', 'attachment', filename='atendimentos.txt')
    return response

# Rota para editar um atendimento
@app.route('/editar_atendimento/<int:id>', methods=['GET', 'POST'])
def editar_atendimento(id):
    atendimento = Atendimento.get_or_none(id=id)
    if atendimento:
        if request.method == 'POST':
            atendimento.empresa = request.form['empresa']
            atendimento.cliente = request.form['cliente']
            atendimento.atendente = request.form['atendente']
            atendimento.observacao = request.form['observacao']
            atendimento.save()
            return redirect(url_for('listar_atendimentos'))
        return render_template('editar_atendimento.html', atendimento=atendimento)
    else:
        return 'Atendimento não encontrado', 404

# Rota para excluir um atendimento
@app.route('/excluir_atendimento/<int:id>')
def excluir_atendimento(id):
    atendimento = Atendimento.get_or_none(id=id)
    if atendimento:
        atendimento.delete_instance()
        return redirect(url_for('listar_atendimentos'))
    else:
        return 'Atendimento não encontrado', 404



if __name__ == '__main__':
    app.run(debug=True)
