from flask import Flask, render_template, request, redirect, url_for, send_file
from peewee import Model, SqliteDatabase, CharField, DateTimeField, TextField, IntegerField
from datetime import datetime
import xlsxwriter

app = Flask(__name__)

# Definindo o banco de dados SQLite
db = SqliteDatabase('atendimentos.db')


# Definindo o modelo de dados
class Cliente(Model):
    codigo = CharField()
    nome = CharField()
    responsavel = CharField()
    telefone = CharField()
    numero_maquinas = IntegerField()
    valor_mensalidade = IntegerField()
    acessos_maquinas = TextField()

    class Meta:
        database = db


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
db.create_tables([Cliente, Atendimento])


# Rota para listar os clientes
@app.route('/clientes')
def listar_clientes():
    clientes = Cliente.select()
    return render_template('lista_clientes.html', clientes=clientes)


# Rota para adicionar um novo cliente
@app.route('/novo_cliente', methods=['GET', 'POST'])
def novo_cliente():
    if request.method == 'POST':
        codigo = request.form['codigo']
        nome = request.form['nome']
        responsavel = request.form['responsavel']
        telefone = request.form['telefone']
        numero_maquinas = request.form['numero_maquinas']
        valor_mensalidade = request.form['valor_mensalidade']
        acessos_maquinas = request.form['acessos_maquinas']
        Cliente.create(
            codigo=codigo,
            nome=nome,
            responsavel=responsavel,
            telefone=telefone,
            numero_maquinas=numero_maquinas,
            valor_mensalidade=valor_mensalidade,
            acessos_maquinas=acessos_maquinas
        )
        return redirect(url_for('listar_clientes'))
    return render_template('novo_cliente.html')


# Rota para editar um cliente
@app.route('/editar_cliente/<int:id>', methods=['GET', 'POST'])
def editar_cliente(id):
    cliente = Cliente.get_or_none(id=id)
    if cliente:
        if request.method == 'POST':
            cliente.nome = request.form['nome']
            cliente.email = request.form['email']
            cliente.save()
            return redirect(url_for('listar_clientes'))
        return render_template('editar_cliente.html', cliente=cliente)
    else:
        return 'Cliente não encontrado', 404



# Rota para excluir um cliente
@app.route('/excluir_cliente/<int:id>')
def excluir_cliente(id):
    cliente = Cliente.get_or_none(id=id)
    if cliente:
        cliente.delete_instance()
        return redirect(url_for('listar_clientes'))
    else:
        return 'Cliente não encontrado', 404


@app.route('/')
def pagina_inicial():
    return render_template('menu.html')

# Rota para listar os atendimentos
@app.route('/lista_atendimentos')
def listar_atendimentos():
    atendimentos = Atendimento.select()
    return render_template('lista_atendimentos.html', atendimentos=atendimentos)


# Rota para adicionar um novo atendimento
@app.route('/novo_atendimento', methods=['GET', 'POST'])
def novo_atendimento():
    clientes = Cliente.select()
    colaboradores = ['Colaborador 1', 'Colaborador 2',
                     'Colaborador 3']
    if request.method == 'POST':
        empresa = request.form['empresa']
        cliente = request.form.get('cliente')
        atendente = request.form['atendente']
        observacao = request.form['observacao']
        Atendimento.create(
            empresa=empresa,
            cliente=cliente,
            atendente=atendente,
            observacao=observacao
        )
        return redirect(url_for('listar_atendimentos'))
    return render_template('novo_atendimento.html', clientes=clientes, colaboradores=colaboradores)


# Rota para visualizar a observação completa de um atendimento
@app.route('/visualizar_observacao/<int:id>')
def visualizar_observacao(id):
    atendimento = Atendimento.get_or_none(id=id)
    if atendimento:
        return render_template('visualizar_observacao.html', atendimento=atendimento)
    else:
        return 'Atendimento não encontrado', 404

# Rota para buscar atendimentos
@app.route('/buscar_atendimentos')
def buscar_atendimentos():
    query = request.args.get('query')
    atendimentos = Atendimento.select().where(Atendimento.empresa.contains(query) | Atendimento.cliente.contains(query) | Atendimento.atendente.contains(query))
    return render_template('lista_atendimentos.html', atendimentos=atendimentos)


# Rota para buscar clientes
@app.route('/buscar_clientes')
def buscar_clientes():
    query = request.args.get('query')
    clientes = Cliente.select().where(Cliente.nome.contains(query))
    return render_template('lista_clientes.html', clientes=clientes)

@app.route('/baixar_clientes_xlsx')
def baixar_clientes_xlsx():
    # Dados dos clientes
    clientes = Cliente.select()
    clientes_data = [(cliente.nome, cliente.telefone) for cliente in clientes]

    # Criar o arquivo XLSX para clientes
    output_path = 'clientes.xlsx'
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('Clientes')

    # Escrever dados dos clientes
    for row, data in enumerate(clientes_data):
        for col, value in enumerate(data):
            worksheet.write(row, col, value)

    workbook.close()

    # Enviar o arquivo de clientes para download
    return send_file(output_path, as_attachment=True)

@app.route('/baixar_atendimentos_xlsx')
def baixar_atendimentos_xlsx():
    # Dados dos atendimentos
    atendimentos = Atendimento.select()
    atendimentos_data = [(atendimento.empresa, atendimento.cliente, atendimento.atendente, atendimento.observacao) for atendimento in atendimentos]

    # Criar o arquivo XLSX para atendimentos
    output_path = 'atendimentos.xlsx'
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('Atendimentos')

    # Escrever dados dos atendimentos
    for row, data in enumerate(atendimentos_data):
        for col, value in enumerate(data):
            worksheet.write(row, col, value)

    workbook.close()

    # Enviar o arquivo de atendimentos para download
    return send_file(output_path, as_attachment=True)




if __name__ == '__main__':
    app.run(debug=True)
