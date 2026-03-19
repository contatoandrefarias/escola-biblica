import os
from datetime import date, datetime
from flask import (Flask, render_template, request,
                   redirect, url_for, flash, send_file)
from flask_login import (LoginManager, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco, DATABASE
import sqlite3
import shutil
from functools import wraps

from auth import carregar_usuario, verificar_login

# Importações para PDF e DOC
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "escola_biblica_2026")

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
login_manager.login_message = "Faça login para continuar."
login_manager.login_message_category = "warning"

@login_manager.user_loader
def load_user(user_id):
    return carregar_usuario(user_id)

# Garante que o banco de dados seja inicializado/migrado ao iniciar o app
# A chamada já está aqui, mas vamos garantir que ela seja sempre executada
# e que a conexão seja fechada corretamente.
try:
    inicializar_banco()
    print("Inicialização do banco de dados tentada com sucesso ao iniciar o app.")
except Exception as e:
    print(f"ERRO CRÍTICO NA INICIALIZAÇÃO DO BANCO DE DADOS AO INICIAR O APP: {e}")


# Decorador para exigir perfil de administrador
def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not current_user.is_authenticated or current_user.perfil != 'admin':
            flash("Acesso negado. Apenas administradores podem acessar esta página.", "erro")
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


# ══════════════════════════════════════
# AUTH
# ══════════════════════════════════════
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))
    if request.method == "POST":
        email   = request.form.get("email", "").strip()
        senha   = request.form.get("senha", "")
        usuario = verificar_login(email, senha)
        if usuario:
            login_user(usuario, remember=True)
            flash(f"Bem-vindo(a), {usuario.nome}!", "sucesso")
            return redirect(
                request.args.get("next") or url_for("index"))
        flash("E-mail ou senha incorretos!", "erro")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    nome = current_user.nome
    logout_user()
    flash(f"Até logo, {nome}!", "sucesso")
    return redirect(url_for("login"))


# ══════════════════════════════════════
# PAINEL
# ══════════════════════════════════════
@app.route("/")
@login_required
def index():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) as t FROM alunos")
        total_alunos = cursor.fetchone()["t"]
        cursor.execute("SELECT COUNT(*) as t FROM professores")
        total_professores = cursor.fetchone()["t"]
        cursor.execute(
            "SELECT COUNT(*) as t FROM disciplinas WHERE ativa=1")
        total_disciplinas = cursor.fetchone()["t"]
        cursor.execute(
            "SELECT COUNT(*) as t FROM turmas WHERE ativa=1")
        total_turmas = cursor.fetchone()["t"]
        cursor.execute(
            "SELECT COUNT(*) as t FROM matriculas WHERE status='aprovado'")
        aprovados = cursor.fetchone()["t"]
        cursor.execute(
            "SELECT COUNT(*) as t FROM matriculas WHERE status='reprovado'")
        reprovados = cursor.fetchone()["t"]
        cursor.execute(
            "SELECT COUNT(*) as t FROM matriculas WHERE status='cursando'")
        cursando = cursor.fetchone()["t"]
        return render_template("index.html",
            total_alunos=total_alunos,
            total_professores=total_professores,
            total_disciplinas=total_disciplinas,
            total_turmas=total_turmas,
            aprovados=aprovados,
            reprovados=reprovados,
            cursando=cursando)
    except sqlite3.OperationalError as e:
        flash(f"Erro ao carregar dados do painel: {e}. O banco de dados pode estar sendo inicializado ou corrompido. Tente novamente em breve.", "erro")
        print(f"ERRO OPERACIONAL NO INDEX: {e}")
        # Retorna a página com contadores zerados para evitar o Erro 500
        return render_template("index.html",
            total_alunos=0, total_professores=0, total_disciplinas=0,
            total_turmas=0, aprovados=0, reprovados=0, cursando=0)
    except Exception as e:
        flash(f"Erro inesperado ao carregar o painel: {e}", "erro")
        print(f"ERRO INESPERADO NO INDEX: {e}")
        return render_template("index.html",
            total_alunos=0, total_professores=0, total_disciplinas=0,
            total_turmas=0, aprovados=0, reprovados=0, cursando=0)
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# TURMAS
# ══════════════════════════════════════
@app.route("/turmas")
@login_required
def turmas():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT t.*, COUNT(a.id) as total_alunos
            FROM turmas t
            LEFT JOIN alunos a ON a.turma_id = t.id
            GROUP BY t.id ORDER BY t.nome
        """)
        lista = cursor.fetchall()
        return render_template("turmas.html", turmas=lista)
    except Exception as e:
        flash(f"Erro ao carregar turmas: {e}", "erro")
        print(f"ERRO EM TURMAS: {e}")
        return render_template("turmas.html", turmas=[])
    finally:
        if conn:
            conn.close()


@app.route("/turmas/nova", methods=["GET", "POST"])
@login_required
def nova_turma():
    conn = None
    try:
        if request.method == "POST":
            nome         = request.form.get("nome", "").strip()
            descricao    = request.form.get("descricao", "").strip()
            faixa_etaria = request.form.get("faixa_etaria", "").strip()
            if not nome:
                flash("Nome é obrigatório!", "erro")
                return redirect(url_for("nova_turma"))
            conn   = conectar()
            cursor = conn.cursor()
            try:
                cursor.execute(
                    "INSERT INTO turmas (nome,descricao,faixa_etaria) VALUES (?,?,?)",
                    (nome, descricao, faixa_etaria))
                conn.commit()
                flash(f"Turma '{nome}' criada!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "turmas.nome" in str(e):
                    flash("Já existe uma turma com este nome!", "erro")
                else:
                    flash(f"Erro ao criar turma: {e}", "erro")
                print(f"ERRO AO CRIAR TURMA (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao criar turma: {e}", "erro")
                print(f"ERRO AO CRIAR TURMA: {e}")
            return redirect(url_for("turmas"))
        return render_template("nova_turma.html")
    except Exception as e:
        flash(f"Erro ao carregar página de nova turma: {e}", "erro")
        print(f"ERRO EM NOVA TURMA (GET): {e}")
        return redirect(url_for("turmas"))
    finally:
        if conn:
            conn.close()


@app.route("/turmas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_turma(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            nome         = request.form.get("nome", "").strip()
            descricao    = request.form.get("descricao", "").strip()
            faixa_etaria = request.form.get("faixa_etaria", "").strip()
            ativa        = 1 if request.form.get("ativa") == "on" else 0
            if not nome:
                flash("Nome é obrigatório!", "erro")
                return redirect(url_for("editar_turma", id=id))
            try:
                cursor.execute(
                    "UPDATE turmas SET nome=?, descricao=?, faixa_etaria=?, ativa=? WHERE id=?",
                    (nome, descricao, faixa_etaria, ativa, id))
                conn.commit()
                flash(f"Turma '{nome}' atualizada!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "turmas.nome" in str(e):
                    flash("Já existe uma turma com este nome!", "erro")
                else:
                    flash(f"Erro ao atualizar turma: {e}", "erro")
                print(f"ERRO AO ATUALIZAR TURMA (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar turma: {e}", "erro")
                print(f"ERRO AO ATUALIZAR TURMA: {e}")
            return redirect(url_for("turmas"))
        else:
            cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
            turma = cursor.fetchone()
            if not turma:
                flash("Turma não encontrada!", "erro")
                return redirect(url_for("turmas"))
            return render_template("editar_turma.html", turma=turma)
    except Exception as e:
        flash(f"Erro ao carregar/editar turma: {e}", "erro")
        print(f"ERRO EM EDITAR TURMA (GET/POST): {e}")
        return redirect(url_for("turmas"))
    finally:
        if conn:
            conn.close()


@app.route("/turmas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_turma(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        # Verificar se há alunos matriculados nesta turma
        cursor.execute("SELECT COUNT(*) FROM alunos WHERE turma_id = ?", (id,))
        total_alunos = cursor.fetchone()[0]
        if total_alunos > 0:
            flash(f"Não é possível excluir a turma. Existem {total_alunos} alunos associados a ela.", "erro")
            return redirect(url_for("turmas"))

        cursor.execute("DELETE FROM turmas WHERE id=?", (id,))
        conn.commit()
        flash("Turma excluída!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir turma: {e}", "erro")
        print(f"ERRO AO EXCLUIR TURMA: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("turmas"))


# ══════════════════════════════════════
# DISCIPLINAS
# ══════════════════════════════════════
@app.route("/disciplinas")
@login_required
def disciplinas():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT d.*, u.nome as professor_nome
            FROM disciplinas d
            LEFT JOIN usuarios u ON d.professor_id = u.id
            ORDER BY d.nome
        """)
        lista = cursor.fetchall()
        return render_template("disciplinas.html", disciplinas=lista)
    except Exception as e:
        flash(f"Erro ao carregar disciplinas: {e}", "erro")
        print(f"ERRO EM DISCIPLINAS: {e}")
        return render_template("disciplinas.html", disciplinas=[])
    finally:
        if conn:
            conn.close()


@app.route("/disciplinas/nova", methods=["GET", "POST"])
@login_required
def nova_disciplina():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            nome            = request.form.get("nome", "").strip()
            descricao       = request.form.get("descricao", "").strip()
            professor_id    = request.form.get("professor_id", type=int)
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = request.form.get("frequencia_minima", type=float)

            if not nome:
                flash("Nome é obrigatório!", "erro")
                professores = cursor.execute("SELECT id, nome FROM usuarios WHERE perfil='professor' OR perfil='admin' ORDER BY nome").fetchall()
                return render_template("nova_disciplina.html", professores=professores)

            try:
                cursor.execute(
                    "INSERT INTO disciplinas (nome, descricao, professor_id, tem_atividades, frequencia_minima) VALUES (?,?,?,?,?)",
                    (nome, descricao, professor_id, tem_atividades, frequencia_minima))
                conn.commit()
                flash(f"Disciplina '{nome}' criada!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "disciplinas.nome" in str(e):
                    flash("Já existe uma disciplina com este nome!", "erro")
                else:
                    flash(f"Erro ao criar disciplina: {e}", "erro")
                print(f"ERRO AO CRIAR DISCIPLINA (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao criar disciplina: {e}", "erro")
                print(f"ERRO AO CRIAR DISCIPLINA: {e}")
            return redirect(url_for("disciplinas"))
        else:
            professores = cursor.execute("SELECT id, nome FROM usuarios WHERE perfil='professor' OR perfil='admin' ORDER BY nome").fetchall()
            return render_template("nova_disciplina.html", professores=professores)
    except Exception as e:
        flash(f"Erro ao carregar página de nova disciplina: {e}", "erro")
        print(f"ERRO EM NOVA DISCIPLINA (GET): {e}")
        return redirect(url_for("disciplinas"))
    finally:
        if conn:
            conn.close()


@app.route("/disciplinas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_disciplina(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            nome            = request.form.get("nome", "").strip()
            descricao       = request.form.get("descricao", "").strip()
            professor_id    = request.form.get("professor_id", type=int)
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = request.form.get("frequencia_minima", type=float)
            ativa           = 1 if request.form.get("ativa") == "on" else 0

            if not nome:
                flash("Nome é obrigatório!", "erro")
                professores = cursor.execute("SELECT id, nome FROM usuarios WHERE perfil='professor' OR perfil='admin' ORDER BY nome").fetchall()
                disciplina = cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,)).fetchone()
                return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)

            try:
                cursor.execute(
                    "UPDATE disciplinas SET nome=?, descricao=?, professor_id=?, tem_atividades=?, frequencia_minima=?, ativa=? WHERE id=?",
                    (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa, id))
                conn.commit()
                flash(f"Disciplina '{nome}' atualizada!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "disciplinas.nome" in str(e):
                    flash("Já existe uma disciplina com este nome!", "erro")
                else:
                    flash(f"Erro ao atualizar disciplina: {e}", "erro")
                print(f"ERRO AO ATUALIZAR DISCIPLINA (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar disciplina: {e}", "erro")
                print(f"ERRO AO ATUALIZAR DISCIPLINA: {e}")
            return redirect(url_for("disciplinas"))
        else:
            disciplina = cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,)).fetchone()
            if not disciplina:
                flash("Disciplina não encontrada!", "erro")
                return redirect(url_for("disciplinas"))
            professores = cursor.execute("SELECT id, nome FROM usuarios WHERE perfil='professor' OR perfil='admin' ORDER BY nome").fetchall()
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except Exception as e:
        flash(f"Erro ao carregar/editar disciplina: {e}", "erro")
        print(f"ERRO EM EDITAR DISCIPLINA (GET/POST): {e}")
        return redirect(url_for("disciplinas"))
    finally:
        if conn:
            conn.close()


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        # Verificar se há matrículas nesta disciplina
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id = ?", (id,))
        total_matriculas = cursor.fetchone()[0]
        if total_matriculas > 0:
            flash(f"Não é possível excluir a disciplina. Existem {total_matriculas} matrículas associadas a ela.", "erro")
            return redirect(url_for("disciplinas"))

        cursor.execute("DELETE FROM disciplinas WHERE id=?", (id,))
        conn.commit()
        flash("Disciplina excluída!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir disciplina: {e}", "erro")
        print(f"ERRO AO EXCLUIR DISCIPLINA: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("disciplinas"))


# ══════════════════════════════════════
# ALUNOS
# ══════════════════════════════════════
@app.route("/alunos")
@login_required
def alunos():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT a.*, t.nome as turma_nome, t.faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            ORDER BY a.nome
        """)
        lista = cursor.fetchall()
        return render_template("alunos.html", alunos=lista)
    except Exception as e:
        flash(f"Erro ao carregar alunos: {e}", "erro")
        print(f"ERRO EM ALUNOS: {e}")
        return render_template("alunos.html", alunos=[])
    finally:
        if conn:
            conn.close()


@app.route("/alunos/novo", methods=["GET", "POST"])
@login_required
def novo_aluno():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            nome            = request.form.get("nome", "").strip()
            data_nascimento = request.form.get("data_nascimento", "").strip()
            telefone        = request.form.get("telefone", "").strip()
            email           = request.form.get("email", "").strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id", type=int)
            # Novos campos
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            if not nome:
                flash("Nome do aluno é obrigatório!", "erro")
                turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
                return render_template("novo_aluno.html", turmas=turmas,
                                       aluno={'nome': nome, 'data_nascimento': data_nascimento,
                                              'telefone': telefone, 'email': email,
                                              'membro_igreja': membro_igreja, 'turma_id': turma_id,
                                              'nome_pai': nome_pai, 'nome_mae': nome_mae, 'endereco': endereco})

            try:
                cursor.execute(
                    """INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco))
                conn.commit()
                flash(f"Aluno '{nome}' cadastrado!", "sucesso")
            except Exception as e:
                flash(f"Erro ao cadastrar aluno: {e}", "erro")
                print(f"ERRO AO CADASTRAR ALUNO: {e}")
            return redirect(url_for("alunos"))
        else:
            turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
            return render_template("novo_aluno.html", turmas=turmas, aluno={})
    except Exception as e:
        flash(f"Erro ao carregar página de novo aluno: {e}", "erro")
        print(f"ERRO EM NOVO ALUNO (GET): {e}")
        return redirect(url_for("alunos"))
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_aluno(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            nome            = request.form.get("nome", "").strip()
            data_nascimento = request.form.get("data_nascimento", "").strip()
            telefone        = request.form.get("telefone", "").strip()
            email           = request.form.get("email", "").strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id", type=int)
            # Novos campos
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            if not nome:
                flash("Nome do aluno é obrigatório!", "erro")
                turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
                aluno = cursor.execute("SELECT * FROM alunos WHERE id=?", (id,)).fetchone()
                return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)

            try:
                cursor.execute(
                    """UPDATE alunos SET nome=?, data_nascimento=?, telefone=?, email=?, membro_igreja=?, turma_id=?, nome_pai=?, nome_mae=?, endereco=?
                       WHERE id=?""",
                    (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco, id))
                conn.commit()
                flash(f"Aluno '{nome}' atualizado!", "sucesso")
            except Exception as e:
                flash(f"Erro ao atualizar aluno: {e}", "erro")
                print(f"ERRO AO ATUALIZAR ALUNO: {e}")
            return redirect(url_for("alunos"))
        else:
            aluno = cursor.execute("SELECT * FROM alunos WHERE id=?", (id,)).fetchone()
            if not aluno:
                flash("Aluno não encontrado!", "erro")
                return redirect(url_for("alunos"))
            turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
    except Exception as e:
        flash(f"Erro ao carregar/editar aluno: {e}", "erro")
        print(f"ERRO EM EDITAR ALUNO (GET/POST): {e}")
        return redirect(url_for("alunos"))
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_aluno(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        # Verificar se há matrículas para este aluno
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE aluno_id = ?", (id,))
        total_matriculas = cursor.fetchone()[0]
        if total_matriculas > 0:
            flash(f"Não é possível excluir o aluno. Existem {total_matriculas} matrículas associadas a ele.", "erro")
            return redirect(url_for("alunos"))

        cursor.execute("DELETE FROM alunos WHERE id=?", (id,))
        conn.commit()
        flash("Aluno excluído!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir aluno: {e}", "erro")
        print(f"ERRO AO EXCLUIR ALUNO: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("alunos"))


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha_aluno(id):
    conn = None
    aluno_dict = {}
    matriculas_processadas = []
    try:
        conn = conectar()
        cursor = conn.cursor()

        # Buscar dados do aluno
        aluno = cursor.execute("SELECT a.*, t.faixa_etaria as turma_faixa_etaria FROM alunos a LEFT JOIN turmas t ON a.turma_id = t.id WHERE a.id = ?", (id,)).fetchone()
        if not aluno:
            flash("Aluno não encontrado!", "erro")
            return redirect(url_for("alunos"))
        aluno_dict = dict(aluno) # Converter para dict mutável

        # Buscar matrículas do aluno
        raw_matriculas = cursor.execute("""
            SELECT m.id as matricula_id, m.data_inicio, m.data_conclusao, m.status,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                   d.nome as disciplina_nome, d.tem_atividades, d.frequencia_minima,
                   t.faixa_etaria as turma_faixa_etaria
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE m.aluno_id = ?
            ORDER BY d.nome
        """, (id,)).fetchall()

        for mat in raw_matriculas:
            matricula_dict = dict(mat) # Converter para dict mutável

            # --- Cálculo de Frequência ---
            historico_chamadas_raw = cursor.execute("""
                SELECT data_aula, presente, fez_atividade
                FROM presencas
                WHERE matricula_id = ?
                ORDER BY data_aula DESC
            """, (matricula_dict['matricula_id'],)).fetchall()

            historico_chamadas = [dict(c) for c in historico_chamadas_raw] # Converter para dict mutável
            matricula_dict['historico_chamadas'] = historico_chamadas

            presencas = sum(1 for c in historico_chamadas if c['presente'])
            total_aulas = len(historico_chamadas)
            atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            # --- Cálculo de Média e Status ---
            faixa_etaria_matricula = matricula_dict['turma_faixa_etaria'] or aluno_dict['turma_faixa_etaria'] # Prioriza da matrícula, senão do aluno

            nota_final_calc = None
            status = "cursando"
            media_display = "—"

            if faixa_etaria_matricula and "criancas" in faixa_etaria_matricula:
                # Crianças: Apenas frequência
                if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status = "aprovado"
                else:
                    status = "reprovado"
                media_display = "N/A" # Crianças não têm média de notas
            else: # Adolescentes/Jovens e Adultos
                # Lógica para Adolescentes/Jovens
                if faixa_etaria_matricula in ['adolescentes_13_15', 'jovens_16_17']:
                    meditacao_val = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                    versiculos_val = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                    desafio_nota_val = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                    visitante_val = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
                    nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
                    media_display = f"{nota_final_calc:.1f}"
                # Lógica para Adultos
                elif faixa_etaria_matricula == 'adultos':
                    nota1_val = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                    nota2_val = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
                    participacao_val = matricula_dict['participacao'] if matricula_dict['participacao'] is not None else 0
                    desafio_val = matricula_dict['desafio'] if matricula_dict['desafio'] is not None else 0
                    prova_val = matricula_dict['prova'] if matricula_dict['prova'] is not None else 0
                    nota_final_calc = nota1_val + nota2_val + participacao_val + desafio_val + prova_val
                    media_display = f"{nota_final_calc:.1f}"

                # Determinar status para Adolescentes/Jovens e Adultos
                if matricula_dict['status'] == 'cursando':
                    if matricula_dict['data_conclusao']: # Se já tem data de conclusão, mas status é 'cursando', recalcular
                        if nota_final_calc is not None and nota_final_calc >= 6.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                            status = "aprovado"
                        else:
                            status = "reprovado"
                else: # Se o status já foi definido (aprovado, reprovado, trancado)
                    status = matricula_dict['status']

            matricula_dict['nota_final_calc'] = nota_final_calc
            matricula_dict['media_display'] = media_display
            matricula_dict['status'] = status # Atualiza o status da matrícula no dicionário

            # Adicionar status_display para exibir no template
            if status == 'aprovado' and matricula_dict['data_conclusao'] is None:
                matricula_dict['status_display'] = "Aprovado (Provisório)"
            elif status == 'aprovado':
                matricula_dict['status_display'] = "Aprovado"
            elif status == 'reprovado':
                matricula_dict['status_display'] = "Reprovado"
            elif status == 'trancado':
                matricula_dict['status_display'] = "Trancado"
            else:
                matricula_dict['status_display'] = "Cursando"

            matriculas_processadas.append(matricula_dict)

        return render_template("trilha_aluno.html", aluno=aluno_dict, matriculas=matriculas_processadas)
    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        print(f"ERRO EM TRILHA DO ALUNO: {e}")
        return redirect(url_for("alunos"))
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# MATRÍCULAS
# ══════════════════════════════════════
@app.route("/matriculas")
@login_required
def matriculas():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT m.id, a.nome as aluno_nome, d.nome as disciplina_nome,
                   m.data_inicio, m.data_conclusao, m.status
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            ORDER BY a.nome, d.nome
        """)
        lista = cursor.fetchall()
        return render_template("matriculas.html", matriculas=lista)
    except Exception as e:
        flash(f"Erro ao carregar matrículas: {e}", "erro")
        print(f"ERRO EM MATRÍCULAS: {e}")
        return render_template("matriculas.html", matriculas=[])
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/nova", methods=["GET", "POST"])
@login_required
def nova_matricula():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            aluno_id      = request.form.get("aluno_id", type=int)
            disciplina_id = request.form.get("disciplina_id", type=int)
            data_inicio   = request.form.get("data_inicio", "").strip()

            if not aluno_id or not disciplina_id or not data_inicio:
                flash("Todos os campos são obrigatórios!", "erro")
                alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
                disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
                return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas)

            try:
                cursor.execute(
                    "INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio) VALUES (?,?,?)",
                    (aluno_id, disciplina_id, data_inicio))
                conn.commit()
                flash("Matrícula criada com sucesso!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                    flash("Este aluno já está matriculado nesta disciplina!", "erro")
                else:
                    flash(f"Erro ao criar matrícula: {e}", "erro")
                print(f"ERRO AO CRIAR MATRÍCULA (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao criar matrícula: {e}", "erro")
                print(f"ERRO AO CRIAR MATRÍCULA: {e}")
            return redirect(url_for("matriculas"))
        else:
            alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
            disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
            return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas)
    except Exception as e:
        flash(f"Erro ao carregar página de nova matrícula: {e}", "erro")
        print(f"ERRO EM NOVA MATRÍCULA (GET): {e}")
        return redirect(url_for("matriculas"))
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/novo_aluno_disciplina", methods=["GET", "POST"])
@login_required
def novo_aluno_disciplina():
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            # Dados do Aluno
            nome_aluno      = request.form.get("nome_aluno", "").strip()
            data_nascimento = request.form.get("data_nascimento", "").strip()
            telefone        = request.form.get("telefone", "").strip()
            email           = request.form.get("email", "").strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id", type=int)
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            # Dados da Matrícula
            disciplina_id   = request.form.get("disciplina_id", type=int)
            data_inicio     = request.form.get("data_inicio", "").strip()

            if not nome_aluno or not disciplina_id or not data_inicio:
                flash("Nome do aluno, disciplina e data de início são obrigatórios!", "erro")
                turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
                disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
                return render_template("novo_aluno_disciplina.html", turmas=turmas, disciplinas=disciplinas,
                                       aluno={'nome': nome_aluno, 'data_nascimento': data_nascimento,
                                              'telefone': telefone, 'email': email,
                                              'membro_igreja': membro_igreja, 'turma_id': turma_id,
                                              'nome_pai': nome_pai, 'nome_mae': nome_mae, 'endereco': endereco},
                                       matricula={'disciplina_id': disciplina_id, 'data_inicio': data_inicio})

            try:
                # 1. Inserir o novo aluno
                cursor.execute(
                    """INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (nome_aluno, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco))
                aluno_id = cursor.lastrowid # Pega o ID do aluno recém-criado

                # 2. Matricular o aluno na disciplina
                cursor.execute(
                    "INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio) VALUES (?,?,?)",
                    (aluno_id, disciplina_id, data_inicio))
                conn.commit()
                flash(f"Aluno '{nome_aluno}' cadastrado e matriculado na disciplina com sucesso!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                    flash("Este aluno já está matriculado nesta disciplina!", "erro")
                else:
                    flash(f"Erro de integridade ao cadastrar/matricular: {e}", "erro")
                print(f"ERRO AO CADASTRAR/MATRICULAR (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao cadastrar/matricular aluno: {e}", "erro")
                print(f"ERRO AO CADASTRAR/MATRICULAR: {e}")
            return redirect(url_for("matriculas"))
        else:
            turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
            disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
            return render_template("novo_aluno_disciplina.html", turmas=turmas, disciplinas=disciplinas, aluno={}, matricula={})
    except Exception as e:
        flash(f"Erro ao carregar página de novo aluno e matrícula: {e}", "erro")
        print(f"ERRO EM NOVO ALUNO DISCIPLINA (GET): {e}")
        return redirect(url_for("matriculas"))
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            aluno_id          = request.form.get("aluno_id", type=int)
            disciplina_id     = request.form.get("disciplina_id", type=int)
            data_inicio       = request.form.get("data_inicio", "").strip()
            data_conclusao    = request.form.get("data_conclusao", "").strip() or None
            status            = request.form.get("status", "").strip()

            # Campos de notas para Adultos
            nota1_adulto      = request.form.get("nota1_adulto", type=float)
            nota2_adulto      = request.form.get("nota2_adulto", type=float)
            participacao_adulto = request.form.get("participacao_adulto", type=float)
            desafio_adulto    = request.form.get("desafio_adulto", type=float)
            prova_adulto      = request.form.get("prova_adulto", type=float)

            # Campos de notas para Adolescentes/Jovens
            meditacao_aj      = request.form.get("meditacao_aj", type=float)
            versiculos_aj     = request.form.get("versiculos_aj", type=float)
            desafio_nota_aj   = request.form.get("desafio_nota_aj", type=float)
            visitante_aj      = request.form.get("visitante_aj", type=float)

            if not aluno_id or not disciplina_id or not data_inicio or not status:
                flash("Todos os campos obrigatórios devem ser preenchidos!", "erro")
                alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
                disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
                matricula = cursor.execute("SELECT * FROM matriculas WHERE id=?", (id,)).fetchone()
                return render_template("editar_matricula.html", matricula=matricula, alunos=alunos, disciplinas=disciplinas)

            try:
                # Primeiro, obter a faixa etária da turma do aluno para saber qual lógica de nota aplicar
                cursor.execute("""
                    SELECT t.faixa_etaria
                    FROM matriculas m
                    JOIN alunos a ON m.aluno_id = a.id
                    LEFT JOIN turmas t ON a.turma_id = t.id
                    WHERE m.id = ?
                """, (id,))
                faixa_etaria_matricula = cursor.fetchone()
                faixa_etaria_matricula = faixa_etaria_matricula['faixa_etaria'] if faixa_etaria_matricula else 'adultos' # Default para adultos

                # Resetar todas as notas para NULL antes de aplicar as novas
                update_query = """
                    UPDATE matriculas SET
                        aluno_id=?, disciplina_id=?, data_inicio=?, data_conclusao=?, status=?,
                        nota1=NULL, nota2=NULL, participacao=NULL, desafio=NULL, prova=NULL,
                        meditacao=NULL, versiculos=NULL, desafio_nota=NULL, visitante=NULL
                    WHERE id=?
                """
                update_params = [aluno_id, disciplina_id, data_inicio, data_conclusao, status, id]
                cursor.execute(update_query, update_params)

                # Aplicar notas com base na faixa etária
                if faixa_etaria_matricula == 'adultos':
                    cursor.execute("""
                        UPDATE matriculas SET
                            nota1=?, nota2=?, participacao=?, desafio=?, prova=?
                        WHERE id=?
                    """, (nota1_adulto, nota2_adulto, participacao_adulto, desafio_adulto, prova_adulto, id))
                elif faixa_etaria_matricula in ['adolescentes_13_15', 'jovens_16_17']:
                    cursor.execute("""
                        UPDATE matriculas SET
                            meditacao=?, versiculos=?, desafio_nota=?, visitante=?
                        WHERE id=?
                    """, (meditacao_aj, versiculos_aj, desafio_nota_aj, visitante_aj, id))
                # Crianças não têm campos de nota aqui

                conn.commit()
                flash("Matrícula atualizada com sucesso!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                    flash("Este aluno já está matriculado nesta disciplina!", "erro")
                else:
                    flash(f"Erro de integridade ao atualizar matrícula: {e}", "erro")
                print(f"ERRO AO ATUALIZAR MATRÍCULA (INTEGRITY): {e}")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar matrícula: {e}", "erro")
                print(f"ERRO AO ATUALIZAR MATRÍCULA: {e}")
            return redirect(url_for("matriculas"))
        else:
            matricula = cursor.execute("SELECT * FROM matriculas WHERE id=?", (id,)).fetchone()
            if not matricula:
                flash("Matrícula não encontrada!", "erro")
                return redirect(url_for("matriculas"))

            # Obter a faixa etária da turma do aluno para exibir os campos de nota corretos
            cursor.execute("""
                SELECT t.faixa_etaria
                FROM alunos a
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE a.id = ?
            """, (matricula['aluno_id'],))
            faixa_etaria_aluno = cursor.fetchone()
            faixa_etaria_aluno = faixa_etaria_aluno['faixa_etaria'] if faixa_etaria_aluno else 'adultos' # Default

            alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
            disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
            return render_template("editar_matricula.html",
                                   matricula=matricula,
                                   alunos=alunos,
                                   disciplinas=disciplinas,
                                   faixa_etaria_aluno=faixa_etaria_aluno)
    except Exception as e:
        flash(f"Erro ao carregar/editar matrícula: {e}", "erro")
        print(f"ERRO EM EDITAR MATRÍCULA (GET/POST): {e}")
        return redirect(url_for("matriculas"))
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_matricula(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        # Excluir presenças associadas à matrícula primeiro
        cursor.execute("DELETE FROM presencas WHERE matricula_id = ?", (id,))
        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula e presenças associadas excluídas!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
        print(f"ERRO AO EXCLUIR MATRÍCULA: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# PRESENÇA / CHAMADA
# ══════════════════════════════════════
@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()

        selected_disciplina_id = request.form.get("disciplina_id", type=int)
        selected_turma_id = request.form.get("turma_id", type=int)
        selected_data_aula = request.form.get("data_aula", str(date.today()))

        alunos_chamada = []
        if request.method == "POST" and selected_disciplina_id and selected_turma_id and selected_data_aula:
            # Buscar alunos matriculados na disciplina e pertencentes à turma selecionada
            cursor.execute("""
                SELECT m.id as matricula_id, a.id as aluno_id, a.nome as aluno_nome, d.tem_atividades
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                WHERE m.disciplina_id = ? AND a.turma_id = ?
                ORDER BY a.nome
            """, (selected_disciplina_id, selected_turma_id))
            raw_alunos_matriculados = cursor.fetchall()

            # Para cada aluno, verificar se já existe registro de presença para a data
            for aluno_mat in raw_alunos_matriculados:
                aluno_dict = dict(aluno_mat) # Converter para dict mutável
                cursor.execute("""
                    SELECT presente, fez_atividade
                    FROM presencas
                    WHERE matricula_id = ? AND data_aula = ?
                """, (aluno_dict['matricula_id'], selected_data_aula))
                presenca_existente = cursor.fetchone()

                if presenca_existente:
                    aluno_dict['presente'] = presenca_existente['presente']
                    aluno_dict['fez_atividade'] = presenca_existente['fez_atividade']
                else:
                    aluno_dict['presente'] = 0 # Default para falta
                    aluno_dict['fez_atividade'] = 0 # Default para não fez atividade
                alunos_chamada.append(aluno_dict)

            # Se a requisição for para salvar a chamada
            if 'salvar_chamada' in request.form:
                try:
                    for aluno_data in alunos_chamada:
                        matricula_id = aluno_data['matricula_id']
                        presente = 1 if request.form.get(f"presente_{matricula_id}") == "on" else 0
                        fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") == "on" else 0

                        # Verificar se já existe e atualizar, senão inserir
                        cursor.execute("""
                            INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                            VALUES (?, ?, ?, ?)
                            ON CONFLICT(matricula_id, data_aula) DO UPDATE SET
                                presente = EXCLUDED.presente,
                                fez_atividade = EXCLUDED.fez_atividade
                        """, (matricula_id, selected_data_aula, presente, fez_atividade))
                    conn.commit()
                    flash("Chamada salva com sucesso!", "sucesso")
                    # Recarregar os alunos da chamada para refletir as mudanças salvas
                    return redirect(url_for('chamada', disciplina_id=selected_disciplina_id, turma_id=selected_turma_id, data_aula=selected_data_aula))
                except Exception as e:
                    flash(f"Erro ao salvar chamada: {e}", "erro")
                    print(f"ERRO AO SALVAR CHAMADA: {e}")

        # Histórico de Chamadas Recentes (para a disciplina e turma selecionadas)
        historico_chamadas_recentes = []
        if selected_disciplina_id and selected_turma_id:
            cursor.execute("""
                SELECT p.data_aula, COUNT(CASE WHEN p.presente = 1 THEN 1 END) as total_presentes,
                       COUNT(m.id) as total_alunos_matriculados
                FROM presencas p
                JOIN matriculas m ON p.matricula_id = m.id
                JOIN alunos a ON m.aluno_id = a.id
                WHERE m.disciplina_id = ? AND a.turma_id = ?
                GROUP BY p.data_aula
                ORDER BY p.data_aula DESC
                LIMIT 10
            """, (selected_disciplina_id, selected_turma_id))
            historico_chamadas_recentes = [dict(row) for row in cursor.fetchall()]


        return render_template("chamada.html",
                               disciplinas=disciplinas,
                               turmas=turmas,
                               selected_disciplina_id=selected_disciplina_id,
                               selected_turma_id=selected_turma_id,
                               selected_data_aula=selected_data_aula,
                               alunos_chamada=alunos_chamada,
                               historico_chamadas_recentes=historico_chamadas_recentes)
    except Exception as e:
        flash(f"Erro ao carregar página de chamada: {e}", "erro")
        print(f"ERRO EM CHAMADA: {e}")
        return render_template("chamada.html",
                               disciplinas=[], turmas=[], alunos_chamada=[],
                               selected_disciplina_id=None, selected_turma_id=None,
                               selected_data_aula=str(date.today()),
                               historico_chamadas_recentes=[])
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# RELATÓRIOS DE MATRÍCULAS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id=None, turma_id=None, aluno_id=None, status=None):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        query = """
            SELECT m.id as matricula_id, a.nome as aluno_nome, a.data_nascimento, a.telefone, a.email,
                   a.membro_igreja, a.nome_pai, a.nome_mae, a.endereco,
                   d.nome AS disciplina_nome,
                   t.nome AS turma_nome,
                   t.faixa_etaria AS turma_faixa_etaria,
                   d.tem_atividades, d.frequencia_minima,
                   m.data_inicio, m.data_conclusao, m.status,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE 1=1
        """
        params = []

        if disciplina_id:
            query += " AND m.disciplina_id = ?"
            params.append(disciplina_id)
        if turma_id:
            query += " AND t.id = ?"
            params.append(turma_id)
        if aluno_id:
            query += " AND a.id = ?"
            params.append(aluno_id)
        if status:
            query += " AND m.status = ?"
            params.append(status)

        query += " ORDER BY a.nome, d.nome"
        cursor.execute(query, params)
        raw_matriculas = cursor.fetchall()
        processed_matriculas = []

        for mat in raw_matriculas:
            matricula_dict = dict(mat) # Converter para dict mutável

            # --- Cálculo de Frequência ---
            cursor.execute("""
                SELECT presente, fez_atividade
                FROM presencas
                WHERE matricula_id = ?
            """, (matricula_dict['matricula_id'],))
            historico_chamadas = cursor.fetchall()

            presencas = sum(1 for c in historico_chamadas if c['presente'])
            total_aulas = len(historico_chamadas)
            atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            # --- Cálculo de Média e Status ---
            faixa_etaria_matricula = matricula_dict['turma_faixa_etaria']

            nota_final_calc = None
            media_display = "—"

            if faixa_etaria_matricula and "criancas" in faixa_etaria_matricula:
                media_display = "N/A" # Crianças não têm média de notas
            else: # Adolescentes/Jovens e Adultos
                if faixa_etaria_matricula in ['adolescentes_13_15', 'jovens_16_17']:
                    meditacao_val = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                    versiculos_val = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                    desafio_nota_val = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                    visitante_val = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
                    nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
                    media_display = f"{nota_final_calc:.1f}"
                elif faixa_etaria_matricula == 'adultos':
                    nota1_val = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                    nota2_val = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
                    participacao_val = matricula_dict['participacao'] if matricula_dict['participacao'] is not None else 0
                    desafio_val = matricula_dict['desafio'] if matricula_dict['desafio'] is not None else 0
                    prova_val = matricula_dict['prova'] if matricula_dict['prova'] is not None else 0
                    nota_final_calc = nota1_val + nota2_val + participacao_val + desafio_val + prova_val
                    media_display = f"{nota_final_calc:.1f}"

            matricula_dict['nota_final_calc'] = nota_final_calc
            matricula_dict['media_display'] = media_display

            # Determinar status final para o relatório
            status_final = matricula_dict['status']
            if status_final == 'cursando':
                if matricula_dict['data_conclusao']: # Se já tem data de conclusão, mas status é 'cursando', recalcular
                    if faixa_etaria_matricula and "criancas" in faixa_etaria_matricula:
                        if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                            status_final = "aprovado"
                        else:
                            status_final = "reprovado"
                    elif nota_final_calc is not None and nota_final_calc >= 6.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                        status_final = "aprovado"
                    else:
                        status_final = "reprovado"
            matricula_dict['status_final'] = status_final

            processed_matriculas.append(matricula_dict)

        return processed_matriculas
    except Exception as e:
        print(f"ERRO EM GET_RELATORIO_DATA: {e}")
        return []
    finally:
        if conn:
            conn.close()


@app.route("/relatorios")
@login_required
def relatorios():
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
        alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
        return render_template("relatorios.html", disciplinas=disciplinas, turmas=turmas, alunos=alunos)
    except Exception as e:
        flash(f"Erro ao carregar página de relatórios: {e}", "erro")
        print(f"ERRO EM RELATORIOS (GET): {e}")
        return render_template("relatorios.html", disciplinas=[], turmas=[], alunos=[])
    finally:
        if conn:
            conn.close()


@app.route("/relatorios/gerar", methods=["POST"])
@login_required
def gerar_relatorio():
    disciplina_id = request.form.get("disciplina_id", type=int)
    turma_id = request.form.get("turma_id", type=int)
    aluno_id = request.form.get("aluno_id", type=int)
    status = request.form.get("status", "").strip()

    relatorio_data = get_relatorio_data(disciplina_id, turma_id, aluno_id, status)

    if not relatorio_data:
        flash("Nenhum dado encontrado para os filtros selecionados.", "warning")
        return redirect(url_for("relatorios"))

    return render_template("relatorios.html",
                           relatorio_data=relatorio_data,
                           disciplinas=get_relatorio_data_aux("disciplinas"),
                           turmas=get_relatorio_data_aux("turmas"),
                           alunos=get_relatorio_data_aux("alunos"),
                           selected_disciplina_id=disciplina_id,
                           selected_turma_id=turma_id,
                           selected_aluno_id=aluno_id,
                           selected_status=status)


def get_relatorio_data_aux(entity_type):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        if entity_type == "disciplinas":
            return cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        elif entity_type == "turmas":
            return cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
        elif entity_type == "alunos":
            return cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
        return []
    except Exception as e:
        print(f"ERRO EM GET_RELATORIO_DATA_AUX ({entity_type}): {e}")
        return []
    finally:
        if conn:
            conn.close()


@app.route("/relatorios/download/<format>", methods=["POST"])
@login_required
def download_relatorio(format):
    disciplina_id = request.form.get("disciplina_id", type=int)
    turma_id = request.form.get("turma_id", type=int)
    aluno_id = request.form.get("aluno_id", type=int)
    status = request.form.get("status", "").strip()

    relatorio_data = get_relatorio_data(disciplina_id, turma_id, aluno_id, status)

    if not relatorio_data:
        flash("Nenhum dado para download com os filtros selecionados.", "warning")
        return redirect(url_for("relatorios"))

    if format == "pdf":
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        story = []

        story.append(Paragraph("Relatório de Matrículas", styles['h1']))
        story.append(Spacer(1, 0.2 * inch))

        data = [['Aluno', 'Disciplina', 'Turma', 'Faixa Etária', 'Início', 'Conclusão', 'Status', 'Média', 'Freq. %']]
        for item in relatorio_data:
            data.append([
                item['aluno_nome'],
                item['disciplina_nome'],
                item['turma_nome'] or 'N/A',
                item['turma_faixa_etaria'].replace('_', ' ').title() if item['turma_faixa_etaria'] else 'N/A',
                item['data_inicio'],
                item['data_conclusao'] or 'N/A',
                item['status_final'].title(),
                item['media_display'],
                f"{item['frequencia_porcentagem']:.1f}%"
            ])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(table)
        doc.build(story)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.pdf", mimetype="application/pdf")

    elif format == "docx":
        document = Document()
        document.add_heading('Relatório de Matrículas', level=1)

        table = document.add_table(rows=1, cols=9)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        headers = ['Aluno', 'Disciplina', 'Turma', 'Faixa Etária', 'Início', 'Conclusão', 'Status', 'Média', 'Freq. %']
        for i, header in enumerate(headers):
            hdr_cells[i].text = header
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        for item in relatorio_data:
            row_cells = table.add_row().cells
            row_cells[0].text = item['aluno_nome']
            row_cells[1].text = item['disciplina_nome']
            row_cells[2].text = item['turma_nome'] or 'N/A'
            row_cells[3].text = item['turma_faixa_etaria'].replace('_', ' ').title() if item['turma_faixa_etaria'] else 'N/A'
            row_cells[4].text = item['data_inicio']
            row_cells[5].text = item['data_conclusao'] or 'N/A'
            row_cells[6].text = item['status_final'].title()
            row_cells[7].text = item['media_display']
            row_cells[8].text = f"{item['frequencia_porcentagem']:.1f}%"
            for cell in row_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        buffer = BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    flash("Formato de download inválido.", "erro")
    return redirect(url_for("relatorios"))


# ══════════════════════════════════════
# RELATÓRIOS DE FREQUÊNCIA
# ══════════════════════════════════════
@app.route("/relatorios/frequencia", methods=["GET", "POST"])
@login_required
def relatorios_frequencia():
    conn = None
    disciplinas = []
    turmas = []
    alunos = []
    frequencia_data = []
    selected_disciplina_id = None
    selected_turma_id = None
    selected_aluno_id = None
    data_inicio_filtro = None
    data_fim_filtro = None

    try:
        conn = conectar()
        cursor = conn.cursor()

        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
        alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()

        if request.method == "POST":
            selected_disciplina_id = request.form.get("disciplina_id", type=int)
            selected_turma_id = request.form.get("turma_id", type=int)
            selected_aluno_id = request.form.get("aluno_id", type=int)
            data_inicio_filtro = request.form.get("data_inicio", "").strip() or None
            data_fim_filtro = request.form.get("data_fim", "").strip() or None

            query = """
                SELECT a.nome as aluno_nome, d.nome as disciplina_nome, t.nome as turma_nome,
                       m.id as matricula_id, d.tem_atividades, d.frequencia_minima
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE 1=1
            """
            params = []

            if selected_disciplina_id:
                query += " AND m.disciplina_id = ?"
                params.append(selected_disciplina_id)
            if selected_turma_id:
                query += " AND t.id = ?"
                params.append(selected_turma_id)
            if selected_aluno_id:
                query += " AND a.id = ?"
                params.append(selected_aluno_id)

            query += " ORDER BY a.nome, d.nome"
            cursor.execute(query, params)
            raw_frequencia_data = cursor.fetchall()

            for row in raw_frequencia_data:
                item_dict = dict(row) # Converter para dict mutável

                # Obter presenças para a matrícula e período
                presenca_query = """
                    SELECT presente, fez_atividade, data_aula
                    FROM presencas
                    WHERE matricula_id = ?
                """
                presenca_params = [item_dict['matricula_id']]

                if data_inicio_filtro:
                    presenca_query += " AND data_aula >= ?"
                    presenca_params.append(data_inicio_filtro)
                if data_fim_filtro:
                    presenca_query += " AND data_aula <= ?"
                    presenca_params.append(data_fim_filtro)

                cursor.execute(presenca_query, presenca_params)
                historico_chamadas = cursor.fetchall()

                presencas = sum(1 for c in historico_chamadas if c['presente'])
                total_aulas = len(historico_chamadas)
                atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

                item_dict['presencas'] = presencas
                item_dict['total_aulas'] = total_aulas
                item_dict['atividades_feitas'] = atividades_feitas
                item_dict['historico_chamadas'] = [dict(c) for c in historico_chamadas] # Para detalhes no download

                frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
                item_dict['frequencia_porcentagem'] = frequencia_porcentagem

                frequencia_data.append(item_dict)

    except Exception as e:
        flash(f"Erro no relatório de frequência: {e}", "erro")
        print(f"ERRO NO RELATÓRIO DE FREQUÊNCIA: {e}")
        # Resetar dados para evitar erros no template
        frequencia_data = []
    finally:
        if conn:
            conn.close()

    return render_template("relatorio_frequencia.html",
                           disciplinas=disciplinas,
                           turmas=turmas,
                           alunos=alunos,
                           frequencia_data=frequencia_data,
                           selected_disciplina_id=selected_disciplina_id,
                           selected_turma_id=selected_turma_id,
                           selected_aluno_id=selected_aluno_id,
                           data_inicio_filtro=data_inicio_filtro,
                           data_fim_filtro=data_fim_filtro)


@app.route("/relatorios/frequencia/download/<format>", methods=["POST"])
@login_required
def download_relatorio_frequencia(format):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        selected_disciplina_id = request.form.get("disciplina_id", type=int)
        selected_turma_id = request.form.get("turma_id", type=int)
        selected_aluno_id = request.form.get("aluno_id", type=int)
        data_inicio_filtro = request.form.get("data_inicio", "").strip() or None
        data_fim_filtro = request.form.get("data_fim", "").strip() or None

        query = """
            SELECT a.nome as aluno_nome, d.nome as disciplina_nome, t.nome as turma_nome,
                   m.id as matricula_id, d.tem_atividades, d.frequencia_minima
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE 1=1
        """
        params = []

        if selected_disciplina_id:
            query += " AND m.disciplina_id = ?"
            params.append(selected_disciplina_id)
        if selected_turma_id:
            query += " AND t.id = ?"
            params.append(selected_turma_id)
        if selected_aluno_id:
            query += " AND a.id = ?"
            params.append(selected_aluno_id)

        query += " ORDER BY a.nome, d.nome"
        cursor.execute(query, params)
        raw_frequencia_data = cursor.fetchall()
        frequencia_data = []

        for row in raw_frequencia_data:
            item_dict = dict(row) # Converter para dict mutável

            presenca_query = """
                SELECT presente, fez_atividade, data_aula
                FROM presencas
                WHERE matricula_id = ?
            """
            presenca_params = [item_dict['matricula_id']]

            if data_inicio_filtro:
                presenca_query += " AND data_aula >= ?"
                presenca_params.append(data_inicio_filtro)
            if data_fim_filtro:
                presenca_query += " AND data_aula <= ?"
                presenca_params.append(data_fim_filtro)

            cursor.execute(presenca_query, presenca_params)
            historico_chamadas = cursor.fetchall()

            presencas = sum(1 for c in historico_chamadas if c['presente'])
            total_aulas = len(historico_chamadas)
            atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

            item_dict['presencas'] = presencas
            item_dict['total_aulas'] = total_aulas
            item_dict['atividades_feitas'] = atividades_feitas
            item_dict['historico_chamadas'] = [dict(c) for c in historico_chamadas]

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            item_dict['frequencia_porcentagem'] = frequencia_porcentagem

            frequencia_data.append(item_dict)

        if not frequencia_data:
            flash("Nenhum dado para download com os filtros selecionados.", "warning")
            return redirect(url_for("relatorios_frequencia"))

        if format == "pdf":
            buffer = BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
            styles = getSampleStyleSheet()
            story = []

            story.append(Paragraph("Relatório de Frequência", styles['h1']))
            story.append(Spacer(1, 0.2 * inch))

            data = [['Aluno', 'Disciplina', 'Turma', 'Presenças', 'Total Aulas', 'Freq. %', 'Ativ. Feitas']]
            for item in frequencia_data:
                data.append([
                    item['aluno_nome'],
                    item['disciplina_nome'],
                    item['turma_nome'] or 'N/A',
                    str(item['presencas']),
                    str(item['total_aulas']),
                    f"{item['frequencia_porcentagem']:.1f}%",
                    str(item['atividades_feitas']) if item['tem_atividades'] else 'N/A'
                ])

            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(table)
            doc.build(story)
            buffer.seek(0)
            return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.pdf", mimetype="application/pdf")

        elif format == "docx":
            document = Document()
            document.add_heading('Relatório de Frequência', level=1)

            table = document.add_table(rows=1, cols=7)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            headers = ['Aluno', 'Disciplina', 'Turma', 'Presenças', 'Total Aulas', 'Freq. %', 'Ativ. Feitas']
            for i, header in enumerate(headers):
                hdr_cells[i].text = header
                hdr_cells[i].paragraphs[0].runs[0].font.bold = True
                hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            for item in frequencia_data:
                row_cells = table.add_row().cells
                row_cells[0].text = item['aluno_nome']
                row_cells[1].text = item['disciplina_nome']
                row_cells[2].text = item['turma_nome'] or 'N/A'
                row_cells[3].text = str(item['presencas'])
                row_cells[4].text = str(item['total_aulas'])
                row_cells[5].text = f"{item['frequencia_porcentagem']:.1f}%"
                row_cells[6].text = str(item['atividades_feitas']) if item['tem_atividades'] else 'N/A'
                for cell in row_cells:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

            buffer = BytesIO()
            document.save(buffer)
            buffer.seek(0)
            return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        flash("Formato de download inválido.", "erro")
        return redirect(url_for("relatorios_frequencia"))
    except Exception as e:
        flash(f"Erro ao gerar download do relatório de frequência: {e}", "erro")
        print(f"ERRO NO DOWNLOAD DO RELATÓRIO DE FREQUÊNCIA: {e}")
        return redirect(url_for("relatorios_frequencia"))
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# BACKUP E RESTAURAÇÃO
# ══════════════════════════════════════
@app.route("/admin/backup", methods=["GET", "POST"])
@login_required
@admin_required
def backup_restauracao():
    if request.method == "POST":
        if 'backup_action' in request.form: # Requisição para baixar backup
            try:
                backup_filename = f"escola_backup_{date.today().strftime('%Y-%m-%d')}.db"
                return send_file(DATABASE, as_attachment=True, download_name=backup_filename, mimetype="application/x-sqlite3")
            except Exception as e:
                flash(f"Erro ao gerar backup: {e}", "erro")
                print(f"ERRO NO BACKUP: {e}")
        elif 'restore_file' in request.files: # Requisição para restaurar
            file = request.files['restore_file']
            if file and file.filename.endswith('.db'):
                try:
                    # Criar um backup temporário do banco de dados atual antes de sobrescrever
                    temp_backup_path = DATABASE + ".temp_bak"
                    shutil.copy(DATABASE, temp_backup_path)

                    file.save(DATABASE)
                    flash("Banco de dados restaurado com sucesso! Pode ser necessário reiniciar o servidor para que as mudanças sejam totalmente aplicadas.", "sucesso")
                    # Remover o backup temporário após sucesso
                    if os.path.exists(temp_backup_path):
                        os.remove(temp_backup_path)
                except Exception as e:
                    # Se houver erro, tentar restaurar do backup temporário
                    if os.path.exists(temp_backup_path):
                        shutil.copy(temp_backup_path, DATABASE)
                        os.remove(temp_backup_path) # Remover o temp_bak
                    flash(f"Erro ao restaurar banco de dados: {e}", "erro")
                    print(f"ERRO NA RESTAURAÇÃO: {e}")
            else:
                flash("Por favor, selecione um arquivo de backup válido (.db).", "erro")
        else:
            flash("Ação inválida.", "erro")
        return redirect(url_for("backup_restauracao"))
    return render_template("backup_restauracao.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)