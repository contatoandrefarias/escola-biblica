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
        cursor.execute("SELECT COUNT(*) as t FROM usuarios WHERE perfil='professor'")
        total_professores = cursor.fetchone()["t"] # Ajustado para contar professores
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
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM turmas ORDER BY nome")
        lista = cursor.fetchall()
        return render_template("turmas.html", turmas=lista)
    except Exception as e:
        flash(f"Erro ao carregar turmas: {e}", "erro")
        print(f"ERRO EM TURMAS: {e}")
        return render_template("turmas.html", turmas=[])
    finally:
        if conn:
            conn.close()


@app.route("/turmas/novo", methods=["GET", "POST"])
@login_required
@admin_required
def nova_turma():
    if request.method == "POST":
        nome        = request.form["nome"].strip()
        descricao   = request.form["descricao"].strip()
        faixa_etaria = request.form["faixa_etaria"]
        ativa       = 1 if request.form.get("ativa") == "on" else 0

        conn = None
        try:
            conn = conectar()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO turmas (nome, descricao, faixa_etaria, ativa) VALUES (?, ?, ?, ?)",
                (nome, descricao, faixa_etaria, ativa))
            conn.commit()
            flash("Turma adicionada com sucesso!", "sucesso")
            return redirect(url_for("turmas"))
        except sqlite3.IntegrityError:
            flash("Já existe uma turma com este nome.", "erro")
        except Exception as e:
            flash(f"Erro ao adicionar turma: {e}", "erro")
            print(f"ERRO EM NOVA_TURMA: {e}")
        finally:
            if conn:
                conn.close()
    return render_template("nova_turma.html")


@app.route("/turmas/<int:id>/editar", methods=["GET", "POST"])
@login_required
@admin_required
def editar_turma(id):
    conn = None
    turma = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
        turma = cursor.fetchone()

        if not turma:
            flash("Turma não encontrada.", "erro")
            return redirect(url_for("turmas"))

        if request.method == "POST":
            nome        = request.form["nome"].strip()
            descricao   = request.form["descricao"].strip()
            faixa_etaria = request.form["faixa_etaria"]
            ativa       = 1 if request.form.get("ativa") == "on" else 0

            cursor.execute(
                "UPDATE turmas SET nome=?, descricao=?, faixa_etaria=?, ativa=? WHERE id=?",
                (nome, descricao, faixa_etaria, ativa, id))
            conn.commit()
            flash("Turma atualizada com sucesso!", "sucesso")
            return redirect(url_for("turmas"))
        return render_template("editar_turma.html", turma=turma)
    except sqlite3.IntegrityError:
        flash("Já existe uma turma com este nome.", "erro")
    except Exception as e:
        flash(f"Erro ao editar turma: {e}", "erro")
        print(f"ERRO EM EDITAR_TURMA: {e}")
    finally:
        if conn:
            conn.close()
    return render_template("editar_turma.html", turma=turma)


@app.route("/turmas/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_turma(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se há alunos associados a esta turma
        cursor.execute("SELECT COUNT(*) FROM alunos WHERE turma_id=?", (id,))
        total_alunos = cursor.fetchone()[0]
        if total_alunos > 0:
            flash(f"Não é possível excluir a turma. Existem {total_alunos} alunos associados a ela.", "erro")
            return redirect(url_for("turmas"))

        cursor.execute("DELETE FROM turmas WHERE id=?", (id,))
        conn.commit()
        flash("Turma excluída com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir turma: {e}", "erro")
        print(f"ERRO EM EXCLUIR_TURMA: {e}")
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
        conn = conectar()
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


@app.route("/disciplinas/novo", methods=["GET", "POST"])
@login_required
@admin_required
def nova_disciplina():
    conn = None
    professores = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome FROM usuarios WHERE perfil='professor' ORDER BY nome")
        professores = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            nome            = request.form["nome"].strip()
            descricao       = request.form["descricao"].strip()
            professor_id    = request.form.get("professor_id")
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = float(request.form.get("frequencia_minima", 75.0))
            ativa           = 1 if request.form.get("ativa") == "on" else 0

            cursor.execute(
                "INSERT INTO disciplinas (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa) VALUES (?, ?, ?, ?, ?, ?)",
                (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa))
            conn.commit()
            flash("Disciplina adicionada com sucesso!", "sucesso")
            return redirect(url_for("disciplinas"))
        return render_template("nova_disciplina.html", professores=professores)
    except sqlite3.IntegrityError:
        flash("Já existe uma disciplina com este nome.", "erro")
        return render_template("nova_disciplina.html", professores=professores)
    except Exception as e:
        flash(f"Erro ao adicionar disciplina: {e}", "erro")
        print(f"ERRO EM NOVA_DISCIPLINA: {e}")
        return render_template("nova_disciplina.html", professores=professores)
    finally:
        if conn:
            conn.close()


@app.route("/disciplinas/<int:id>/editar", methods=["GET", "POST"])
@login_required
@admin_required
def editar_disciplina(id):
    conn = None
    disciplina = None
    professores = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
        disciplina = cursor.fetchone()

        if not disciplina:
            flash("Disciplina não encontrada.", "erro")
            return redirect(url_for("disciplinas"))

        cursor.execute("SELECT id, nome FROM usuarios WHERE perfil='professor' ORDER BY nome")
        professores = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            nome            = request.form["nome"].strip()
            descricao       = request.form["descricao"].strip()
            professor_id    = request.form.get("professor_id")
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = float(request.form.get("frequencia_minima", 75.0))
            ativa           = 1 if request.form.get("ativa") == "on" else 0

            cursor.execute(
                "UPDATE disciplinas SET nome=?, descricao=?, professor_id=?, tem_atividades=?, frequencia_minima=?, ativa=? WHERE id=?",
                (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa, id))
            conn.commit()
            flash("Disciplina atualizada com sucesso!", "sucesso")
            return redirect(url_for("disciplinas"))
        return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except sqlite3.IntegrityError:
        flash("Já existe uma disciplina com este nome.", "erro")
        return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except Exception as e:
        flash(f"Erro ao editar disciplina: {e}", "erro")
        print(f"ERRO EM EDITAR_DISCIPLINA: {e}")
        return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    finally:
        if conn:
            conn.close()


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_disciplina(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se há matrículas associadas a esta disciplina
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id=?", (id,))
        total_matriculas = cursor.fetchone()[0]
        if total_matriculas > 0:
            flash(f"Não é possível excluir a disciplina. Existem {total_matriculas} matrículas associadas a ela.", "erro")
            return redirect(url_for("disciplinas"))

        cursor.execute("DELETE FROM disciplinas WHERE id=?", (id,))
        conn.commit()
        flash("Disciplina excluída com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir disciplina: {e}", "erro")
        print(f"ERRO EM EXCLUIR_DISCIPLINA: {e}")
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
    lista = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT a.*, t.nome as turma_nome
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            ORDER BY a.nome
        """)
        lista = [dict(row) for row in cursor.fetchall()]
        print(f"DEBUG ALUNOS: {lista}") # Log para depuração
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
@admin_required
def novo_aluno():
    conn = None
    turmas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            nome            = request.form["nome"].strip()
            data_nascimento = request.form["data_nascimento"]
            telefone        = request.form["telefone"].strip()
            email           = request.form["email"].strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id") if request.form.get("turma_id") != '' else None
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            cursor.execute(
                "INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco))
            conn.commit()
            flash("Aluno adicionado com sucesso!", "sucesso")
            return redirect(url_for("alunos"))
        return render_template("novo_aluno.html", turmas=turmas)
    except Exception as e:
        flash(f"Erro inesperado ao cadastrar aluno: {e}", "erro")
        print(f"ERRO EM NOVO_ALUNO: {e}")
        return render_template("novo_aluno.html", turmas=turmas)
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
@admin_required
def editar_aluno(id):
    conn = None
    aluno = None
    turmas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
        aluno = cursor.fetchone()

        if not aluno:
            flash("Aluno não encontrado.", "erro")
            return redirect(url_for("alunos"))

        cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            nome            = request.form["nome"].strip()
            data_nascimento = request.form["data_nascimento"]
            telefone        = request.form["telefone"].strip()
            email           = request.form["email"].strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id") if request.form.get("turma_id") != '' else None
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            cursor.execute(
                "UPDATE alunos SET nome=?, data_nascimento=?, telefone=?, email=?, membro_igreja=?, turma_id=?, nome_pai=?, nome_mae=?, endereco=? WHERE id=?",
                (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco, id))
            conn.commit()
            flash("Aluno atualizado com sucesso!", "sucesso")
            return redirect(url_for("alunos"))
        return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
    except Exception as e:
        flash(f"Erro ao editar aluno: {e}", "erro")
        print(f"ERRO EM EDITAR_ALUNO: {e}")
        return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_aluno(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se há matrículas associadas a este aluno
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE aluno_id=?", (id,))
        total_matriculas = cursor.fetchone()[0]
        if total_matriculas > 0:
            flash(f"Não é possível excluir o aluno. Existem {total_matriculas} matrículas associadas a ele.", "erro")
            return redirect(url_for("alunos"))

        cursor.execute("DELETE FROM alunos WHERE id=?", (id,))
        conn.commit()
        flash("Aluno excluído com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir aluno: {e}", "erro")
        print(f"ERRO EM EXCLUIR_ALUNO: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("alunos"))


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha_aluno(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT a.*, t.nome as turma_nome
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno = cursor.fetchone()

        if not aluno:
            flash("Aluno não encontrado.", "erro")
            return redirect(url_for("alunos"))

        aluno_dict = dict(aluno)

        cursor.execute("""
            SELECT m.id as matricula_id, d.nome as disciplina_nome,
                   m.data_inicio, m.data_conclusao, m.status,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                   d.tem_atividades, d.frequencia_minima
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.aluno_id = ?
            ORDER BY d.nome
        """, (id,))
        matriculas = cursor.fetchall()

        matriculas_processadas = []
        for matricula in matriculas:
            matricula_dict = dict(matricula)
            status = matricula_dict['status']
            if status == 'aprovado':
                matricula_dict['status_display'] = 'Aprovado'
            elif status == 'reprovado':
                matricula_dict['status_display'] = 'Reprovado'
            elif status == 'cursando':
                matricula_dict['status_display'] = 'Cursando'
            elif status == 'trancado':
                matricula_dict['status_display'] = 'Trancado'
            else:
                matricula_dict['status_display'] = 'Desconhecido'

            matriculas_processadas.append(matricula_dict)

        return render_template("trilha_aluno.html", aluno=aluno_dict, matriculas=matriculas_processadas)
    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        print(f"ERRO NA TRILHA DO ALUNO: {e}")
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
    lista = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT m.id, a.nome as aluno_nome, d.nome as disciplina_nome,
                   m.data_inicio, m.data_conclusao, m.status
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            ORDER BY a.nome, d.nome
        """)
        lista = [dict(row) for row in cursor.fetchall()]
        return render_template("matriculas.html", matriculas=lista)
    except Exception as e:
        flash(f"Erro ao carregar matrículas: {e}", "erro")
        print(f"ERRO EM MATRÍCULAS: {e}")
        return render_template("matriculas.html", matriculas=[])
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/novo", methods=["GET", "POST"])
@login_required
def nova_matricula():
    conn = None
    alunos = []
    disciplinas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
        alunos = [dict(row) for row in cursor.fetchall()]
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            aluno_id        = request.form["aluno_id"]
            disciplina_id   = request.form["disciplina_id"]
            data_inicio     = request.form["data_inicio"]
            data_conclusao  = request.form.get("data_conclusao")
            status          = request.form.get("status", "cursando")

            cursor.execute(
                "INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, data_conclusao, status) VALUES (?, ?, ?, ?, ?)",
                (aluno_id, disciplina_id, data_inicio, data_conclusao, status))
            conn.commit()
            flash("Matrícula realizada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas)
    except sqlite3.IntegrityError:
        flash("Este aluno já está matriculado nesta disciplina.", "erro")
        return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas)
    except Exception as e:
        flash(f"Erro ao realizar matrícula: {e}", "erro")
        print(f"ERRO AO REALIZAR MATRÍCULA: {e}")
        return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas)
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn = None
    matricula = None
    alunos = []
    disciplinas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM matriculas WHERE id=?", (id,))
        matricula = cursor.fetchone()

        if not matricula:
            flash("Matrícula não encontrada.", "erro")
            return redirect(url_for("matriculas"))

        cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
        alunos = [dict(row) for row in cursor.fetchall()]
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            aluno_id        = request.form["aluno_id"]
            disciplina_id   = request.form["disciplina_id"]
            data_inicio     = request.form["data_inicio"]
            data_conclusao  = request.form.get("data_conclusao")
            status          = request.form.get("status", "cursando")
            nota1           = request.form.get("nota1")
            nota2           = request.form.get("nota2")
            participacao    = request.form.get("participacao")
            desafio         = request.form.get("desafio")
            prova           = request.form.get("prova")
            meditacao       = request.form.get("meditacao")
            versiculos      = request.form.get("versiculos")
            desafio_nota    = request.form.get("desafio_nota")
            visitante       = request.form.get("visitante")

            cursor.execute(
                """UPDATE matriculas SET aluno_id=?, disciplina_id=?, data_inicio=?, data_conclusao=?, status=?,
                   nota1=?, nota2=?, participacao=?, desafio=?, prova=?, meditacao=?, versiculos=?, desafio_nota=?, visitante=?
                   WHERE id=?""",
                (aluno_id, disciplina_id, data_inicio, data_conclusao, status,
                 nota1, nota2, participacao, desafio, prova, meditacao, versiculos, desafio_nota, visitante, id))
            conn.commit()
            flash("Matrícula atualizada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        return render_template("editar_matricula.html", matricula=matricula, alunos=alunos, disciplinas=disciplinas)
    except Exception as e:
        flash(f"Erro ao editar matrícula: {e}", "erro")
        print(f"ERRO EM EDITAR_MATRICULA: {e}")
        return render_template("editar_matricula.html", matricula=matricula, alunos=alunos, disciplinas=disciplinas)
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_matricula(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se há presenças associadas a esta matrícula
        cursor.execute("SELECT COUNT(*) FROM presencas WHERE matricula_id=?", (id,))
        total_presencas = cursor.fetchone()[0]
        if total_presencas > 0:
            flash(f"Não é possível excluir a matrícula. Existem {total_presencas} registros de presença associados a ela.", "erro")
            return redirect(url_for("matriculas"))

        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula excluída com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
        print(f"ERRO EM EXCLUIR_MATRICULA: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# PRESENÇAS / CHAMADA
# ══════════════════════════════════════
@app.route("/presencas")
@login_required
def presencas():
    conn = None
    chamadas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT p.id, a.nome as aluno_nome, d.nome as disciplina_nome,
                   p.data_aula, p.presente, p.fez_atividade
            FROM presencas p
            JOIN matriculas m ON p.matricula_id = m.id
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            ORDER BY p.data_aula DESC, d.nome, a.nome
        """)
        chamadas = [dict(row) for row in cursor.fetchall()]
        return render_template("presencas.html", chamadas=chamadas)
    except Exception as e:
        flash(f"Erro ao carregar presenças: {e}", "erro")
        print(f"ERRO EM PRESENCAS: {e}")
        return render_template("presencas.html", chamadas=[])
    finally:
        if conn:
            conn.close()


@app.route("/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = None
    disciplinas = []
    turmas = []
    alunos_matriculados = []
    data_aula = date.today().isoformat()
    disciplina_selecionada = request.args.get('disciplina_id', type=int)
    turma_selecionada = request.args.get('turma_id', type=int)

    try:
        conn = conectar()
        cursor = conn.cursor()

        # Carregar disciplinas (apenas as ativas)
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas = [dict(row) for row in cursor.fetchall()]

        # Carregar turmas (apenas as ativas)
        cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = [dict(row) for row in cursor.fetchall()]

        if disciplina_selecionada and turma_selecionada:
            # Buscar alunos matriculados na disciplina e turma selecionadas
            cursor.execute("""
                SELECT a.id as aluno_id, a.nome as aluno_nome, m.id as matricula_id
                FROM alunos a
                JOIN matriculas m ON a.id = m.aluno_id
                WHERE m.disciplina_id = ? AND a.turma_id = ?
                ORDER BY a.nome
            """, (disciplina_selecionada, turma_selecionada))
            alunos_matriculados = [dict(row) for row in cursor.fetchall()]

            # Verificar presenças existentes para a data e disciplina/turma
            for aluno in alunos_matriculados:
                cursor.execute("""
                    SELECT presente, fez_atividade FROM presencas
                    WHERE matricula_id = ? AND data_aula = ?
                """, (aluno['matricula_id'], data_aula))
                presenca_existente = cursor.fetchone()
                if presenca_existente:
                    aluno['presente'] = presenca_existente['presente']
                    aluno['fez_atividade'] = presenca_existente['fez_atividade']
                else:
                    aluno['presente'] = 0
                    aluno['fez_atividade'] = 0

        if request.method == "POST":
            data_aula = request.form["data_aula"]
            disciplina_id = request.form["disciplina_id"]
            turma_id = request.form["turma_id"]

            # Obter novamente os alunos para a disciplina e turma para processar o POST
            cursor.execute("""
                SELECT a.id as aluno_id, a.nome as aluno_nome, m.id as matricula_id
                FROM alunos a
                JOIN matriculas m ON a.id = m.aluno_id
                WHERE m.disciplina_id = ? AND a.turma_id = ?
                ORDER BY a.nome
            """, (disciplina_id, turma_id))
            alunos_para_salvar = [dict(row) for row in cursor.fetchall()]

            for aluno in alunos_para_salvar:
                matricula_id = aluno['matricula_id']
                presente = 1 if request.form.get(f"presente_{matricula_id}") == "on" else 0
                fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") == "on" else 0

                cursor.execute("""
                    INSERT OR REPLACE INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                    VALUES (?, ?, ?, ?)
                """, (matricula_id, data_aula, presente, fez_atividade))
            conn.commit()
            flash("Chamada salva com sucesso!", "sucesso")
            return redirect(url_for("chamada", disciplina_id=disciplina_id, turma_id=turma_id))

        return render_template("chamada.html",
                               disciplinas=disciplinas,
                               turmas=turmas,
                               alunos_matriculados=alunos_matriculados,
                               data_aula=data_aula,
                               disciplina_selecionada=disciplina_selecionada,
                               turma_selecionada=turma_selecionada)
    except Exception as e:
        flash(f"Erro ao carregar ou salvar chamada: {e}", "erro")
        print(f"ERRO EM CHAMADA: {e}")
        return render_template("chamada.html",
                               disciplinas=disciplinas,
                               turmas=turmas,
                               alunos_matriculados=[],
                               data_aula=data_aula,
                               disciplina_selecionada=disciplina_selecionada,
                               turma_selecionada=turma_selecionada)
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# RELATÓRIOS
# ══════════════════════════════════════
@app.route("/relatorios")
@login_required
def relatorios():
    return render_template("relatorios.html")


@app.route("/relatorios/frequencia", methods=["GET", "POST"])
@login_required
def relatorios_frequencia():
    conn = None
    disciplinas_lista = []
    turmas_lista = []
    alunos_lista = []
    frequencia_data = []
    selected_disciplina = request.form.get('disciplina_id', type=int) or request.args.get('disciplina_id', type=int)
    selected_turma = request.form.get('turma_id', type=int) or request.args.get('turma_id', type=int)
    selected_aluno = request.form.get('aluno_id', type=int) or request.args.get('aluno_id', type=int)
    data_inicio_filtro = request.form.get('data_inicio') or request.args.get('data_inicio')
    data_fim_filtro = request.form.get('data_fim') or request.args.get('data_fim')

    try:
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas_lista = [dict(row) for row in cursor.fetchall()]

        cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas_lista = [dict(row) for row in cursor.fetchall()]

        cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
        alunos_lista = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST" or (selected_disciplina or selected_turma or selected_aluno or data_inicio_filtro or data_fim_filtro):
            query = """
                SELECT m.id as matricula_id, a.nome as aluno_nome, d.nome as disciplina_nome,
                       t.nome as turma_nome, d.tem_atividades, d.frequencia_minima
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE 1=1
            """
            params = []

            if selected_disciplina:
                query += " AND d.id = ?"
                params.append(selected_disciplina)
            if selected_turma:
                query += " AND t.id = ?"
                params.append(selected_turma)
            if selected_aluno:
                query += " AND a.id = ?"
                params.append(selected_aluno)

            cursor.execute(query, params)
            matriculas_filtradas = [dict(row) for row in cursor.fetchall()]

            frequencia_data = []
            for matricula in matriculas_filtradas:
                matricula_id = matricula['matricula_id']
                presenca_query = "SELECT COUNT(*) FROM presencas WHERE matricula_id=?"
                presenca_params = [matricula_id]

                if data_inicio_filtro:
                    presenca_query += " AND data_aula >= ?"
                    presenca_params.append(data_inicio_filtro)
                if data_fim_filtro:
                    presenca_query += " AND data_aula <= ?"
                    presenca_params.append(data_fim_filtro)

                cursor.execute(presenca_query, presenca_params)
                total_aulas = cursor.fetchone()[0]

                presenca_query_presente = presenca_query + " AND presente=1"
                cursor.execute(presenca_query_presente, presenca_params)
                presencas = cursor.fetchone()[0]

                frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

                atividades_feitas = 'N/A'
                if matricula['tem_atividades']:
                    presenca_query_atividade = presenca_query + " AND fez_atividade=1"
                    cursor.execute(presenca_query_atividade, presenca_params)
                    atividades_feitas = cursor.fetchone()[0]

                frequencia_data.append({
                    'aluno_nome': matricula['aluno_nome'],
                    'disciplina_nome': matricula['disciplina_nome'],
                    'turma_nome': matricula['turma_nome'] or 'N/A',
                    'presencas': presencas,
                    'total_aulas': total_aulas,
                    'frequencia_porcentagem': frequencia_porcentagem,
                    'tem_atividades': matricula['tem_atividades'],
                    'atividades_feitas': atividades_feitas,
                    'frequencia_minima': matricula['frequencia_minima']
                })

        return render_template("relatorios_frequencia.html",
                               disciplinas=disciplinas_lista,
                               turmas=turmas_lista,
                               alunos=alunos_lista,
                               frequencia_data=frequencia_data,
                               selected_disciplina=selected_disciplina,
                               selected_turma=selected_turma,
                               selected_aluno=selected_aluno,
                               data_inicio_filtro=data_inicio_filtro,
                               data_fim_filtro=data_fim_filtro)
    except Exception as e:
        flash(f"Erro no relatório de frequência: {e}", "erro")
        print(f"ERRO NO RELATÓRIO DE FREQUÊNCIA: {e}")
        return render_template("relatorios_frequencia.html",
                               disciplinas=[], turmas=[], alunos=[], frequencia_data=[],
                               selected_disciplina=selected_disciplina,
                               selected_turma=selected_turma,
                               selected_aluno=selected_aluno,
                               data_inicio_filtro=data_inicio_filtro,
                               data_fim_filtro=data_fim_filtro)
    finally:
        if conn:
            conn.close()


@app.route("/download_relatorio_frequencia", methods=["POST"])
@login_required
def download_relatorio_frequencia():
    conn = None
    try:
        format = request.form.get("format")
        selected_disciplina = request.form.get('disciplina_id', type=int)
        selected_turma = request.form.get('turma_id', type=int)
        selected_aluno = request.form.get('aluno_id', type=int)
        data_inicio_filtro = request.form.get('data_inicio')
        data_fim_filtro = request.form.get('data_fim')

        conn = conectar()
        cursor = conn.cursor()

        query = """
            SELECT m.id as matricula_id, a.nome as aluno_nome, d.nome as disciplina_nome,
                   t.nome as turma_nome, d.tem_atividades, d.frequencia_minima
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE 1=1
        """
        params = []

        if selected_disciplina:
            query += " AND d.id = ?"
            params.append(selected_disciplina)
        if selected_turma:
            query += " AND t.id = ?"
            params.append(selected_turma)
        if selected_aluno:
            query += " AND a.id = ?"
            params.append(selected_aluno)

        cursor.execute(query, params)
        matriculas_filtradas = [dict(row) for row in cursor.fetchall()]

        frequencia_data = []
        for matricula in matriculas_filtradas:
            matricula_id = matricula['matricula_id']
            presenca_query = "SELECT COUNT(*) FROM presencas WHERE matricula_id=?"
            presenca_params = [matricula_id]

            if data_inicio_filtro:
                presenca_query += " AND data_aula >= ?"
                presenca_params.append(data_inicio_filtro)
            if data_fim_filtro:
                presenca_query += " AND data_aula <= ?"
                presenca_params.append(data_fim_filtro)

            cursor.execute(presenca_query, presenca_params)
            total_aulas = cursor.fetchone()[0]

            presenca_query_presente = presenca_query + " AND presente=1"
            cursor.execute(presenca_query_presente, presenca_params)
            presencas = cursor.fetchone()[0]

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

            atividades_feitas = 'N/A'
            if matricula['tem_atividades']:
                presenca_query_atividade = presenca_query + " AND fez_atividade=1"
                cursor.execute(presenca_query_atividade, presenca_params)
                atividades_feitas = cursor.fetchone()[0]

            frequencia_data.append({
                'aluno_nome': matricula['aluno_nome'],
                'disciplina_nome': matricula['disciplina_nome'],
                'turma_nome': matricula['turma_nome'] or 'N/A',
                'presencas': presencas,
                'total_aulas': total_aulas,
                'frequencia_porcentagem': frequencia_porcentagem,
                'tem_atividades': matricula['tem_atividades'],
                'atividades_feitas': atividades_feitas,
                'frequencia_minima': matricula['frequencia_minima']
            })

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
# USUÁRIOS (ADMIN)
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
@admin_required
def usuarios():
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, email, perfil FROM usuarios ORDER BY nome")
        lista = [dict(row) for row in cursor.fetchall()]
        return render_template("usuarios.html", usuarios=lista)
    except Exception as e:
        flash(f"Erro ao carregar usuários: {e}", "erro")
        print(f"ERRO EM USUARIOS: {e}")
        return render_template("usuarios.html", usuarios=[])
    finally:
        if conn:
            conn.close()


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
@admin_required
def novo_usuario():
    if request.method == "POST":
        nome    = request.form["nome"].strip()
        email   = request.form["email"].strip()
        senha   = request.form["senha"]
        perfil  = request.form["perfil"]

        if not nome or not email or not senha or not perfil:
            flash("Todos os campos são obrigatórios.", "erro")
            return render_template("novo_usuario.html")

        senha_hash = generate_password_hash(senha)

        conn = None
        try:
            conn = conectar()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO usuarios (nome, email, senha_hash, perfil) VALUES (?, ?, ?, ?)",
                (nome, email, senha_hash, perfil))
            conn.commit()
            flash("Usuário adicionado com sucesso!", "sucesso")
            return redirect(url_for("usuarios"))
        except sqlite3.IntegrityError:
            flash("Já existe um usuário com este e-mail.", "erro")
        except Exception as e:
            flash(f"Erro ao adicionar usuário: {e}", "erro")
            print(f"ERRO EM NOVO_USUARIO: {e}")
        finally:
            if conn:
                conn.close()
    return render_template("novo_usuario.html")


@app.route("/usuarios/<int:id>/editar", methods=["GET", "POST"])
@login_required
@admin_required
def editar_usuario(id):
    conn = None
    usuario = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, email, perfil FROM usuarios WHERE id=?", (id,))
        usuario = cursor.fetchone()

        if not usuario:
            flash("Usuário não encontrado.", "erro")
            return redirect(url_for("usuarios"))

        if request.method == "POST":
            nome    = request.form["nome"].strip()
            email   = request.form["email"].strip()
            perfil  = request.form["perfil"]
            senha   = request.form["senha"] # Pode ser vazio se não for alterada

            if not nome or not email or not perfil:
                flash("Nome, e-mail e perfil são obrigatórios.", "erro")
                return render_template("editar_usuario.html", usuario=usuario)

            if senha: # Se a senha foi fornecida, atualiza o hash
                senha_hash = generate_password_hash(senha)
                cursor.execute(
                    "UPDATE usuarios SET nome=?, email=?, perfil=?, senha_hash=? WHERE id=?",
                    (nome, email, perfil, senha_hash, id))
            else: # Caso contrário, atualiza sem alterar a senha
                cursor.execute(
                    "UPDATE usuarios SET nome=?, email=?, perfil=? WHERE id=?",
                    (nome, email, perfil, id))
            conn.commit()
            flash("Usuário atualizado com sucesso!", "sucesso")
            return redirect(url_for("usuarios"))
        return render_template("editar_usuario.html", usuario=usuario)
    except sqlite3.IntegrityError:
        flash("Já existe um usuário com este e-mail.", "erro")
        return render_template("editar_usuario.html", usuario=usuario)
    except Exception as e:
        flash(f"Erro ao editar usuário: {e}", "erro")
        print(f"ERRO EM EDITAR_USUARIO: {e}")
        return render_template("editar_usuario.html", usuario=usuario)
    finally:
        if conn:
            conn.close()


@app.route("/usuarios/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_usuario(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se o usuário é professor e está associado a alguma disciplina
        cursor.execute("SELECT perfil FROM usuarios WHERE id=?", (id,))
        usuario_perfil = cursor.fetchone()
        if usuario_perfil and usuario_perfil['perfil'] == 'professor':
            cursor.execute("SELECT COUNT(*) FROM disciplinas WHERE professor_id=?", (id,))
            total_disciplinas = cursor.fetchone()[0]
            if total_disciplinas > 0:
                flash(f"Não é possível excluir este professor. Ele está associado a {total_disciplinas} disciplina(s).", "erro")
                return redirect(url_for("usuarios"))

        cursor.execute("DELETE FROM usuarios WHERE id=?", (id,))
        conn.commit()
        flash("Usuário excluído com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir usuário: {e}", "erro")
        print(f"ERRO EM EXCLUIR_USUARIO: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("usuarios"))


# ══════════════════════════════════════
# PROFESSORES (LISTAGEM ESPECÍFICA)
# ══════════════════════════════════════
@app.route("/professores")
@login_required
def professores(): # O nome da função é 'professores'
    conn = None
    lista = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, email FROM usuarios WHERE perfil='professor' ORDER BY nome")
        lista = [dict(row) for row in cursor.fetchall()]
        print(f"DEBUG PROFESSORES: {lista}") # Log para depuração
        return render_template("professores.html", professores=lista)
    except Exception as e:
        flash(f"Erro ao carregar professores: {e}", "erro")
        print(f"ERRO EM PROFESSORES: {e}")
        return render_template("professores.html", professores=[])
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


# Adiciona um log para listar todas as rotas registradas pelo Flask
with app.test_request_context():
    print("\n══════════════════════════════════════")
    print("ROTAS REGISTRADAS NO FLASK:")
    print("══════════════════════════════════════")
    for rule in app.url_map.iter_rules():
        print(f"Endpoint: {rule.endpoint}, Métodos: {rule.methods}, Rota: {rule.rule}")
    print("══════════════════════════════════════\n")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)