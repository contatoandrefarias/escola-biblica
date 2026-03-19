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
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT t.*, COUNT(a.id) as total_alunos
            FROM turmas t
            LEFT JOIN alunos a ON a.turma_id = t.id
            GROUP BY t.id
            ORDER BY t.nome
        """)
        lista = [dict(row) for row in cursor.fetchall()]
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
            print(f"ERRO AO ADICIONAR TURMA: {e}")
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
        cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
        turma = dict(cursor.fetchone())
        if not turma:
            flash("Turma não encontrada.", "erro")
            return redirect(url_for("turmas"))
        return render_template("editar_turma.html", turma=turma)
    except sqlite3.IntegrityError:
        flash("Já existe uma turma com este nome.", "erro")
        # Recarrega a turma para o template em caso de erro de integridade
        if conn:
            cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
            turma = dict(cursor.fetchone())
        return render_template("editar_turma.html", turma=turma)
    except Exception as e:
        flash(f"Erro ao editar turma: {e}", "erro")
        print(f"ERRO AO EDITAR TURMA: {e}")
        return redirect(url_for("turmas"))
    finally:
        if conn:
            conn.close()


@app.route("/turmas/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_turma(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se há alunos matriculados nesta turma
        cursor.execute("SELECT COUNT(*) FROM alunos WHERE turma_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir a turma, pois há alunos associados a ela.", "erro")
            return redirect(url_for("turmas"))

        cursor.execute("DELETE FROM turmas WHERE id=?", (id,))
        conn.commit()
        flash("Turma excluída com sucesso!", "sucesso")
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
    lista = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT d.*, u.nome as professor_nome
            FROM disciplinas d
            LEFT JOIN usuarios u ON d.professor_id = u.id
            ORDER BY d.nome
        """)
        lista = [dict(row) for row in cursor.fetchall()]
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
                (nome, descricao, professor_id if professor_id else None, tem_atividades, frequencia_minima, ativa))
            conn.commit()
            flash("Disciplina adicionada com sucesso!", "sucesso")
            return redirect(url_for("disciplinas"))
        return render_template("nova_disciplina.html", professores=professores)
    except sqlite3.IntegrityError:
        flash("Já existe uma disciplina com este nome.", "erro")
        return render_template("nova_disciplina.html", professores=professores)
    except Exception as e:
        flash(f"Erro ao adicionar disciplina: {e}", "erro")
        print(f"ERRO AO ADICIONAR DISCIPLINA: {e}")
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
                (nome, descricao, professor_id if professor_id else None, tem_atividades, frequencia_minima, ativa, id))
            conn.commit()
            flash("Disciplina atualizada com sucesso!", "sucesso")
            return redirect(url_for("disciplinas"))

        cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
        disciplina = dict(cursor.fetchone())
        if not disciplina:
            flash("Disciplina não encontrada.", "erro")
            return redirect(url_for("disciplinas"))
        return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except sqlite3.IntegrityError:
        flash("Já existe uma disciplina com este nome.", "erro")
        # Recarrega a disciplina para o template em caso de erro de integridade
        if conn:
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = dict(cursor.fetchone())
        return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except Exception as e:
        flash(f"Erro ao editar disciplina: {e}", "erro")
        print(f"ERRO AO EDITAR DISCIPLINA: {e}")
        return redirect(url_for("disciplinas"))
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
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir a disciplina, pois há matrículas associadas a ela.", "erro")
            return redirect(url_for("disciplinas"))

        cursor.execute("DELETE FROM disciplinas WHERE id=?", (id,))
        conn.commit()
        flash("Disciplina excluída com sucesso!", "sucesso")
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
    turmas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            nome            = request.form["nome"].strip()
            data_nascimento = request.form["data_nascimento"]
            telefone        = request.form["telefone"].strip()
            email           = request.form["email"].strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id")
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            cursor.execute(
                "INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (nome, data_nascimento, telefone, email, membro_igreja, turma_id if turma_id else None, nome_pai, nome_mae, endereco))
            conn.commit()
            flash("Aluno adicionado com sucesso!", "sucesso")
            return redirect(url_for("alunos"))
        return render_template("novo_aluno.html", turmas=turmas)
    except Exception as e:
        flash(f"Erro ao adicionar aluno: {e}", "erro")
        print(f"ERRO AO ADICIONAR ALUNO: {e}")
        return render_template("novo_aluno.html", turmas=turmas)
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_aluno(id):
    conn = None
    aluno = None
    turmas = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            nome            = request.form["nome"].strip()
            data_nascimento = request.form["data_nascimento"]
            telefone        = request.form["telefone"].strip()
            email           = request.form["email"].strip()
            membro_igreja   = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id        = request.form.get("turma_id")
            nome_pai        = request.form.get("nome_pai", "").strip()
            nome_mae        = request.form.get("nome_mae", "").strip()
            endereco        = request.form.get("endereco", "").strip()

            cursor.execute(
                "UPDATE alunos SET nome=?, data_nascimento=?, telefone=?, email=?, membro_igreja=?, turma_id=?, nome_pai=?, nome_mae=?, endereco=? WHERE id=?",
                (nome, data_nascimento, telefone, email, membro_igreja, turma_id if turma_id else None, nome_pai, nome_mae, endereco, id))
            conn.commit()
            flash("Aluno atualizado com sucesso!", "sucesso")
            return redirect(url_for("alunos"))

        cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
        aluno = dict(cursor.fetchone())
        if not aluno:
            flash("Aluno não encontrado.", "erro")
            return redirect(url_for("alunos"))
        return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
    except Exception as e:
        flash(f"Erro ao editar aluno: {e}", "erro")
        print(f"ERRO AO EDITAR ALUNO: {e}")
        return render_template("editar_aluno.html", aluno=aluno, turmas=turmas) # Tenta renderizar mesmo com erro
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_aluno(id):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        # Verificar se há matrículas associadas a este aluno
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE aluno_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir o aluno, pois há matrículas associadas a ele.", "erro")
            return redirect(url_for("alunos"))

        cursor.execute("DELETE FROM alunos WHERE id=?", (id,))
        conn.commit()
        flash("Aluno excluído com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir aluno: {e}", "erro")
        print(f"ERRO AO EXCLUIR ALUNO: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("alunos"))


# ══════════════════════════════════════
# PROFESSORES (NOVA ROTA)
# ══════════════════════════════════════
@app.route("/professores")
@login_required
def professores():
    conn = None
    lista_professores = []
    try:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, email FROM usuarios WHERE perfil='professor' ORDER BY nome")
        lista_professores = [dict(row) for row in cursor.fetchall()]
        return render_template("professores.html", professores=lista_professores)
    except Exception as e:
        flash(f"Erro ao carregar professores: {e}", "erro")
        print(f"ERRO EM PROFESSORES: {e}")
        return render_template("professores.html", professores=[])
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# USUÁRIOS (GERAL)
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
@admin_required
def usuarios():
    conn = None
    lista = []
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
        conn = None
        try:
            conn = conectar()
            cursor = conn.cursor()
            senha_hash = generate_password_hash(senha)
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
            print(f"ERRO AO ADICIONAR USUARIO: {e}")
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
        if request.method == "POST":
            nome    = request.form["nome"].strip()
            email   = request.form["email"].strip()
            perfil  = request.form["perfil"]
            senha   = request.form.get("senha") # Senha é opcional na edição

            # Não permitir que o admin logado altere seu próprio perfil para algo diferente de admin
            if current_user.id == id and perfil != 'admin':
                flash("Você não pode alterar seu próprio perfil de administrador.", "erro")
                # Recarrega o usuário para o template
                cursor.execute("SELECT id, nome, email, perfil FROM usuarios WHERE id=?", (id,))
                usuario = dict(cursor.fetchone())
                return render_template("editar_usuario.html", usuario=usuario)

            if senha: # Se uma nova senha foi fornecida
                senha_hash = generate_password_hash(senha)
                cursor.execute(
                    "UPDATE usuarios SET nome=?, email=?, perfil=?, senha_hash=? WHERE id=?",
                    (nome, email, perfil, senha_hash, id))
            else:
                cursor.execute(
                    "UPDATE usuarios SET nome=?, email=?, perfil=? WHERE id=?",
                    (nome, email, perfil, id))
            conn.commit()
            flash("Usuário atualizado com sucesso!", "sucesso")
            return redirect(url_for("usuarios"))

        cursor.execute("SELECT id, nome, email, perfil FROM usuarios WHERE id=?", (id,))
        usuario = dict(cursor.fetchone())
        if not usuario:
            flash("Usuário não encontrado.", "erro")
            return redirect(url_for("usuarios"))
        return render_template("editar_usuario.html", usuario=usuario)
    except sqlite3.IntegrityError:
        flash("Já existe um usuário com este e-mail.", "erro")
        # Recarrega o usuário para o template em caso de erro de integridade
        if conn:
            cursor.execute("SELECT id, nome, email, perfil FROM usuarios WHERE id=?", (id,))
            usuario = dict(cursor.fetchone())
        return render_template("editar_usuario.html", usuario=usuario)
    except Exception as e:
        flash(f"Erro ao editar usuário: {e}", "erro")
        print(f"ERRO AO EDITAR USUARIO: {e}")
        return redirect(url_for("usuarios"))
    finally:
        if conn:
            conn.close()


@app.route("/usuarios/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_usuario(id):
    conn = None
    try:
        # Não permitir que o admin logado se auto-exclua
        if current_user.id == id:
            flash("Você não pode excluir sua própria conta de administrador.", "erro")
            return redirect(url_for("usuarios"))

        conn = conectar()
        cursor = conn.cursor()
        # Verificar se o usuário é professor e está associado a alguma disciplina
        cursor.execute("SELECT perfil FROM usuarios WHERE id = ?", (id,))
        user_profile = cursor.fetchone()
        if user_profile and user_profile['perfil'] == 'professor':
            cursor.execute("SELECT COUNT(*) FROM disciplinas WHERE professor_id = ?", (id,))
            if cursor.fetchone()[0] > 0:
                flash("Não é possível excluir este professor, pois ele está associado a uma ou mais disciplinas.", "erro")
                return redirect(url_for("usuarios"))

        cursor.execute("DELETE FROM usuarios WHERE id=?", (id,))
        conn.commit()
        flash("Usuário excluído com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir usuário: {e}", "erro")
        print(f"ERRO AO EXCLUIR USUARIO: {e}")
    finally:
        if conn:
            conn.close()
    return redirect(url_for("usuarios"))


@app.route("/minha_conta", methods=["GET", "POST"])
@login_required
def minha_conta():
    if request.method == "POST":
        senha_atual = request.form.get("senha_atual", "")
        nova_senha  = request.form.get("nova_senha", "")
        confirmar   = request.form.get("confirmar", "")
        u = verificar_login(current_user.email, senha_atual)
        if not u:
            flash("Senha atual incorreta!", "erro")
            return redirect(url_for("minha_conta"))
        if nova_senha != confirmar:
            flash("As senhas não coincidem!", "erro")
            return redirect(url_for("minha_conta"))
        if len(nova_senha) < 6:
            flash("Mínimo 6 caracteres!", "erro")
            return redirect(url_for("minha_conta"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "UPDATE usuarios SET senha_hash=? WHERE id=?",
                (generate_password_hash(nova_senha), current_user.id))
            conn.commit()
            flash("Senha alterada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao alterar senha: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("index"))
    return render_template("minha_conta.html")


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