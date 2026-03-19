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
            nome        = request.form["nome"].strip()
            descricao   = request.form["descricao"].strip()
            faixa_etaria = request.form["faixa_etaria"]
            ativa       = 1 if request.form.get("ativa") == "on" else 0

            if not nome or not faixa_etaria:
                flash("Nome e Faixa Etária são obrigatórios!", "erro")
                return redirect(url_for("nova_turma"))

            conn   = conectar()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO turmas (nome, descricao, faixa_etaria, ativa) VALUES (?, ?, ?, ?)",
                (nome, descricao, faixa_etaria, ativa)
            )
            conn.commit()
            flash("Turma cadastrada com sucesso!", "sucesso")
            return redirect(url_for("turmas"))
        return render_template("nova_turma.html")
    except sqlite3.IntegrityError:
        flash("Já existe uma turma com este nome.", "erro")
        return render_template("nova_turma.html")
    except Exception as e:
        flash(f"Erro ao cadastrar turma: {e}", "erro")
        print(f"ERRO EM NOVA_TURMA: {e}")
        return render_template("nova_turma.html")
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
            nome        = request.form["nome"].strip()
            descricao   = request.form["descricao"].strip()
            faixa_etaria = request.form["faixa_etaria"]
            ativa       = 1 if request.form.get("ativa") == "on" else 0

            if not nome or not faixa_etaria:
                flash("Nome e Faixa Etária são obrigatórios!", "erro")
                return redirect(url_for("editar_turma", id=id))

            cursor.execute(
                "UPDATE turmas SET nome=?, descricao=?, faixa_etaria=?, ativa=? WHERE id=?",
                (nome, descricao, faixa_etaria, ativa, id)
            )
            conn.commit()
            flash("Turma atualizada com sucesso!", "sucesso")
            return redirect(url_for("turmas"))
        else:
            cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
            turma = cursor.fetchone()
            if not turma:
                flash("Turma não encontrada!", "erro")
                return redirect(url_for("turmas"))
            return render_template("editar_turma.html", turma=turma)
    except sqlite3.IntegrityError:
        flash("Já existe uma turma com este nome.", "erro")
        cursor.execute("SELECT * FROM turmas WHERE id=?", (id,)) # Recarregar para renderizar
        turma = cursor.fetchone()
        return render_template("editar_turma.html", turma=turma)
    except Exception as e:
        flash(f"Erro ao editar turma: {e}", "erro")
        print(f"ERRO EM EDITAR_TURMA: {e}")
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
        flash("Turma excluída com sucesso!", "sucesso")
        return redirect(url_for("turmas"))
    except Exception as e:
        flash(f"Erro ao excluir turma: {e}", "erro")
        print(f"ERRO EM EXCLUIR_TURMA: {e}")
        return redirect(url_for("turmas"))
    finally:
        if conn:
            conn.close()


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
            nome            = request.form["nome"].strip()
            descricao       = request.form["descricao"].strip()
            professor_id    = request.form.get("professor_id")
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = float(request.form.get("frequencia_minima", 75.0))
            ativa           = 1 if request.form.get("ativa") == "on" else 0

            if not nome:
                flash("Nome da disciplina é obrigatório!", "erro")
                # Recarregar professores para o template
                cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome")
                professores = cursor.fetchall()
                return render_template("nova_disciplina.html", professores=professores)

            cursor.execute(
                "INSERT INTO disciplinas (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa) VALUES (?, ?, ?, ?, ?, ?)",
                (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa)
            )
            conn.commit()
            flash("Disciplina cadastrada com sucesso!", "sucesso")
            return redirect(url_for("disciplinas"))
        else:
            cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome")
            professores = cursor.fetchall()
            return render_template("nova_disciplina.html", professores=professores)
    except sqlite3.IntegrityError:
        flash("Já existe uma disciplina com este nome.", "erro")
        cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome") # Recarregar
        professores = cursor.fetchall()
        return render_template("nova_disciplina.html", professores=professores)
    except Exception as e:
        flash(f"Erro ao cadastrar disciplina: {e}", "erro")
        print(f"ERRO EM NOVA_DISCIPLINA: {e}")
        cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome") # Recarregar
        professores = cursor.fetchall()
        return render_template("nova_disciplina.html", professores=professores)
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
            nome            = request.form["nome"].strip()
            descricao       = request.form["descricao"].strip()
            professor_id    = request.form.get("professor_id")
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = float(request.form.get("frequencia_minima", 75.0))
            ativa           = 1 if request.form.get("ativa") == "on" else 0

            if not nome:
                flash("Nome da disciplina é obrigatório!", "erro")
                # Recarregar dados para o template
                cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
                disciplina = cursor.fetchone()
                cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome")
                professores = cursor.fetchall()
                return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)

            cursor.execute(
                "UPDATE disciplinas SET nome=?, descricao=?, professor_id=?, tem_atividades=?, frequencia_minima=?, ativa=? WHERE id=?",
                (nome, descricao, professor_id, tem_atividades, frequencia_minima, ativa, id)
            )
            conn.commit()
            flash("Disciplina atualizada com sucesso!", "sucesso")
            return redirect(url_for("disciplinas"))
        else:
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            if not disciplina:
                flash("Disciplina não encontrada!", "erro")
                return redirect(url_for("disciplinas"))
            cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome")
            professores = cursor.fetchall()
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except sqlite3.IntegrityError:
        flash("Já existe uma disciplina com este nome.", "erro")
        # Recarregar dados para o template
        cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
        disciplina = cursor.fetchone()
        cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome")
        professores = cursor.fetchall()
        return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except Exception as e:
        flash(f"Erro ao editar disciplina: {e}", "erro")
        print(f"ERRO EM EDITAR_DISCIPLINA: {e}")
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
        flash("Disciplina excluída com sucesso!", "sucesso")
        return redirect(url_for("disciplinas"))
    except Exception as e:
        flash(f"Erro ao excluir disciplina: {e}", "erro")
        print(f"ERRO EM EXCLUIR_DISCIPLINA: {e}")
        return redirect(url_for("disciplinas"))
    finally:
        if conn:
            conn.close()


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
            SELECT a.*, t.nome as turma_nome
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
        turmas = [] # Inicializa turmas
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()

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

            if not nome or not data_nascimento:
                flash("Nome e Data de Nascimento são obrigatórios!", "erro")
                return render_template("novo_aluno.html", turmas=turmas)

            # Convertendo turma_id para None se for string vazia
            turma_id = int(turma_id) if turma_id else None

            cursor.execute(
                "INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco)
            )
            conn.commit()
            flash("Aluno cadastrado com sucesso!", "sucesso")
            return redirect(url_for("alunos"))
        return render_template("novo_aluno.html", turmas=turmas)
    except Exception as e:
        flash(f"Erro ao cadastrar aluno: {e}", "erro")
        print(f"ERRO EM NOVO_ALUNO: {e}")
        # Recarrega turmas em caso de erro para re-renderizar o formulário
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()
        return render_template("novo_aluno.html", turmas=turmas)
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
        turmas = [] # Inicializa turmas
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()

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

            if not nome or not data_nascimento:
                flash("Nome e Data de Nascimento são obrigatórios!", "erro")
                # Recarregar aluno para o template
                cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
                aluno = cursor.fetchone()
                return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)

            # Convertendo turma_id para None se for string vazia
            turma_id = int(turma_id) if turma_id else None

            cursor.execute(
                "UPDATE alunos SET nome=?, data_nascimento=?, telefone=?, email=?, membro_igreja=?, turma_id=?, nome_pai=?, nome_mae=?, endereco=? WHERE id=?",
                (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco, id)
            )
            conn.commit()
            flash("Aluno atualizado com sucesso!", "sucesso")
            return redirect(url_for("alunos"))
        else:
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            if not aluno:
                flash("Aluno não encontrado!", "erro")
                return redirect(url_for("alunos"))
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
    except Exception as e:
        flash(f"Erro ao editar aluno: {e}", "erro")
        print(f"ERRO EM EDITAR_ALUNO: {e}")
        # Recarrega turmas e aluno em caso de erro para re-renderizar o formulário
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()
        cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
        aluno = cursor.fetchone()
        return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
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
        # Verificar se o aluno tem matrículas
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE aluno_id = ?", (id,))
        total_matriculas = cursor.fetchone()[0]
        if total_matriculas > 0:
            flash(f"Não é possível excluir o aluno. Existem {total_matriculas} matrículas associadas a ele.", "erro")
            return redirect(url_for("alunos"))

        cursor.execute("DELETE FROM alunos WHERE id=?", (id,))
        conn.commit()
        flash("Aluno excluído com sucesso!", "sucesso")
        return redirect(url_for("alunos"))
    except Exception as e:
        flash(f"Erro ao excluir aluno: {e}", "erro")
        print(f"ERRO EM EXCLUIR_ALUNO: {e}")
        return redirect(url_for("alunos"))
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha_aluno(id):
    conn = None
    aluno_dict = None
    matriculas_processadas = []
    try:
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT a.*, t.nome as turma_nome, t.faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno = cursor.fetchone()

        if not aluno:
            flash("Aluno não encontrado!", "erro")
            return redirect(url_for("alunos"))

        aluno_dict = dict(aluno) # Converter para dict mutável

        cursor.execute("""
            SELECT m.id as matricula_id, m.data_inicio, m.data_conclusao, m.status,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                   d.nome as disciplina_nome, d.tem_atividades, d.frequencia_minima
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.aluno_id = ?
            ORDER BY d.nome
        """, (id,))
        matriculas = cursor.fetchall()

        for mat in matriculas:
            matricula_dict = dict(mat) # Converter para dict mutável

            # --- Cálculo de Frequência ---
            cursor.execute("""
                SELECT data_aula, presente, fez_atividade
                FROM presencas
                WHERE matricula_id = ?
                ORDER BY data_aula DESC
            """, (matricula_dict['matricula_id'],))
            historico_chamadas = [dict(c) for c in cursor.fetchall()] # Converter para dict

            presencas = sum(1 for c in historico_chamadas if c['presente'])
            total_aulas = len(historico_chamadas)
            atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas
            matricula_dict['historico_chamadas'] = historico_chamadas # Adicionar histórico

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            # --- Cálculo de Média e Status ---
            faixa_etaria = aluno_dict['faixa_etaria']

            nota_final_calc = None
            status = matricula_dict['status'] # Status padrão do DB

            if faixa_etaria in ['criancas_0_3', 'criancas_4_7', 'criancas_8_12']:
                # Crianças: Aprovado/Reprovado apenas por frequência
                if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status = "aprovado"
                else:
                    status = "reprovado"
                matricula_dict['media_display'] = "N/A"
                matricula_dict['status_display'] = f"{status.capitalize()} (Frequência: {frequencia_porcentagem:.1f}%)"
                matricula_dict['status_frequencia'] = status # Para uso no template
            else:
                # Adolescentes/Jovens/Adultos: Cálculo de notas
                if faixa_etaria in ['adolescentes_13_15', 'jovens_16_17']:
                    # Média para Adolescentes/Jovens
                    notas_validas = [n for n in [
                        matricula_dict['meditacao'],
                        matricula_dict['versiculos'],
                        matricula_dict['desafio_nota'],
                        matricula_dict['visitante']
                    ] if n is not None]
                    if notas_validas:
                        nota_final_calc = sum(notas_validas) / len(notas_validas)
                elif faixa_etaria == 'adultos':
                    # Média para Adultos
                    notas_validas = [n for n in [
                        matricula_dict['nota1'],
                        matricula_dict['nota2'],
                        matricula_dict['participacao'],
                        matricula_dict['desafio'],
                        matricula_dict['prova']
                    ] if n is not None]
                    if notas_validas:
                        nota_final_calc = sum(notas_validas) / len(notas_validas)

                matricula_dict['media_display'] = f"{nota_final_calc:.1f}" if nota_final_calc is not None else "—"

                # Lógica de status combinada (notas e frequência)
                if status == 'cursando':
                    if nota_final_calc is not None and nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                        status = "aprovado (provisório)" # Sugere aprovação, mas ainda cursando
                    elif nota_final_calc is not None and nota_final_calc < 7.0:
                        status = "reprovado (provisório)" # Sugere reprovação
                    elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                        status = "reprovado (provisório - frequência)"
                    else:
                        status = "cursando" # Ainda não há dados suficientes para definir

                matricula_dict['status_display'] = status.replace('_', ' ').title() # Formata para exibição
                matricula_dict['status_frequencia'] = None # Não aplicável para essas faixas etárias

            matriculas_processadas.append(matricula_dict)

        return render_template("trilha_aluno.html", aluno=aluno_dict, matriculas=matriculas_processadas)
    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        print(f"ERRO EM TRILHA_ALUNO: {e}")
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
                   m.data_inicio, m.data_conclusao, m.status,
                   t.nome as turma_nome, d.tem_atividades, d.frequencia_minima,
                   a.faixa_etaria -- Adicionado para lógica de notas
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            ORDER BY a.nome, d.nome
        """)
        raw_matriculas = cursor.fetchall()
        processed_matriculas = []

        for mat in raw_matriculas:
            matricula_dict = dict(mat) # Converter para dict mutável

            # --- Cálculo de Frequência ---
            cursor.execute("""
                SELECT presente, fez_atividade
                FROM presencas
                WHERE matricula_id = ?
            """, (matricula_dict['id'],))
            historico_chamadas = cursor.fetchall()

            presencas = sum(1 for c in historico_chamadas if c['presente'])
            total_aulas = len(historico_chamadas)
            atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            # --- Status de Frequência para Crianças ---
            if matricula_dict['faixa_etaria'] in ['criancas_0_3', 'criancas_4_7', 'criancas_8_12']:
                if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    matricula_dict['status_frequencia_display'] = "Aprovado"
                else:
                    matricula_dict['status_frequencia_display'] = "Reprovado"
            else:
                matricula_dict['status_frequencia_display'] = "N/A" # Não aplicável para outras faixas

            processed_matriculas.append(matricula_dict)

        return render_template("matriculas.html", matriculas=processed_matriculas)
    except Exception as e:
        flash(f"Erro ao carregar matrículas: {e}", "erro")
        print(f"ERRO EM MATRICULAS: {e}")
        return render_template("matriculas.html", matriculas=[])
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/nova_aluno_disciplina", methods=["GET", "POST"])
@login_required
def nova_aluno_disciplina():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        alunos = []
        disciplinas = []
        turmas = []

        cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
        alunos = cursor.fetchall()
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas = cursor.fetchall()
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()

        if request.method == "POST":
            aluno_existente_id = request.form.get("aluno_existente_id")
            disciplina_id      = request.form.get("disciplina_id")
            data_inicio        = request.form.get("data_inicio")
            status             = request.form.get("status", "cursando")

            # Dados para novo aluno, se aplicável
            novo_aluno_nome            = request.form.get("novo_aluno_nome", "").strip()
            novo_aluno_data_nascimento = request.form.get("novo_aluno_data_nascimento", "")
            novo_aluno_telefone        = request.form.get("novo_aluno_telefone", "").strip()
            novo_aluno_email           = request.form.get("novo_aluno_email", "").strip()
            novo_aluno_membro_igreja   = 1 if request.form.get("novo_aluno_membro_igreja") == "on" else 0
            novo_aluno_turma_id        = request.form.get("novo_aluno_turma_id")
            novo_aluno_nome_pai        = request.form.get("novo_aluno_nome_pai", "").strip()
            novo_aluno_nome_mae        = request.form.get("novo_aluno_nome_mae", "").strip()
            novo_aluno_endereco        = request.form.get("novo_aluno_endereco", "").strip()

            aluno_id_para_matricular = None

            if aluno_existente_id:
                aluno_id_para_matricular = int(aluno_existente_id)
            elif novo_aluno_nome and novo_aluno_data_nascimento:
                # Cadastrar novo aluno
                try:
                    novo_aluno_turma_id = int(novo_aluno_turma_id) if novo_aluno_turma_id else None
                    cursor.execute(
                        "INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        (novo_aluno_nome, novo_aluno_data_nascimento, novo_aluno_telefone, novo_aluno_email, novo_aluno_membro_igreja, novo_aluno_turma_id, novo_aluno_nome_pai, novo_aluno_nome_mae, novo_aluno_endereco)
                    )
                    conn.commit()
                    aluno_id_para_matricular = cursor.lastrowid
                    flash(f"Novo aluno '{novo_aluno_nome}' cadastrado com sucesso!", "sucesso")
                except Exception as e:
                    flash(f"Erro ao cadastrar novo aluno para matrícula: {e}", "erro")
                    print(f"ERRO EM NOVA_ALUNO_DISCIPLINA (cadastro de aluno): {e}")
                    return render_template("nova_aluno_disciplina.html", alunos=alunos, disciplinas=disciplinas, turmas=turmas)
            else:
                flash("Selecione um aluno existente ou preencha os dados do novo aluno.", "erro")
                return render_template("nova_aluno_disciplina.html", alunos=alunos, disciplinas=disciplinas, turmas=turmas)

            if not aluno_id_para_matricular or not disciplina_id or not data_inicio:
                flash("Aluno, Disciplina e Data de Início são obrigatórios para a matrícula!", "erro")
                return render_template("nova_aluno_disciplina.html", alunos=alunos, disciplinas=disciplinas, turmas=turmas)

            try:
                cursor.execute(
                    "INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, status) VALUES (?, ?, ?, ?)",
                    (aluno_id_para_matricular, disciplina_id, data_inicio, status)
                )
                conn.commit()
                flash("Matrícula realizada com sucesso!", "sucesso")
                return redirect(url_for("matriculas"))
            except sqlite3.IntegrityError:
                flash("Este aluno já está matriculado nesta disciplina.", "erro")
            except Exception as e:
                flash(f"Erro ao realizar matrícula: {e}", "erro")
                print(f"ERRO EM NOVA_ALUNO_DISCIPLINA (matrícula): {e}")

            return render_template("nova_aluno_disciplina.html", alunos=alunos, disciplinas=disciplinas, turmas=turmas)
        return render_template("nova_aluno_disciplina.html", alunos=alunos, disciplinas=disciplinas, turmas=turmas)
    except Exception as e:
        flash(f"Erro ao carregar página de matrícula: {e}", "erro")
        print(f"ERRO EM NOVA_ALUNO_DISCIPLINA (GET): {e}")
        return render_template("nova_aluno_disciplina.html", alunos=[], disciplinas=[], turmas=[])
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
            data_inicio    = request.form["data_inicio"]
            data_conclusao = request.form.get("data_conclusao")
            status         = request.form["status"]
            nota1          = request.form.get("nota1")
            nota2          = request.form.get("nota2")
            participacao   = request.form.get("participacao")
            desafio        = request.form.get("desafio")
            prova          = request.form.get("prova")
            meditacao      = request.form.get("meditacao")
            versiculos     = request.form.get("versiculos")
            desafio_nota   = request.form.get("desafio_nota")
            visitante      = request.form.get("visitante")

            # Converte strings vazias para None para o banco de dados
            data_conclusao = data_conclusao if data_conclusao else None
            nota1 = float(nota1) if nota1 else None
            nota2 = float(nota2) if nota2 else None
            participacao = float(participacao) if participacao else None
            desafio = float(desafio) if desafio else None
            prova = float(prova) if prova else None
            meditacao = float(meditacao) if meditacao else None
            versiculos = float(versiculos) if versiculos else None
            desafio_nota = float(desafio_nota) if desafio_nota else None
            visitante = float(visitante) if visitante else None

            cursor.execute(
                """UPDATE matriculas SET data_inicio=?, data_conclusao=?, status=?,
                   nota1=?, nota2=?, participacao=?, desafio=?, prova=?,
                   meditacao=?, versiculos=?, desafio_nota=?, visitante=? WHERE id=?""",
                (data_inicio, data_conclusao, status,
                 nota1, nota2, participacao, desafio, prova,
                 meditacao, versiculos, desafio_nota, visitante, id)
            )
            conn.commit()
            flash("Matrícula atualizada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        else:
            cursor.execute("""
                SELECT m.*, a.nome as aluno_nome, d.nome as disciplina_nome,
                       a.faixa_etaria, d.tem_atividades, d.frequencia_minima
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                WHERE m.id = ?
            """, (id,))
            matricula = cursor.fetchone()
            if not matricula:
                flash("Matrícula não encontrada!", "erro")
                return redirect(url_for("matriculas"))

            # Converter para dict mutável para adicionar campos
            matricula_dict = dict(matricula)

            # --- Cálculo de Frequência para exibição ---
            cursor.execute("""
                SELECT presente, fez_atividade
                FROM presencas
                WHERE matricula_id = ?
            """, (matricula_dict['id'],))
            historico_chamadas = cursor.fetchall()

            presencas = sum(1 for c in historico_chamadas if c['presente'])
            total_aulas = len(historico_chamadas)
            atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            return render_template("editar_matricula.html", matricula=matricula_dict)
    except Exception as e:
        flash(f"Erro ao editar matrícula: {e}", "erro")
        print(f"ERRO EM EDITAR_MATRICULA: {e}")
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
        # Excluir presenças associadas a esta matrícula primeiro
        cursor.execute("DELETE FROM presencas WHERE matricula_id = ?", (id,))
        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula excluída com sucesso!", "sucesso")
        return redirect(url_for("matriculas"))
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
        print(f"ERRO EM EXCLUIR_MATRICULA: {e}")
        return redirect(url_for("matriculas"))
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# PRESENÇA
# ══════════════════════════════════════
@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        disciplinas = []
        turmas = []
        alunos_matriculados = []
        chamadas_recentes = []

        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas = cursor.fetchall()
        cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()

        disciplina_id = request.args.get("disciplina_id", type=int)
        turma_id = request.args.get("turma_id", type=int)
        data_chamada = request.args.get("data_chamada", date.today().strftime("%Y-%m-%d"))

        if disciplina_id:
            query = """
                SELECT m.id as matricula_id, a.id as aluno_id, a.nome as aluno_nome,
                       p.presente, p.fez_atividade, p.id as presenca_id
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN presencas p ON p.matricula_id = m.id AND p.data_aula = ?
                WHERE m.disciplina_id = ?
            """
            params = [data_chamada, disciplina_id]
            if turma_id:
                query += " AND a.turma_id = ?"
                params.append(turma_id)
            query += " ORDER BY a.nome"
            cursor.execute(query, params)
            alunos_matriculados = [dict(row) for row in cursor.fetchall()]

        if request.method == "POST":
            disciplina_id_post = request.form.get("disciplina_id", type=int)
            turma_id_post = request.form.get("turma_id", type=int)
            data_chamada_post = request.form.get("data_chamada")

            if not disciplina_id_post or not data_chamada_post:
                flash("Disciplina e Data da Chamada são obrigatórios!", "erro")
                return redirect(url_for("chamada"))

            # Excluir presenças existentes para esta disciplina e data (se houver)
            # Isso evita duplicatas e permite reenvio da chamada
            cursor.execute("""
                DELETE FROM presencas
                WHERE data_aula = ? AND matricula_id IN (
                    SELECT id FROM matriculas WHERE disciplina_id = ?
                )
            """, (data_chamada_post, disciplina_id_post))
            conn.commit()

            for key, value in request.form.items():
                if key.startswith("presente_"):
                    matricula_id = int(key.replace("presente_", ""))
                    presente = 1 if value == "on" else 0
                    fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") == "on" else 0

                    cursor.execute(
                        "INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade) VALUES (?, ?, ?, ?)",
                        (matricula_id, data_chamada_post, presente, fez_atividade)
                    )
            conn.commit()
            flash("Chamada registrada com sucesso!", "sucesso")
            return redirect(url_for("chamada", disciplina_id=disciplina_id_post, turma_id=turma_id_post, data_chamada=data_chamada_post))

        # Histórico de chamadas recentes (últimas 10)
        cursor.execute("""
            SELECT p.data_aula, d.nome as disciplina_nome, t.nome as turma_nome,
                   COUNT(CASE WHEN p.presente = 1 THEN 1 END) as presentes,
                   COUNT(p.id) as total_alunos_chamada
            FROM presencas p
            JOIN matriculas m ON p.matricula_id = m.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            GROUP BY p.data_aula, d.nome, t.nome
            ORDER BY p.data_aula DESC
            LIMIT 10
        """)
        chamadas_recentes = [dict(row) for row in cursor.fetchall()]

        return render_template("chamada.html",
            disciplinas=disciplinas,
            turmas=turmas,
            alunos_matriculados=alunos_matriculados,
            disciplina_id=disciplina_id,
            turma_id=turma_id,
            data_chamada=data_chamada,
            chamadas_recentes=chamadas_recentes
        )
    except Exception as e:
        flash(f"Erro ao carregar página de chamada: {e}", "erro")
        print(f"ERRO EM CHAMADA: {e}")
        return render_template("chamada.html",
            disciplinas=[], turmas=[], alunos_matriculados=[],
            disciplina_id=None, turma_id=None, data_chamada=date.today().strftime("%Y-%m-%d"),
            chamadas_recentes=[]
        )
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
    disciplinas = []
    turmas = []
    alunos = []
    frequencia_data = []
    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    selected_turma_id = request.args.get("turma_id", type=int)
    selected_aluno_id = request.args.get("aluno_id", type=int)
    data_inicio_filtro = request.args.get("data_inicio_filtro")
    data_fim_filtro = request.args.get("data_fim_filtro")

    try:
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas = cursor.fetchall()
        cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
        turmas = cursor.fetchall()
        cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
        alunos = cursor.fetchall()

        query = """
            SELECT m.id as matricula_id, a.nome as aluno_nome, d.nome as disciplina_nome,
                   t.nome as turma_nome, d.tem_atividades, d.frequencia_minima,
                   a.faixa_etaria
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE 1=1
        """
        params = []

        if selected_disciplina_id:
            query += " AND d.id = ?"
            params.append(selected_disciplina_id)
        if selected_turma_id:
            query += " AND t.id = ?"
            params.append(selected_turma_id)
        if selected_aluno_id:
            query += " AND a.id = ?"
            params.append(selected_aluno_id)

        query += " ORDER BY d.nome, a.nome"
        cursor.execute(query, params)
        raw_matriculas = cursor.fetchall()

        processed_frequencia_data = []

        for mat in raw_matriculas:
            matricula_dict = dict(mat)

            presenca_query = """
                SELECT presente, fez_atividade, data_aula
                FROM presencas
                WHERE matricula_id = ?
            """
            presenca_params = [matricula_dict['matricula_id']]

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

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            processed_frequencia_data.append(matricula_dict)

        frequencia_data = processed_frequencia_data

        return render_template("relatorios_frequencia.html",
            disciplinas=disciplinas,
            turmas=turmas,
            alunos=alunos,
            frequencia_data=frequencia_data,
            selected_disciplina_id=selected_disciplina_id,
            selected_turma_id=selected_turma_id,
            selected_aluno_id=selected_aluno_id,
            data_inicio_filtro=data_inicio_filtro,
            data_fim_filtro=data_fim_filtro
        )
    except Exception as e:
        flash(f"Erro ao carregar relatório de frequência: {e}", "erro")
        print(f"ERRO NO RELATÓRIO DE FREQUÊNCIA: {e}")
        return render_template("relatorios_frequencia.html",
            disciplinas=disciplinas,
            turmas=turmas,
            alunos=alunos,
            frequencia_data=[],
            selected_disciplina_id=selected_disciplina_id,
            selected_turma_id=selected_turma_id,
            selected_aluno_id=selected_aluno_id,
            data_inicio_filtro=data_inicio_filtro,
            data_fim_filtro=data_fim_filtro
        )
    finally:
        if conn:
            conn.close()


@app.route("/relatorios/frequencia/download/<format>", methods=["GET"])
@login_required
def download_relatorio_frequencia(format):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        disciplina_id = request.args.get("disciplina_id", type=int)
        turma_id = request.args.get("turma_id", type=int)
        aluno_id = request.args.get("aluno_id", type=int)
        data_inicio_filtro = request.args.get("data_inicio_filtro")
        data_fim_filtro = request.args.get("data_fim_filtro")

        query = """
            SELECT m.id as matricula_id, a.nome as aluno_nome, d.nome as disciplina_nome,
                   t.nome as turma_nome, d.tem_atividades, d.frequencia_minima,
                   a.faixa_etaria
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE 1=1
        """
        params = []

        if disciplina_id:
            query += " AND d.id = ?"
            params.append(disciplina_id)
        if turma_id:
            query += " AND t.id = ?"
            params.append(turma_id)
        if aluno_id:
            query += " AND a.id = ?"
            params.append(aluno_id)

        query += " ORDER BY d.nome, a.nome"
        cursor.execute(query, params)
        raw_matriculas = cursor.fetchall()

        frequencia_data = []

        for mat in raw_matriculas:
            matricula_dict = dict(mat)

            presenca_query = """
                SELECT presente, fez_atividade, data_aula
                FROM presencas
                WHERE matricula_id = ?
            """
            presenca_params = [matricula_dict['matricula_id']]

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

            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

            matricula_dict['presencas'] = presencas
            matricula_dict['total_aulas'] = total_aulas
            matricula_dict['atividades_feitas'] = atividades_feitas
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            frequencia_data.append(matricula_dict)

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