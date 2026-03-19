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

# Chamada para inicializar o banco de dados
# Esta linha é crucial e deve ser executada antes de qualquer acesso ao DB
inicializar_banco()

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
        flash(f"Erro ao carregar dados do painel: {e}. O banco de dados pode não estar inicializado corretamente.", "erro")
        print(f"ERRO NO PAINEL (INDEX): {e}")
        return render_template("index.html",
            total_alunos=0, total_professores=0, total_disciplinas=0,
            total_turmas=0, aprovados=0, reprovados=0, cursando=0)
    except Exception as e:
        flash(f"Erro inesperado ao carregar painel: {e}", "erro")
        print(f"ERRO INESPERADO NO PAINEL (INDEX): {e}")
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
        print(f"ERRO NAS TURMAS: {e}")
        return render_template("turmas.html", turmas=[])
    finally:
        if conn:
            conn.close()


@app.route("/turmas/nova", methods=["GET", "POST"])
@login_required
def nova_turma():
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
                flash(f"Erro de integridade ao cadastrar turma: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar turma: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("turmas"))
    return render_template("nova_turma.html")


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
                    flash(f"Erro de integridade ao atualizar turma: {e}", "erro")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar turma: {e}", "erro")
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
        print(f"ERRO EM EDITAR TURMA: {e}")
        return redirect(url_for("turmas"))
    finally:
        if conn:
            conn.close()


@app.route("/turmas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_turma(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há alunos associados a esta turma
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
    finally:
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
        print(f"ERRO NAS DISCIPLINAS: {e}")
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
        professores = cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome").fetchall()

        if request.method == "POST":
            nome            = request.form.get("nome", "").strip()
            descricao       = request.form.get("descricao", "").strip()
            professor_id    = request.form.get("professor_id", type=int)
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = request.form.get("frequencia_minima", type=float)

            if not nome:
                flash("Nome é obrigatório!", "erro")
                return render_template("nova_disciplina.html", professores=professores)
            if frequencia_minima is None or not (0 <= frequencia_minima <= 100):
                flash("Frequência mínima deve ser um valor entre 0 e 100.", "erro")
                return render_template("nova_disciplina.html", professores=professores)

            try:
                cursor.execute(
                    "INSERT INTO disciplinas (nome, descricao, professor_id, tem_atividades, frequencia_minima) VALUES (?, ?, ?, ?, ?)",
                    (nome, descricao, professor_id, tem_atividades, frequencia_minima))
                conn.commit()
                flash(f"Disciplina '{nome}' criada!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "disciplinas.nome" in str(e):
                    flash("Já existe uma disciplina com este nome!", "erro")
                else:
                    flash(f"Erro de integridade ao cadastrar disciplina: {e}", "erro")
            except Exception as e:
                flash(f"Erro inesperado ao cadastrar disciplina: {e}", "erro")
            return redirect(url_for("disciplinas"))
        return render_template("nova_disciplina.html", professores=professores)
    except Exception as e:
        flash(f"Erro ao carregar página de nova disciplina: {e}", "erro")
        print(f"ERRO EM NOVA DISCIPLINA (GET): {e}")
        return render_template("nova_disciplina.html", professores=[])
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
        professores = cursor.execute("SELECT id, nome FROM usuarios WHERE perfil = 'professor' ORDER BY nome").fetchall()

        if request.method == "POST":
            nome            = request.form.get("nome", "").strip()
            descricao       = request.form.get("descricao", "").strip()
            professor_id    = request.form.get("professor_id", type=int)
            tem_atividades  = 1 if request.form.get("tem_atividades") == "on" else 0
            frequencia_minima = request.form.get("frequencia_minima", type=float)
            ativa           = 1 if request.form.get("ativa") == "on" else 0

            if not nome:
                flash("Nome é obrigatório!", "erro")
                return render_template("editar_disciplina.html", disciplina=request.form, professores=professores)
            if frequencia_minima is None or not (0 <= frequencia_minima <= 100):
                flash("Frequência mínima deve ser um valor entre 0 e 100.", "erro")
                return render_template("editar_disciplina.html", disciplina=request.form, professores=professores)

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
                    flash(f"Erro de integridade ao atualizar disciplina: {e}", "erro")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar disciplina: {e}", "erro")
            return redirect(url_for("disciplinas"))
        else:
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            if not disciplina:
                flash("Disciplina não encontrada!", "erro")
                return redirect(url_for("disciplinas"))
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores)
    except Exception as e:
        flash(f"Erro ao carregar/editar disciplina: {e}", "erro")
        print(f"ERRO EM EDITAR DISCIPLINA: {e}")
        return redirect(url_for("disciplinas"))
    finally:
        if conn:
            conn.close()


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há matrículas associadas a esta disciplina
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
    finally:
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
        print(f"ERRO NOS ALUNOS: {e}")
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
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()

        if request.method == "POST":
            nome          = request.form.get("nome", "").strip()
            data_nascimento_str = request.form.get("data_nascimento", "").strip()
            telefone      = request.form.get("telefone", "").strip()
            email         = request.form.get("email", "").strip()
            membro_igreja = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id      = request.form.get("turma_id", type=int)
            nome_pai      = request.form.get("nome_pai", "").strip()
            nome_mae      = request.form.get("nome_mae", "").strip()
            endereco      = request.form.get("endereco", "").strip()

            if not nome:
                flash("Nome é obrigatório!", "erro")
                return render_template("novo_aluno.html", turmas=turmas, aluno=request.form)

            # Validação de data de nascimento
            data_nascimento = None
            if data_nascimento_str:
                try:
                    data_nascimento = datetime.strptime(data_nascimento_str, '%Y-%m-%d').strftime('%Y-%m-%d')
                except ValueError:
                    flash("Formato de data de nascimento inválido. Use AAAA-MM-DD.", "erro")
                    return render_template("novo_aluno.html", turmas=turmas, aluno=request.form)

            try:
                cursor.execute(
                    "INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco))
                conn.commit()
                flash(f"Aluno '{nome}' cadastrado!", "sucesso")
            except Exception as e:
                flash(f"Erro inesperado ao cadastrar aluno: {e}", "erro")
                print(f"ERRO AO CADASTRAR ALUNO: {e}")
            return redirect(url_for("alunos"))
        return render_template("novo_aluno.html", turmas=turmas, aluno={})
    except Exception as e:
        flash(f"Erro ao carregar página de novo aluno: {e}", "erro")
        print(f"ERRO EM NOVO ALUNO (GET): {e}")
        return render_template("novo_aluno.html", turmas=[], aluno={})
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
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()

        if request.method == "POST":
            nome          = request.form.get("nome", "").strip()
            data_nascimento_str = request.form.get("data_nascimento", "").strip()
            telefone      = request.form.get("telefone", "").strip()
            email         = request.form.get("email", "").strip()
            membro_igreja = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id      = request.form.get("turma_id", type=int)
            nome_pai      = request.form.get("nome_pai", "").strip()
            nome_mae      = request.form.get("nome_mae", "").strip()
            endereco      = request.form.get("endereco", "").strip()

            if not nome:
                flash("Nome é obrigatório!", "erro")
                # Recarrega o template com os dados do formulário e turmas
                aluno_data = request.form.to_dict()
                aluno_data['id'] = id # Garante que o ID esteja presente para o template
                return render_template("editar_aluno.html", aluno=aluno_data, turmas=turmas)

            # Validação de data de nascimento
            data_nascimento = None
            if data_nascimento_str:
                try:
                    data_nascimento = datetime.strptime(data_nascimento_str, '%Y-%m-%d').strftime('%Y-%m-%d')
                except ValueError:
                    flash("Formato de data de nascimento inválido. Use AAAA-MM-DD.", "erro")
                    aluno_data = request.form.to_dict()
                    aluno_data['id'] = id
                    return render_template("editar_aluno.html", aluno=aluno_data, turmas=turmas)

            try:
                cursor.execute(
                    "UPDATE alunos SET nome=?, data_nascimento=?, telefone=?, email=?, membro_igreja=?, turma_id=?, nome_pai=?, nome_mae=?, endereco=? WHERE id=?",
                    (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco, id))
                conn.commit()
                flash(f"Aluno '{nome}' atualizado!", "sucesso")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar aluno: {e}", "erro")
                print(f"ERRO AO ATUALIZAR ALUNO: {e}")
            return redirect(url_for("alunos"))
        else:
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            if not aluno:
                flash("Aluno não encontrado!", "erro")
                return redirect(url_for("alunos"))
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas)
    except Exception as e:
        flash(f"Erro ao carregar/editar aluno: {e}", "erro")
        print(f"ERRO EM EDITAR ALUNO: {e}")
        return redirect(url_for("alunos"))
    finally:
        if conn:
            conn.close()


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há matrículas associadas a este aluno
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
    finally:
        conn.close()
    return redirect(url_for("alunos"))


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha_aluno(id):
    conn = None
    aluno_dict = None
    matriculas_processadas = []
    try:
        conn = conectar()
        cursor = conn.cursor()

        # Obter dados do aluno
        cursor.execute("""
            SELECT a.*, t.nome as turma_nome, t.faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno = cursor.fetchone()

        if not aluno:
            flash("Aluno não encontrado!", "erro")
            return render_template("trilha_aluno.html", aluno=None)

        aluno_dict = dict(aluno) # Converter para dict mutável

        # Obter matrículas do aluno
        cursor.execute("""
            SELECT m.id AS matricula_id,
                   d.nome AS disciplina_nome,
                   d.tem_atividades,
                   d.frequencia_minima,
                   m.data_inicio, m.data_conclusao, m.status,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                   t.faixa_etaria AS turma_faixa_etaria
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
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
            matricula_dict['frequencia_display'] = f"{frequencia_porcentagem:.1f}%"

            # --- Cálculo de Notas e Status ---
            faixa_etaria = matricula_dict['turma_faixa_etaria'] or 'adultos' # Default para adultos

            nota_final_calc = None
            status = matricula_dict['status'] # Manter status do DB como base

            if 'criancas' in faixa_etaria:
                # Crianças não têm notas, apenas frequência
                matricula_dict['media_display'] = 'N/A'
                if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    matricula_dict['status_frequencia'] = 'Aprovado'
                else:
                    matricula_dict['status_frequencia'] = 'Reprovado'
                matricula_dict['status_display'] = matricula_dict['status_frequencia'] # Para crianças, o status é a frequência
            elif 'adolescentes' in faixa_etaria or 'jovens' in faixa_etaria:
                meditacao = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                versiculos = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                desafio_nota = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                visitante = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
                # Total 10 pontos: Meditação (4), Versículos (4), Desafio (1), Visitante (1)
                nota_final_calc = meditacao + versiculos + desafio_nota + visitante
                matricula_dict['media_display'] = f"{nota_final_calc:.1f}"
                if nota_final_calc >= 6 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status = 'aprovado'
                elif nota_final_calc < 6 and frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status = 'reprovado'
                elif nota_final_calc < 6:
                    status = 'reprovado'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status = 'reprovado'
                matricula_dict['status_display'] = status.capitalize()
            else: # Adultos
                nota1 = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                nota2 = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
                participacao = matricula_dict['participacao'] if matricula_dict['participacao'] is not None else 0
                desafio = matricula_dict['desafio'] if matricula_dict['desafio'] is not None else 0
                prova = matricula_dict['prova'] if matricula_dict['prova'] is not None else 0
                # Média simples entre Nota 1 e Nota 2, com bônus
                # Nota 1 (0-10), Nota 2 (0-10), Participação (0-1), Desafio (0-1), Prova (0-8)
                # Vamos considerar que Nota 1 e Nota 2 já incorporam outros critérios ou são as principais
                # E os outros são bônus ou parte da composição.
                # Para simplificar, vamos considerar Nota 1 e Nota 2 como as principais, e os outros como bônus.
                # Ou, se Nota 1 e Nota 2 são as notas de prova, e os outros são componentes.
                # Vamos assumir uma média ponderada ou soma para um total de 10.
                # Ex: (Nota1 * 0.4) + (Nota2 * 0.4) + (Participacao * 0.1) + (Desafio * 0.1)
                # Ou, se Nota1 e Nota2 são notas de prova, e os outros são bônus.
                # Vamos usar a lógica anterior: Nota 1 + Participação + Desafio + Prova (total 10)
                nota_final_calc = nota1 + participacao + desafio + prova # Total 10 pontos
                matricula_dict['media_display'] = f"{nota_final_calc:.1f}"
                if nota_final_calc >= 6 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status = 'aprovado'
                elif nota_final_calc < 6 and frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status = 'reprovado'
                elif nota_final_calc < 6:
                    status = 'reprovado'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status = 'reprovado'
                matricula_dict['status_display'] = status.capitalize()

            # Se o status da matrícula no DB for 'aprovado' ou 'reprovado', ele prevalece
            # Apenas se for 'cursando', a lógica acima define o status provisório
            if matricula_dict['status'] == 'aprovado':
                matricula_dict['status_display'] = 'Aprovado'
            elif matricula_dict['status'] == 'reprovado':
                matricula_dict['status_display'] = 'Reprovado'
            elif matricula_dict['status'] == 'trancado':
                matricula_dict['status_display'] = 'Trancado'
            elif matricula_dict['status'] == 'cursando' and nota_final_calc is not None:
                # Se ainda está cursando, mas já tem nota final calculada, mostra o status provisório
                matricula_dict['status_display'] += ' (Provisório)'


            matriculas_processadas.append(matricula_dict)

        return render_template("trilha_aluno.html", aluno=aluno_dict, matriculas=matriculas_processadas)

    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        print(f"ERRO NA TRILHA DO ALUNO: {e}")
        return render_template("trilha_aluno.html", aluno=None, matriculas=[])
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
            SELECT m.id, a.nome as aluno_nome, d.nome as disciplina_nome, m.data_inicio, m.data_conclusao, m.status
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            ORDER BY a.nome, d.nome
        """)
        lista = cursor.fetchall()
        return render_template("matriculas.html", matriculas=lista)
    except Exception as e:
        flash(f"Erro ao carregar matrículas: {e}", "erro")
        print(f"ERRO NAS MATRÍCULAS: {e}")
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
        alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()

        if request.method == "POST":
            aluno_id      = request.form.get("aluno_id", type=int)
            disciplina_id = request.form.get("disciplina_id", type=int)
            data_inicio   = request.form.get("data_inicio", "").strip()

            if not aluno_id or not disciplina_id or not data_inicio:
                flash("Todos os campos são obrigatórios!", "erro")
                return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas, matricula=request.form)

            try:
                cursor.execute(
                    "INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio) VALUES (?, ?, ?)",
                    (aluno_id, disciplina_id, data_inicio))
                conn.commit()
                flash("Matrícula criada com sucesso!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                    flash("Este aluno já está matriculado nesta disciplina!", "erro")
                else:
                    flash(f"Erro de integridade ao matricular: {e}", "erro")
            except Exception as e:
                flash(f"Erro inesperado ao matricular aluno: {e}", "erro")
                print(f"ERRO AO CRIAR MATRÍCULA: {e}")
            return redirect(url_for("matriculas"))
        return render_template("nova_matricula.html", alunos=alunos, disciplinas=disciplinas, matricula={})
    except Exception as e:
        flash(f"Erro ao carregar página de nova matrícula: {e}", "erro")
        print(f"ERRO EM NOVA MATRÍCULA (GET): {e}")
        return render_template("nova_matricula.html", alunos=[], disciplinas=[], matricula={})
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
        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()

        if request.method == "POST":
            # Dados do Aluno
            nome_aluno = request.form.get("nome_aluno", "").strip()
            data_nascimento_str = request.form.get("data_nascimento", "").strip()
            telefone = request.form.get("telefone", "").strip()
            email = request.form.get("email", "").strip()
            membro_igreja = 1 if request.form.get("membro_igreja") == "on" else 0
            turma_id = request.form.get("turma_id", type=int)
            nome_pai = request.form.get("nome_pai", "").strip()
            nome_mae = request.form.get("nome_mae", "").strip()
            endereco = request.form.get("endereco", "").strip()

            # Dados da Matrícula
            disciplina_id = request.form.get("disciplina_id", type=int)
            data_inicio = request.form.get("data_inicio", "").strip()

            if not nome_aluno or not disciplina_id or not data_inicio:
                flash("Nome do aluno, disciplina e data de início são obrigatórios!", "erro")
                return render_template("novo_aluno_disciplina.html", disciplinas=disciplinas, turmas=turmas, form_data=request.form)

            # Validação de data de nascimento
            data_nascimento = None
            if data_nascimento_str:
                try:
                    data_nascimento = datetime.strptime(data_nascimento_str, '%Y-%m-%d').strftime('%Y-%m-%d')
                except ValueError:
                    flash("Formato de data de nascimento inválido. Use AAAA-MM-DD.", "erro")
                    return render_template("novo_aluno_disciplina.html", disciplinas=disciplinas, turmas=turmas, form_data=request.form)

            try:
                # 1. Cadastrar o novo aluno
                cursor.execute(
                    "INSERT INTO alunos (nome, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (nome_aluno, data_nascimento, telefone, email, membro_igreja, turma_id, nome_pai, nome_mae, endereco))
                aluno_id = cursor.lastrowid # Obter o ID do aluno recém-criado

                # 2. Matricular o aluno na disciplina
                cursor.execute(
                    "INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio) VALUES (?, ?, ?)",
                    (aluno_id, disciplina_id, data_inicio))
                conn.commit()
                flash(f"Aluno '{nome_aluno}' cadastrado e matriculado em disciplina!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                    flash("Este aluno já está matriculado nesta disciplina!", "erro")
                else:
                    flash(f"Erro de integridade ao cadastrar/matricular: {e}", "erro")
            except Exception as e:
                flash(f"Erro inesperado ao cadastrar/matricular: {e}", "erro")
                print(f"ERRO EM NOVO ALUNO/DISCIPLINA: {e}")
            return redirect(url_for("matriculas"))
        return render_template("novo_aluno_disciplina.html", disciplinas=disciplinas, turmas=turmas, form_data={})
    except Exception as e:
        flash(f"Erro ao carregar página de cadastro/matrícula: {e}", "erro")
        print(f"ERRO EM NOVO ALUNO/DISCIPLINA (GET): {e}")
        return render_template("novo_aluno_disciplina.html", disciplinas=[], turmas=[], form_data={})
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
        matricula = None
        if request.method == "POST":
            status        = request.form.get("status", "").strip()
            data_conclusao = request.form.get("data_conclusao", "").strip()
            # Notas para Adultos
            nota1         = request.form.get("nota1", type=float)
            nota2         = request.form.get("nota2", type=float)
            participacao  = request.form.get("participacao", type=float)
            desafio       = request.form.get("desafio", type=float)
            prova         = request.form.get("prova", type=float)
            # Notas para Adolescentes/Jovens
            meditacao     = request.form.get("meditacao", type=float)
            versiculos    = request.form.get("versiculos", type=float)
            desafio_nota  = request.form.get("desafio_nota", type=float)
            visitante     = request.form.get("visitante", type=float)

            # Obter a faixa etária da turma do aluno para saber qual lógica de nota aplicar
            cursor.execute("""
                SELECT t.faixa_etaria
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.id = ?
            """, (id,))
            faixa_etaria_matricula = cursor.fetchone()
            faixa_etaria_matricula = faixa_etaria_matricula['faixa_etaria'] if faixa_etaria_matricula else 'adultos' # Default para adultos

            # Resetar todas as notas para NULL antes de aplicar as específicas
            update_query = """
                UPDATE matriculas SET status=?, data_conclusao=?,
                nota1=NULL, nota2=NULL, participacao=NULL, desafio=NULL, prova=NULL,
                meditacao=NULL, versiculos=NULL, desafio_nota=NULL, visitante=NULL
                WHERE id=?
            """
            params = [status, data_conclusao if data_conclusao else None, id]
            cursor.execute(update_query, params)

            # Aplicar notas específicas da faixa etária
            if 'adolescentes' in faixa_etaria_matricula or 'jovens' in faixa_etaria_matricula:
                cursor.execute("""
                    UPDATE matriculas SET meditacao=?, versiculos=?, desafio_nota=?, visitante=?
                    WHERE id=?
                """, (meditacao, versiculos, desafio_nota, visitante, id))
            elif 'adultos' in faixa_etaria_matricula:
                cursor.execute("""
                    UPDATE matriculas SET nota1=?, nota2=?, participacao=?, desafio=?, prova=?
                    WHERE id=?
                """, (nota1, nota2, participacao, desafio, prova, id))
            # Crianças não têm notas para atualizar aqui

            conn.commit()
            flash("Matrícula atualizada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        else:
            cursor.execute("""
                SELECT m.*, a.nome as aluno_nome, d.nome as disciplina_nome, t.faixa_etaria
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.id = ?
            """, (id,))
            matricula = cursor.fetchone()
            if not matricula:
                flash("Matrícula não encontrada!", "erro")
                return redirect(url_for("matriculas"))
            return render_template("editar_matricula.html", matricula=matricula)
    except Exception as e:
        flash(f"Erro ao carregar/editar matrícula: {e}", "erro")
        print(f"ERRO EM EDITAR MATRÍCULA: {e}")
        return redirect(url_for("matriculas"))
    finally:
        if conn:
            conn.close()


@app.route("/matriculas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Excluir presenças associadas primeiro
        cursor.execute("DELETE FROM presencas WHERE matricula_id = ?", (id,))
        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula e presenças associadas excluídas!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# PRESENÇAS
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

        alunos_matriculados = []
        disciplina_selecionada = None
        turma_selecionada = None
        data_chamada = request.form.get("data_chamada", date.today().strftime('%Y-%m-%d'))

        if request.method == "POST":
            disciplina_id = request.form.get("disciplina_id", type=int)
            turma_id = request.form.get("turma_id", type=int)
            data_chamada = request.form.get("data_chamada", date.today().strftime('%Y-%m-%d'))

            if disciplina_id and turma_id and data_chamada:
                disciplina_selecionada = dict(cursor.execute("SELECT id, nome, tem_atividades FROM disciplinas WHERE id=?", (disciplina_id,)).fetchone())
                turma_selecionada = dict(cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE id=?", (turma_id,)).fetchone())

                # Buscar alunos matriculados na disciplina E pertencentes à turma selecionada
                cursor.execute("""
                    SELECT a.id AS aluno_id, a.nome AS aluno_nome, m.id AS matricula_id,
                           p.presente, p.fez_atividade
                    FROM alunos a
                    JOIN matriculas m ON a.id = m.aluno_id
                    LEFT JOIN presencas p ON m.id = p.matricula_id AND p.data_aula = ?
                    WHERE m.disciplina_id = ? AND a.turma_id = ?
                    ORDER BY a.nome
                """, (data_chamada, disciplina_id, turma_id))
                alunos_matriculados = [dict(row) for row in cursor.fetchall()]

                # Se a requisição for para salvar a chamada
                if 'salvar_chamada' in request.form:
                    for aluno in alunos_matriculados:
                        matricula_id = aluno['matricula_id']
                        presente = 1 if request.form.get(f"presente_{matricula_id}") == "on" else 0
                        fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") == "on" else 0

                        # Verificar se já existe um registro de presença para esta matrícula e data
                        cursor.execute("""
                            SELECT id FROM presencas
                            WHERE matricula_id = ? AND data_aula = ?
                        """, (matricula_id, data_chamada))
                        presenca_existente = cursor.fetchone()

                        if presenca_existente:
                            # Atualizar presença existente
                            cursor.execute("""
                                UPDATE presencas SET presente=?, fez_atividade=?
                                WHERE id=?
                            """, (presente, fez_atividade, presenca_existente['id']))
                        else:
                            # Inserir nova presença
                            cursor.execute("""
                                INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                                VALUES (?, ?, ?, ?)
                            """, (matricula_id, data_chamada, presente, fez_atividade))
                    conn.commit()
                    flash("Chamada salva com sucesso!", "sucesso")
                    # Recarregar a lista de alunos para refletir as mudanças
                    cursor.execute("""
                        SELECT a.id AS aluno_id, a.nome AS aluno_nome, m.id AS matricula_id,
                               p.presente, p.fez_atividade
                        FROM alunos a
                        JOIN matriculas m ON a.id = m.aluno_id
                        LEFT JOIN presencas p ON m.id = p.matricula_id AND p.data_aula = ?
                        WHERE m.disciplina_id = ? AND a.turma_id = ?
                        ORDER BY a.nome
                    """, (data_chamada, disciplina_id, turma_id))
                    alunos_matriculados = [dict(row) for row in cursor.fetchall()]
            else:
                flash("Por favor, selecione uma disciplina, turma e data para fazer a chamada.", "aviso")

        # Histórico de chamadas recentes para a disciplina e turma selecionadas
        historico_chamadas_recentes = []
        if disciplina_selecionada and turma_selecionada:
            cursor.execute("""
                SELECT p.data_aula, COUNT(p.id) AS total_registros,
                       SUM(CASE WHEN p.presente = 1 THEN 1 ELSE 0 END) AS total_presentes
                FROM presencas p
                JOIN matriculas m ON p.matricula_id = m.id
                JOIN alunos a ON m.aluno_id = a.id
                WHERE m.disciplina_id = ? AND a.turma_id = ?
                GROUP BY p.data_aula
                ORDER BY p.data_aula DESC
                LIMIT 10 -- Mostrar as 10 chamadas mais recentes
            """, (disciplina_selecionada['id'], turma_selecionada['id']))
            historico_chamadas_recentes = [dict(row) for row in cursor.fetchall()]


        return render_template("chamada.html",
                               disciplinas=disciplinas,
                               turmas=turmas,
                               alunos_matriculados=alunos_matriculados,
                               disciplina_selecionada=disciplina_selecionada,
                               turma_selecionada=turma_selecionada,
                               data_chamada=data_chamada,
                               historico_chamadas_recentes=historico_chamadas_recentes)
    except Exception as e:
        flash(f"Erro ao carregar página de chamada: {e}", "erro")
        print(f"ERRO NA CHAMADA: {e}")
        return render_template("chamada.html",
                               disciplinas=[],
                               turmas=[],
                               alunos_matriculados=[],
                               disciplina_selecionada=None,
                               turma_selecionada=None,
                               data_chamada=date.today().strftime('%Y-%m-%d'),
                               historico_chamadas_recentes=[])
    finally:
        if conn:
            conn.close()


# ══════════════════════════════════════
# RELATÓRIOS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id=None, turma_id=None, aluno_id=None, status=None, data_inicio=None, data_fim=None):
    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()

        query = """
            SELECT m.id AS matricula_id,
                   a.nome AS aluno_nome,
                   d.nome AS disciplina_nome,
                   t.nome AS turma_nome,
                   t.faixa_etaria AS turma_faixa_etaria,
                   m.data_inicio, m.data_conclusao, m.status,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                   d.tem_atividades, d.frequencia_minima
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
        if data_inicio:
            query += " AND m.data_inicio >= ?"
            params.append(data_inicio)
        if data_fim:
            query += " AND m.data_inicio <= ?"
            params.append(data_fim)

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
            matricula_dict['frequencia_display'] = f"{frequencia_porcentagem:.1f}%"

            # --- Cálculo de Notas e Status (simplificado para relatório) ---
            faixa_etaria = matricula_dict['turma_faixa_etaria'] or 'adultos' # Default para adultos
            nota_final_calc = None

            if 'criancas' in faixa_etaria:
                matricula_dict['media_final'] = 'N/A'
            elif 'adolescentes' in faixa_etaria or 'jovens' in faixa_etaria:
                meditacao = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                versiculos = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                desafio_nota = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                visitante = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
                nota_final_calc = meditacao + versiculos + desafio_nota + visitante
                matricula_dict['media_final'] = f"{nota_final_calc:.1f}"
            else: # Adultos
                nota1 = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                nota2 = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
                participacao = matricula_dict['participacao'] if matricula_dict['participacao'] is not None else 0
                desafio = matricula_dict['desafio'] if matricula_dict['desafio'] is not None else 0
                prova = matricula_dict['prova'] if matricula_dict['prova'] is not None else 0
                nota_final_calc = nota1 + participacao + desafio + prova
                matricula_dict['media_final'] = f"{nota_final_calc:.1f}"

            # Status de aprovação/reprovação para o relatório
            if 'criancas' in faixa_etaria:
                matricula_dict['status_relatorio'] = 'Aprovado' if frequencia_porcentagem >= matricula_dict['frequencia_minima'] else 'Reprovado'
            elif nota_final_calc is not None:
                if nota_final_calc >= 6 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    matricula_dict['status_relatorio'] = 'Aprovado'
                else:
                    matricula_dict['status_relatorio'] = 'Reprovado'
            else:
                matricula_dict['status_relatorio'] = matricula_dict['status'].capitalize() # Usa o status do DB se não houver cálculo

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
        conn.close()
        return render_template("relatorios.html", disciplinas=disciplinas, turmas=turmas, alunos=alunos)
    except Exception as e:
        flash(f"Erro ao carregar filtros de relatório: {e}", "erro")
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
    data_inicio = request.form.get("data_inicio", "").strip()
    data_fim = request.form.get("data_fim", "").strip()

    matriculas_relatorio = get_relatorio_data(disciplina_id, turma_id, aluno_id, status, data_inicio, data_fim)

    conn = None
    try:
        conn = conectar()
        cursor = conn.cursor()
        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
        alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
        conn.close()

        return render_template("relatorios.html",
                               disciplinas=disciplinas,
                               turmas=turmas,
                               alunos=alunos,
                               matriculas_relatorio=matriculas_relatorio,
                               filtro_disciplina_id=disciplina_id,
                               filtro_turma_id=turma_id,
                               filtro_aluno_id=aluno_id,
                               filtro_status=status,
                               filtro_data_inicio=data_inicio,
                               filtro_data_fim=data_fim)
    except Exception as e:
        flash(f"Erro ao gerar relatório: {e}", "erro")
        print(f"ERRO EM GERAR RELATORIO: {e}")
        return redirect(url_for("relatorios"))
    finally:
        if conn:
            conn.close()


@app.route("/relatorios/frequencia", methods=["GET", "POST"])
@login_required
def relatorios_frequencia():
    conn = None
    disciplinas = []
    turmas = []
    alunos = []
    frequencia_data = []
    filtro_disciplina_id = None
    filtro_turma_id = None
    filtro_aluno_id = None
    filtro_data_inicio = None
    filtro_data_fim = None

    try:
        conn = conectar()
        cursor = conn.cursor()

        disciplinas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
        turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
        alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()

        if request.method == "POST":
            filtro_disciplina_id = request.form.get("disciplina_id", type=int)
            filtro_turma_id = request.form.get("turma_id", type=int)
            filtro_aluno_id = request.form.get("aluno_id", type=int)
            filtro_data_inicio = request.form.get("data_inicio", "").strip()
            filtro_data_fim = request.form.get("data_fim", "").strip()

            query = """
                SELECT a.nome AS aluno_nome,
                       d.nome AS disciplina_nome,
                       t.nome AS turma_nome,
                       m.id AS matricula_id,
                       d.tem_atividades,
                       d.frequencia_minima
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE 1=1
            """
            params = []

            if filtro_disciplina_id:
                query += " AND m.disciplina_id = ?"
                params.append(filtro_disciplina_id)
            if filtro_turma_id:
                query += " AND t.id = ?"
                params.append(filtro_turma_id)
            if filtro_aluno_id:
                query += " AND a.id = ?"
                params.append(filtro_aluno_id)

            query += " ORDER BY a.nome, d.nome"
            cursor.execute(query, params)
            raw_matriculas = cursor.fetchall()

            for mat in raw_matriculas:
                matricula_dict = dict(mat) # Converter para dict mutável

                # Contar presenças e total de aulas dentro do período filtrado
                presenca_query = """
                    SELECT presente, fez_atividade
                    FROM presencas
                    WHERE matricula_id = ?
                """
                presenca_params = [matricula_dict['matricula_id']]

                if filtro_data_inicio:
                    presenca_query += " AND data_aula >= ?"
                    presenca_params.append(filtro_data_inicio)
                if filtro_data_fim:
                    presenca_query += " AND data_aula <= ?"
                    presenca_params.append(filtro_data_fim)

                cursor.execute(presenca_query, presenca_params)
                historico_chamadas = cursor.fetchall()

                presencas = sum(1 for c in historico_chamadas if c['presente'])
                total_aulas = len(historico_chamadas)
                atividades_feitas = sum(1 for c in historico_chamadas if c['fez_atividade'])

                frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

                matricula_dict['presencas'] = presencas
                matricula_dict['total_aulas'] = total_aulas
                matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem
                matricula_dict['frequencia_display'] = f"{frequencia_porcentagem:.1f}%"
                matricula_dict['atividades_feitas'] = atividades_feitas

                frequencia_data.append(matricula_dict)

    except Exception as e:
        flash(f"Erro ao carregar relatórios de frequência: {e}", "erro")
        print(f"ERRO EM RELATORIOS_FREQUENCIA: {e}")
    finally:
        if conn:
            conn.close()

    return render_template("relatorio_frequencia.html",
                           disciplinas=disciplinas,
                           turmas=turmas,
                           alunos=alunos,
                           frequencia_data=frequencia_data,
                           filtro_disciplina_id=filtro_disciplina_id,
                           filtro_turma_id=filtro_turma_id,
                           filtro_aluno_id=filtro_aluno_id,
                           filtro_data_inicio=filtro_data_inicio,
                           filtro_data_fim=filtro_data_fim)


@app.route("/relatorios/frequencia/download/<format>", methods=["GET"])
@login_required
def download_relatorio_frequencia(format):
    disciplina_id = request.args.get("disciplina_id", type=int)
    turma_id = request.args.get("turma_id", type=int)
    aluno_id = request.args.get("aluno_id", type=int)
    data_inicio = request.args.get("data_inicio", "").strip()
    data_fim = request.args.get("data_fim", "").strip()

    frequencia_data = []
    try:
        frequencia_data = get_relatorio_data(disciplina_id, turma_id, aluno_id, None, data_inicio, data_fim)
    except Exception as e:
        flash(f"Erro ao preparar dados para download: {e}", "erro")
        print(f"ERRO EM DOWNLOAD_RELATORIO_FREQUENCIA (GET_RELATORIO_DATA): {e}")
        return redirect(url_for("relatorios_frequencia"))

    if not frequencia_data:
        flash("Nenhum dado para gerar o relatório de frequência.", "aviso")
        return redirect(url_for("relatorios_frequencia"))

    # Obter nomes dos filtros para o título do relatório
    conn = None
    disciplina_nome = "Todas"
    turma_nome = "Todas"
    aluno_nome = "Todos"
    try:
        conn = conectar()
        cursor = conn.cursor()
        if disciplina_id:
            d = cursor.execute("SELECT nome FROM disciplinas WHERE id=?", (disciplina_id,)).fetchone()
            if d: disciplina_nome = d['nome']
        if turma_id:
            t = cursor.execute("SELECT nome FROM turmas WHERE id=?", (turma_id,)).fetchone()
            if t: turma_nome = t['nome']
        if aluno_id:
            a = cursor.execute("SELECT nome FROM alunos WHERE id=?", (aluno_id,)).fetchone()
            if a: aluno_nome = a['nome']
    except Exception as e:
        print(f"Erro ao buscar nomes para o relatório: {e}")
    finally:
        if conn:
            conn.close()

    titulo_relatorio = f"Relatório de Frequência - {disciplina_nome} / {turma_nome} / {aluno_nome}"
    if data_inicio and data_fim:
        titulo_relatorio += f" ({data_inicio} a {data_fim})"
    elif data_inicio:
        titulo_relatorio += f" (A partir de {data_inicio})"
    elif data_fim:
        titulo_relatorio += f" (Até {data_fim})"

    if format == "pdf":
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph(titulo_relatorio, styles['h2']))
        elements.append(Spacer(1, 0.2 * inch))

        data = [['Aluno', 'Disciplina', 'Turma', 'Presenças', 'Total Aulas', '% Frequência', 'Atividades Feitas', 'Frequência Mínima']]
        for item in frequencia_data:
            data.append([
                item['aluno_nome'],
                item['disciplina_nome'],
                item['turma_nome'],
                str(item['presencas']),
                str(item['total_aulas']),
                f"{item['frequencia_porcentagem']:.1f}%",
                str(item['atividades_feitas']) if item['tem_atividades'] else 'N/A',
                f"{item['frequencia_minima']:.1f}%"
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
        elements.append(table)

        doc.build(elements)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name=f"relatorio_frequencia_{date.today().strftime('%Y-%m-%d')}.pdf", mimetype="application/pdf")

    elif format == "docx":
        document = Document()
        document.add_heading(titulo_relatorio, level=1)

        table = document.add_table(rows=1, cols=8)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Aluno'
        hdr_cells[1].text = 'Disciplina'
        hdr_cells[2].text = 'Turma'
        hdr_cells[3].text = 'Presenças'
        hdr_cells[4].text = 'Total Aulas'
        hdr_cells[5].text = '% Frequência'
        hdr_cells[6].text = 'Atividades Feitas'
        hdr_cells[7].text = 'Frequência Mínima'

        for item in frequencia_data:
            row_cells = table.add_row().cells
            row_cells[0].text = item['aluno_nome']
            row_cells[1].text = item['disciplina_nome']
            row_cells[2].text = item['turma_nome']
            row_cells[3].text = str(item['presencas'])
            row_cells[4].text = str(item['total_aulas'])
            row_cells[5].text = f"{item['frequencia_porcentagem']:.1f}%"
            row_cells[6].text = str(item['atividades_feitas']) if item['tem_atividades'] else 'N/A'
            row_cells[7].text = f"{item['frequencia_minima']:.1f}%"

        buffer = BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name=f"relatorio_frequencia_{date.today().strftime('%Y-%m-%d')}.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        flash("Formato de download inválido.", "erro")
        return redirect(url_for("relatorios_frequencia"))


# ══════════════════════════════════════
# USUÁRIOS
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
@admin_required
def usuarios():
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT id, nome, email, perfil FROM usuarios ORDER BY nome")
        lista = cursor.fetchall()
        return render_template("usuarios.html", usuarios=lista)
    except Exception as e:
        flash(f"Erro ao carregar usuários: {e}", "erro")
        print(f"ERRO NOS USUÁRIOS: {e}")
        return render_template("usuarios.html", usuarios=[])
    finally:
        if conn:
            conn.close()


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
@admin_required
def novo_usuario():
    if request.method == "POST":
        nome  = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip()
        senha = request.form.get("senha", "")
        perfil = request.form.get("perfil", "").strip()

        if not nome or not email or not senha or not perfil:
            flash("Todos os campos são obrigatórios!", "erro")
            return render_template("novo_usuario.html", form_data=request.form)
        if len(senha) < 6:
            flash("A senha deve ter no mínimo 6 caracteres!", "erro")
            return render_template("novo_usuario.html", form_data=request.form)

        conn   = conectar()
        cursor = conn.cursor()
        try:
            senha_hash = generate_password_hash(senha)
            cursor.execute(
                "INSERT INTO usuarios (nome, email, senha_hash, perfil) VALUES (?, ?, ?, ?)",
                (nome, email, senha_hash, perfil))
            conn.commit()
            flash(f"Usuário '{nome}' criado!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "usuarios.email" in str(e):
                flash("Já existe um usuário com este e-mail!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar usuário: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar usuário: {e}", "erro")
            print(f"ERRO AO CADASTRAR USUÁRIO: {e}")
        finally:
            conn.close()
        return redirect(url_for("usuarios"))
    return render_template("novo_usuario.html", form_data={})


@app.route("/usuarios/<int:id>/editar", methods=["GET", "POST"])
@login_required
@admin_required
def editar_usuario(id):
    conn = None
    try:
        conn   = conectar()
        cursor = conn.cursor()
        if request.method == "POST":
            nome  = request.form.get("nome", "").strip()
            email = request.form.get("email", "").strip()
            perfil = request.form.get("perfil", "").strip()

            if not nome or not email or not perfil:
                flash("Nome, e-mail e perfil são obrigatórios!", "erro")
                usuario_data = request.form.to_dict()
                usuario_data['id'] = id
                return render_template("editar_usuario.html", usuario=usuario_data)

            try:
                cursor.execute(
                    "UPDATE usuarios SET nome=?, email=?, perfil=? WHERE id=?",
                    (nome, email, perfil, id))
                conn.commit()
                flash(f"Usuário '{nome}' atualizado!", "sucesso")
            except sqlite3.IntegrityError as e:
                if "usuarios.email" in str(e):
                    flash("Já existe um usuário com este e-mail!", "erro")
                else:
                    flash(f"Erro de integridade ao atualizar usuário: {e}", "erro")
            except Exception as e:
                flash(f"Erro inesperado ao atualizar usuário: {e}", "erro")
                print(f"ERRO AO ATUALIZAR USUÁRIO: {e}")
            return redirect(url_for("usuarios"))
        else:
            cursor.execute("SELECT id, nome, email, perfil FROM usuarios WHERE id=?", (id,))
            usuario = cursor.fetchone()
            if not usuario:
                flash("Usuário não encontrado!", "erro")
                return redirect(url_for("usuarios"))
            return render_template("editar_usuario.html", usuario=usuario)
    except Exception as e:
        flash(f"Erro ao carregar/editar usuário: {e}", "erro")
        print(f"ERRO EM EDITAR USUÁRIO: {e}")
        return redirect(url_for("usuarios"))
    finally:
        if conn:
            conn.close()


@app.route("/usuarios/<int:id>/excluir", methods=["POST"])
@login_required
@admin_required
def excluir_usuario(id):
    if id == current_user.id:
        flash("Você não pode excluir sua própria conta!", "erro")
        return redirect(url_for("usuarios"))

    conn   = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM usuarios WHERE id=?", (id,))
        conn.commit()
        flash("Usuário excluído!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir usuário: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("usuarios"))


@app.route("/minha-conta", methods=["GET", "POST"])
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
                    # Fechar a conexão atual com o banco de dados antes de substituir o arquivo
                    # Isso é crucial para evitar erros de arquivo em uso
                    # (A função conectar() cria uma nova conexão, então não precisamos fechar explicitamente aqui)

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