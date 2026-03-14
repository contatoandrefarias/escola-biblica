import os
from datetime import date
from flask import (Flask, render_template, request,
                   redirect, url_for, flash, send_file)
from flask_login import (LoginManager, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco
import sqlite3

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
login_manager.login_message = "Faca login para continuar."
login_manager.login_message_category = "warning"

@login_manager.user_loader
def load_user(user_id):
    return carregar_usuario(user_id)

inicializar_banco()


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
    flash(f"Ate logo, {nome}!", "sucesso")
    return redirect(url_for("login"))


# ══════════════════════════════════════
# PAINEL
# ══════════════════════════════════════
@app.route("/")
@login_required
def index():
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
    conn.close()
    return render_template("index.html",
        total_alunos=total_alunos,
        total_professores=total_professores,
        total_disciplinas=total_disciplinas,
        total_turmas=total_turmas,
        aprovados=aprovados,
        reprovados=reprovados,
        cursando=cursando)


# ══════════════════════════════════════
# TURMAS
# ══════════════════════════════════════
@app.route("/turmas")
@login_required
def turmas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT t.*, COUNT(a.id) as total_alunos
        FROM turmas t
        LEFT JOIN alunos a ON a.turma_id = t.id
        GROUP BY t.id ORDER BY t.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("turmas.html", turmas=lista)


@app.route("/turmas/nova", methods=["GET", "POST"])
@login_required
def nova_turma():
    if request.method == "POST":
        nome         = request.form.get("nome", "").strip()
        descricao    = request.form.get("descricao", "").strip()
        faixa_etaria = request.form.get("faixa_etaria", "").strip()
        if not nome:
            flash("Nome e obrigatorio!", "erro")
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
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome         = request.form.get("nome", "").strip()
        descricao    = request.form.get("descricao", "").strip()
        faixa_etaria = request.form.get("faixa_etaria", "").strip()
        ativa        = 1 if request.form.get("ativa") else 0
        try:
            cursor.execute("""
                UPDATE turmas
                SET nome=?,descricao=?,faixa_etaria=?,ativa=?
                WHERE id=?
            """, (nome, descricao, faixa_etaria, ativa, id))
            conn.commit()
            flash("Turma atualizada!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "turmas.nome" in str(e):
                flash("Já existe outra turma com este nome!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar turma: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao atualizar turma: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("turmas"))
    cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
    turma = cursor.fetchone()
    cursor.execute(
        "SELECT * FROM alunos WHERE turma_id=? ORDER BY nome", (id,))
    alunos_turma = cursor.fetchall()
    conn.close()
    if not turma:
        flash("Turma nao encontrada!", "erro")
        return redirect(url_for("turmas"))
    return render_template("editar_turma.html",
        turma=turma, alunos_turma=alunos_turma)


# ══════════════════════════════════════
# ALUNOS
# ══════════════════════════════════════
@app.route("/alunos")
@login_required
def alunos():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT a.*, t.nome as turma_nome
        FROM alunos a
        LEFT JOIN turmas t ON a.turma_id = t.id
        ORDER BY a.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("alunos.html", alunos=lista)


@app.route("/alunos/novo", methods=["GET", "POST"])
@login_required
def novo_aluno():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        telefone  = request.form.get("telefone", "").strip()
        email     = request.form.get("email", "").strip()
        data_nasc = request.form.get("data_nascimento", "").strip()
        membro    = 1 if request.form.get("membro_igreja") else 0
        turma_id  = request.form.get("turma_id") or None

        if not nome:
            flash("Nome e obrigatorio!", "erro")
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_lista)

        try:
            cursor.execute("""
                INSERT INTO alunos
                    (nome,telefone,email,data_nascimento,
                     membro_igreja,turma_id)
                VALUES (?,?,?,?,?,?)
            """, (nome, telefone, email, data_nasc, membro, turma_id))
            conn.commit()
            flash(f"Aluno '{nome}' cadastrado!", "sucesso")
            return redirect(url_for("alunos"))
        except sqlite3.IntegrityError as e:
            flash(f"Erro de integridade ao cadastrar aluno: {e}", "erro")
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_lista)
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar aluno: {e}", "erro")
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_lista)
        finally:
            conn.close()
    cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    conn.close()
    return render_template("novo_aluno.html", turmas=turmas_lista)


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        telefone  = request.form.get("telefone", "").strip()
        email     = request.form.get("email", "").strip()
        data_nasc = request.form.get("data_nascimento", "").strip()
        membro    = 1 if request.form.get("membro_igreja") else 0
        turma_id  = request.form.get("turma_id") or None

        if not nome:
            flash("Nome e obrigatorio!", "erro")
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=cursor.fetchone(), turmas=turmas_lista)

        try:
            cursor.execute("""
                UPDATE alunos
                SET nome=?,telefone=?,email=?,data_nascimento=?,
                    membro_igreja=?,turma_id=?
                WHERE id=?
            """, (nome, telefone, email, data_nasc, membro, turma_id, id))
            conn.commit()
            flash("Aluno atualizado!", "sucesso")
            return redirect(url_for("alunos"))
        except sqlite3.IntegrityError as e:
            flash(f"Erro de integridade ao atualizar aluno: {e}", "erro")
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=cursor.fetchone(), turmas=turmas_lista)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar aluno: {e}", "erro")
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=cursor.fetchone(), turmas=turmas_lista)
        finally:
            conn.close()
    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    conn.close()
    if not aluno:
        flash("Aluno nao encontrado!", "erro")
        return redirect(url_for("alunos"))
    return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_lista)


# ══════════════════════════════════════
# PROFESSORES
# ══════════════════════════════════════
@app.route("/professores")
@login_required
def professores():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    lista = cursor.fetchall()
    conn.close()
    return render_template("professores.html", professores=lista)


@app.route("/professores/novo", methods=["GET", "POST"])
@login_required
def novo_professor():
    if request.method == "POST":
        nome         = request.form.get("nome", "").strip()
        telefone     = request.form.get("telefone", "").strip()
        email        = request.form.get("email", "").strip()
        especialidade = request.form.get("especialidade", "").strip()
        if not nome:
            flash("Nome e obrigatorio!", "erro")
            return redirect(url_for("novo_professor"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO professores (nome,telefone,email,especialidade)
                VALUES (?,?,?,?)
            """, (nome, telefone, email, especialidade))
            conn.commit()
            flash(f"Professor '{nome}' cadastrado!", "sucesso")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar professor: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("professores"))
    return render_template("novo_professor.html")


@app.route("/professores/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_professor(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome         = request.form.get("nome", "").strip()
        telefone     = request.form.get("telefone", "").strip()
        email        = request.form.get("email", "").strip()
        especialidade = request.form.get("especialidade", "").strip()
        try:
            cursor.execute("""
                UPDATE professores
                SET nome=?,telefone=?,email=?,especialidade=?
                WHERE id=?
            """, (nome, telefone, email, especialidade, id))
            conn.commit()
            flash("Professor atualizado!", "sucesso")
        except Exception as e:
            flash(f"Erro inesperado ao atualizar professor: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("professores"))
    cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
    professor = cursor.fetchone()
    conn.close()
    if not professor:
        flash("Professor nao encontrado!", "erro")
        return redirect(url_for("professores"))
    return render_template("editar_professor.html", professor=professor)


# ══════════════════════════════════════
# DISCIPLINAS
# ══════════════════════════════════════
@app.route("/disciplinas")
@login_required
def disciplinas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT d.*, p.nome as professor_nome
        FROM disciplinas d
        LEFT JOIN professores p ON d.professor_id = p.id
        ORDER BY d.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("disciplinas.html", disciplinas=lista)


@app.route("/disciplinas/nova", methods=["GET", "POST"])
@login_required
def nova_disciplina():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome             = request.form.get("nome", "").strip()
        descricao        = request.form.get("descricao", "").strip()
        duracao_semanas  = request.form.get("duracao_semanas", type=int)
        nota_minima      = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades   = 1 if request.form.get("tem_atividades") else 0
        professor_id     = request.form.get("professor_id") or None

        if not nome or not duracao_semanas or nota_minima is None or frequencia_minima is None:
            flash("Nome, duracao, nota minima e frequencia minima sao obrigatorios!", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=professores_lista)

        try:
            cursor.execute("""
                INSERT INTO disciplinas
                    (nome,descricao,duracao_semanas,nota_minima,
                     frequencia_minima,tem_atividades,professor_id)
                VALUES (?,?,?,?,?,?,?)
            """, (nome, descricao, duracao_semanas, nota_minima,
                  frequencia_minima, tem_atividades, professor_id))
            conn.commit()
            flash(f"Disciplina '{nome}' cadastrada!", "sucesso")
            return redirect(url_for("disciplinas"))
        except sqlite3.IntegrityError as e:
            if "disciplinas.nome" in str(e):
                flash("Já existe uma disciplina com este nome!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar disciplina: {e}", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=professores_lista)
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar disciplina: {e}", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=professores_lista)
        finally:
            conn.close()
    cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
    professores_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_disciplina.html", professores=professores_lista)


@app.route("/disciplinas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome             = request.form.get("nome", "").strip()
        descricao        = request.form.get("descricao", "").strip()
        duracao_semanas  = request.form.get("duracao_semanas", type=int)
        nota_minima      = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades   = 1 if request.form.get("tem_atividades") else 0
        professor_id     = request.form.get("professor_id") or None
        ativa            = 1 if request.form.get("ativa") else 0

        if not nome or not duracao_semanas or nota_minima is None or frequencia_minima is None:
            flash("Nome, duracao, nota minima e frequencia minima sao obrigatorios!", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=cursor.fetchone(), professores=professores_lista)

        try:
            cursor.execute("""
                UPDATE disciplinas
                SET nome=?,descricao=?,duracao_semanas=?,nota_minima=?,
                    frequencia_minima=?,tem_atividades=?,professor_id=?,ativa=?
                WHERE id=?
            """, (nome, descricao, duracao_semanas, nota_minima,
                  frequencia_minima, tem_atividades, professor_id, ativa, id))
            conn.commit()
            flash("Disciplina atualizada!", "sucesso")
            return redirect(url_for("disciplinas"))
        except sqlite3.IntegrityError as e:
            if "disciplinas.nome" in str(e):
                flash("Já existe outra disciplina com este nome!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar disciplina: {e}", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=cursor.fetchone(), professores=professores_lista)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar disciplina: {e}", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=cursor.fetchone(), professores=professores_lista)
        finally:
            conn.close()
    cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
    disciplina = cursor.fetchone()
    cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
    professores_lista = cursor.fetchall()
    conn.close()
    if not disciplina:
        flash("Disciplina nao encontrada!", "erro")
        return redirect(url_for("disciplinas"))
    return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores_lista)


# ══════════════════════════════════════
# MATRICULAS
# ══════════════════════════════════════
@app.route("/matriculas")
@login_required
def matriculas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            m.id,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.nome as turma_nome,
            t.faixa_etaria,
            m.data_inicio,
            m.data_conclusao,
            m.nota_final,
            m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        ORDER BY a.nome, d.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("matriculas.html", matriculas=lista)


@app.route("/matriculas/nova", methods=["GET", "POST"])
@login_required
def nova_matricula():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        aluno_id      = request.form.get("aluno_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio")

        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Aluno, disciplina e data de inicio sao obrigatorios!", "erro")
            # Recarrega listas para o template
            cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome") # <<-- AQUI
            turmas_lista = cursor.fetchall()
            cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   alunos=alunos_lista,
                                   turmas=turmas_lista, # <<-- AQUI
                                   disciplinas=disciplinas_lista)

        # Buscar todos os alunos da turma selecionada
        cursor.execute("SELECT id FROM alunos WHERE turma_id = ?", (aluno_id,)) # aluno_id aqui é na verdade turma_id
        alunos_da_turma = cursor.fetchall()

        if not alunos_da_turma:
            flash("Nenhum aluno encontrado nesta turma para matricular.", "erro")
            cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome") # <<-- AQUI
            turmas_lista = cursor.fetchall()
            cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   alunos=alunos_lista,
                                   turmas=turmas_lista, # <<-- AQUI
                                   disciplinas=disciplinas_lista)

        matriculas_inseridas = 0
        for aluno_turma in alunos_da_turma:
            try:
                cursor.execute("""
                    INSERT INTO matriculas
                        (aluno_id,disciplina_id,data_inicio,status)
                    VALUES (?,?,?,?)
                """, (aluno_turma['id'], disciplina_id, data_inicio, 'cursando'))
                matriculas_inseridas += 1
            except sqlite3.IntegrityError:
                # Ignora se a matrícula já existe para este aluno nesta disciplina
                pass
            except Exception as e:
                flash(f"Erro ao matricular aluno {aluno_turma['id']}: {e}", "erro")
                conn.rollback() # Desfaz as operações se houver um erro grave
                conn.close()
                return redirect(url_for("nova_matricula"))

        conn.commit()
        conn.close()
        flash(f"{matriculas_inseridas} matrículas criadas ou atualizadas para a turma!", "sucesso")
        return redirect(url_for("matriculas"))

    # GET request
    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos_lista = cursor.fetchall()
    # MODIFICADO: Buscar faixa_etaria para turmas
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html",
                           alunos=alunos_lista,
                           turmas=turmas_lista,
                           disciplinas=disciplinas_lista)


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        data_inicio   = request.form.get("data_inicio")
        data_conclusao = request.form.get("data_conclusao") or None # Pode ser NULL

        # Notas para Adultos
        nota1_adulto = request.form.get("nota1", type=float)
        nota2_adulto = request.form.get("nota2", type=float)

        # Notas para Adolescentes/Jovens
        participacao_aj = request.form.get("participacao", type=float)
        desafio_aj      = request.form.get("desafio", type=float)
        prova_aj        = request.form.get("prova", type=float)

        # Buscar dados da matrícula e da turma para a lógica de notas
        cursor.execute("""
            SELECT
                m.aluno_id,
                m.disciplina_id,
                t.faixa_etaria,
                d.nota_minima,
                d.frequencia_minima
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.id = ?
        """, (id,))
        matricula_info = cursor.fetchone()

        if not matricula_info:
            flash("Matrícula não encontrada!", "erro")
            conn.close()
            return redirect(url_for("matriculas"))

        faixa_etaria = matricula_info['faixa_etaria']
        nota_minima = matricula_info['nota_minima']
        frequencia_minima = matricula_info['frequencia_minima']

        # Chamar a função de atualização de status e notas
        _atualizar_status_matricula(
            conn, id, faixa_etaria, nota_minima, frequencia_minima,
            data_inicio, data_conclusao,
            nota1_adulto, nota2_adulto,
            participacao_aj, desafio_aj, prova_aj
        )

        conn.commit()
        conn.close()
        flash("Matrícula atualizada!", "sucesso")
        return redirect(url_for("matriculas"))

    # GET request
    cursor.execute("""
        SELECT
            m.id,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.faixa_etaria,
            m.data_inicio,
            m.data_conclusao,
            m.nota1, m.nota2, m.nota_final,
            m.participacao, m.desafio, m.prova,
            m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE m.id = ?
    """, (id,))
    matricula = cursor.fetchone()
    conn.close()
    if not matricula:
        flash("Matrícula não encontrada!", "erro")
        return redirect(url_for("matriculas"))
    return render_template("editar_matricula.html", matricula=matricula)


# Função auxiliar para atualizar status e notas
def _atualizar_status_matricula(
    conn, matricula_id, faixa_etaria, nota_minima, frequencia_minima,
    data_inicio, data_conclusao,
    nota1_adulto, nota2_adulto,
    participacao_aj, desafio_aj, prova_aj
):
    cursor = conn.cursor()

    new_nota1 = None
    new_nota2 = None
    new_nota_final = None
    new_participacao = None
    new_desafio = None
    new_prova = None

    if faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
        new_participacao = participacao_aj
        new_desafio = desafio_aj
        new_prova = prova_aj

        # Calcula N1 para adolescentes/jovens
        calc_nota1 = 0.0
        if new_participacao is not None:
            calc_nota1 += new_participacao
        if new_desafio is not None:
            calc_nota1 += new_desafio
        if new_prova is not None:
            calc_nota1 += new_prova

        new_nota1 = calc_nota1
        new_nota_final = new_nota1 # Para adolescentes/jovens, N1 é a nota final
        new_nota2 = None # Garante que N2 é nulo para adolescentes/jovens
    elif faixa_etaria == 'adultos':
        new_nota1 = nota1_adulto
        new_nota2 = nota2_adulto

        # Calcula nota final para adultos
        if new_nota1 is not None and new_nota2 is not None:
            new_nota_final = (new_nota1 + new_nota2) / 2
        elif new_nota1 is not None: # Se apenas N1 foi preenchido, usa N1 como final
            new_nota_final = new_nota1
        elif new_nota2 is not None: # Se apenas N2 foi preenchido, usa N2 como final
            new_nota_final = new_nota2
        else:
            new_nota_final = None
        new_participacao = None # Garante que são nulos para adultos
        new_desafio = None
        new_prova = None
    else: # Criancas (criancas_0_3, criancas_4_7, criancas_8_12)
        new_nota1 = None
        new_nota2 = None
        new_nota_final = None
        new_participacao = None
        new_desafio = None
        new_prova = None

    # Calcular frequência
    cursor.execute("""
        SELECT
            COUNT(p.id) as total_aulas,
            SUM(p.presente) as presencas
        FROM presencas p
        WHERE p.matricula_id = ?
    """, (matricula_id,))
    frequencia_data = cursor.fetchone()
    total_aulas = frequencia_data['total_aulas'] or 0
    presencas = frequencia_data['presencas'] or 0

    frequencia_percentual = 0.0
    if total_aulas > 0:
        frequencia_percentual = (presencas / total_aulas) * 100

    # Lógica para determinar o status
    new_status = 'cursando'
    if data_conclusao: # Se a disciplina foi concluída
        if faixa_etaria and faixa_etaria.startswith('criancas'):
            if frequencia_percentual >= frequencia_minima:
                new_status = 'aprovado'
            else:
                new_status = 'reprovado'
        elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
            if new_nota_final is not None and new_nota_final >= nota_minima and \
               frequencia_percentual >= frequencia_minima:
                new_status = 'aprovado'
            else:
                new_status = 'reprovado'
        elif faixa_etaria == 'adultos':
            if new_nota_final is not None and new_nota_final >= nota_minima and \
               frequencia_percentual >= frequencia_minima:
                new_status = 'aprovado'
            else:
                new_status = 'reprovado'

    # Atualizar a matrícula no banco de dados
    cursor.execute("""
        UPDATE matriculas
        SET nota1 = ?, nota2 = ?, nota_final = ?, status = ?,
            participacao = ?, desafio = ?, prova = ?
        WHERE id = ?
    """, (new_nota1, new_nota2, new_nota_final, new_status,
          new_participacao, new_desafio, new_prova, matricula_id))

# ══════════════════════════════════════
# RELATORIOS
# ══════════════════════════════════════
@app.route("/relatorios")
@login_required
def relatorios():
    conn   = conectar()
    cursor = conn.cursor()

    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro", "todos")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro, conn)

    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()

    return render_template("relatorios.html",
                           dados=dados,
                           disciplinas=disciplinas_lista,
                           selected_disciplina=disciplina_id,
                           selected_data_inicio=data_inicio,
                           selected_data_fim=data_fim,
                           selected_status=status_filtro)


def get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro, conn):
    cursor = conn.cursor()
    query = """
        SELECT
            a.nome  as aluno,
            d.nome  as disciplina,
            d.nota_minima,
            d.frequencia_minima,
            t.faixa_etaria,
            m.nota1, m.nota2, m.nota_final,
            m.participacao, m.desafio, m.prova,
            m.status,
            m.data_inicio,
            m.data_conclusao,
            (SELECT COUNT(p_sub.id) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as total_aulas,
            (SELECT SUM(p_sub.presente) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as presencas,
            (SELECT SUM(p_sub.fez_atividade) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as atividades
        FROM matriculas m
        JOIN alunos      a ON m.aluno_id      = a.id
        LEFT JOIN turmas t ON a.turma_id      = t.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE 1=1
    """
    params = []

    if disciplina_id:
        query += " AND d.id = ?"
        params.append(disciplina_id)
    if data_inicio:
        query += " AND m.data_inicio >= ?"
        params.append(data_inicio)
    if data_fim:
        query += " AND m.data_conclusao <= ?"
        params.append(data_fim)
    if status_filtro and status_filtro != "todos":
        query += " AND m.status = ?"
        params.append(status_filtro)

    query += " ORDER BY a.nome, d.nome"
    cursor.execute(query, tuple(params))
    return cursor.fetchall()


@app.route("/relatorios/download/pdf")
@login_required
def download_relatorio_pdf():
    conn = conectar()
    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro, conn)
    conn.close()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    elements = []
    styles = getSampleStyleSheet()

    # Título
    elements.append(Paragraph("Relatório de Matrículas", styles['h1']))
    elements.append(Spacer(1, 0.2 * inch))

    # Filtros aplicados
    filter_text = "Filtros: "
    if disciplina_id:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT nome FROM disciplinas WHERE id = ?", (disciplina_id,))
        disc_nome = cursor.fetchone()['nome']
        conn.close()
        filter_text += f"Disciplina: {disc_nome}; "
    if data_inicio:
        filter_text += f"Início: {data_inicio}; "
    if data_fim:
        filter_text += f"Fim: {data_fim}; "
    if status_filtro and status_filtro != 'todos':
        filter_text += f"Status: {status_filtro.capitalize()}; "
    if filter_text == "Filtros: ":
        filter_text += "Nenhum"
    elements.append(Paragraph(filter_text, styles['Normal']))
    elements.append(Spacer(1, 0.2 * inch))

    # Dados da tabela
    if dados:
        table_data = []
        # Cabeçalho
        table_data.append([
            "Aluno", "Disciplina", "Início", "Conclusão",
            "Média", "Status", "Frequência", "Atividades"
        ])

        for item in dados:
            freq_val = "—"
            if item['total_aulas'] is not None and item['total_aulas'] > 0:
                freq = ((item['presencas'] or 0) / item['total_aulas'] * 100)
                freq_val = f"{freq:.1f}% ({item['presencas'] or 0}/{item['total_aulas']})"

            media_val = "—"
            if item['faixa_etaria'] and item['faixa_etaria'].startswith('criancas'):
                media_val = "N/A"
            elif item['faixa_etaria'] and (item['faixa_etaria'].startswith('adolescentes') or item['faixa_etaria'].startswith('jovens')):
                if item['nota1'] is not None:
                    media_val = f"{item['nota1']:.1f}"
            else: # Adultos
                if item['nota_final'] is not None:
                    media_val = f"{item['nota_final']:.1f}"

            status_val = item['status'].capitalize() if item['status'] else "—"

            table_data.append([
                item['aluno'],
                item['disciplina'],
                item['data_inicio'] or "—",
                item['data_conclusao'] or "Em andamento",
                media_val,
                status_val,
                freq_val,
                item['atividades'] or "—"
            ])

        table = Table(table_data, colWidths=[1.5*inch, 1.5*inch, 0.8*inch, 1.0*inch, 0.8*inch, 0.8*inch, 1.2*inch, 0.8*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#dee2e6')),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        elements.append(table)
    else:
        elements.append(Paragraph("Nenhum relatório encontrado com os filtros aplicados.", styles['Normal']))

    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.pdf", mimetype="application/pdf")


@app.route("/relatorios/download/doc")
@login_required
def download_relatorio_doc():
    conn = conectar()
    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro, conn)
    conn.close()

    document = Document()
    document.settings.element.xpath('//w:settings')[0].append(
        OxmlElement('w:lang', {'val': 'pt-BR'})
    )

    # Título
    document.add_heading('Relatório de Matrículas', level=1)

    # Filtros aplicados
    filter_text = "Filtros: "
    if disciplina_id:
        conn = conectar()
        cursor = conn.cursor()
        cursor.execute("SELECT nome FROM disciplinas WHERE id = ?", (disciplina_id,))
        disc_nome = cursor.fetchone()['nome']
        conn.close()
        filter_text += f"Disciplina: {disc_nome}; "
    if data_inicio:
        filter_text += f"Início: {data_inicio}; "
    if data_fim:
        filter_text += f"Fim: {data_fim}; "
    if status_filtro and status_filtro != 'todos':
        filter_text += f"Status: {status_filtro.capitalize()}; "
    if filter_text == "Filtros: ":
        filter_text += "Nenhum"
    document.add_paragraph(filter_text)
    document.add_paragraph() # Espaço

    # Dados da tabela
    if dados:
        # Cabeçalho
        headers = [
            "Aluno", "Disciplina", "Início", "Conclusão",
            "Média", "Status", "Frequência", "Atividades"
        ]
        table = document.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            p = hdr_cells[i].paragraphs[0]
            p.text = header
            p.runs[0].bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hdr_cells[i].width = Inches(1.0)

        # Linhas de dados
        for item in dados:
            row_cells = table.add_row().cells

            freq_val = "—"
            if item['total_aulas'] is not None and item['total_aulas'] > 0:
                freq = ((item['presencas'] or 0) / item['total_aulas'] * 100)
                freq_val = f"{freq:.1f}% ({item['presencas'] or 0}/{item['total_aulas']})"

            media_val = "—"
            if item['faixa_etaria'] and item['faixa_etaria'].startswith('criancas'):
                media_val = "N/A"
            elif item['faixa_etaria'] and (item['faixa_etaria'].startswith('adolescentes') or item['faixa_etaria'].startswith('jovens')):
                if item['nota1'] is not None:
                    media_val = f"{item['nota1']:.1f}"
            else: # Adultos
                if item['nota_final'] is not None:
                    media_val = f"{item['nota_final']:.1f}"

            status_val = item['status'].capitalize() if item['status'] else "—"

            row_cells[0].text = item['aluno']
            row_cells[1].text = item['disciplina']
            row_cells[2].text = item['data_inicio'] or "—"
            row_cells[3].text = item['data_conclusao'] or "Em andamento"
            row_cells[4].text = media_val
            row_cells[5].text = status_val
            row_cells[6].text = freq_val
            row_cells[7].text = item['atividades'] or "—"

            # Centralizar todas as células, exceto a primeira (Aluno)
            for i in range(1, len(headers)):
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    else:
        document.add_paragraph("Nenhum relatório encontrado com os filtros aplicados.")

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ══════════════════════════════════════
# USUARIOS
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
def usuarios():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM usuarios ORDER BY nome")
    lista = cursor.fetchall()
    conn.close()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
def novo_usuario():
    if request.method == "POST":
        nome   = request.form.get("nome", "").strip()
        email  = request.form.get("email", "").strip()
        senha  = request.form.get("senha", "")
        perfil = request.form.get("perfil", "usuario")
        if not nome or not email or not senha:
            flash("Todos os campos sao obrigatorios!", "erro")
            return redirect(url_for("novo_usuario"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO usuarios (nome,email,senha_hash,perfil)
                VALUES (?,?,?,?)
            """, (nome, email,
                  generate_password_hash(senha), perfil))
            conn.commit()
            flash(f"Usuario '{nome}' criado!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "usuarios.email" in str(e):
                flash("Este e-mail já está cadastrado para outro usuário!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar usuário: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar usuário: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("usuarios"))
    return render_template("novo_usuario.html")


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
            flash("As senhas nao coincidem!", "erro")
            return redirect(url_for("minha_conta"))
        if len(nova_senha) < 6:
            flash("Minimo 6 caracteres!", "erro")
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)