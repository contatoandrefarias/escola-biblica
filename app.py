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
            # Obter turmas novamente para renderizar o template
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
            # Capturar outros erros de integridade para alunos
            flash(f"Erro de integridade ao cadastrar aluno: {e}", "erro")
            # Obter turmas novamente para renderizar o template
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_lista)
        except Exception as e:
            # Capturar outros erros genéricos
            flash(f"Erro inesperado ao cadastrar aluno: {e}", "erro")
            # Obter turmas novamente para renderizar o template
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_lista)
        finally:
            conn.close() # Garantir que a conexão seja fechada

    cursor.execute(
        "SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    conn.close()
    return render_template("novo_aluno.html", turmas=turmas_lista)


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome     = request.form.get("nome", "").strip()
        telefone = request.form.get("telefone", "").strip()
        email    = request.form.get("email", "").strip()
        membro   = 1 if request.form.get("membro_igreja") else 0
        turma_id = request.form.get("turma_id") or None

        if not nome:
            flash("Nome e obrigatorio!", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_lista)

        try:
            cursor.execute("""
                UPDATE alunos
                SET nome=?,telefone=?,email=?,
                    membro_igreja=?,turma_id=?
                WHERE id=?
            """, (nome, telefone, email, membro, turma_id, id))
            conn.commit()
            flash("Aluno atualizado!", "sucesso")
            return redirect(url_for("alunos"))
        except sqlite3.IntegrityError as e:
            # Capturar outros erros de integridade para alunos
            flash(f"Erro de integridade ao atualizar aluno: {e}", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_lista)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar aluno: {e}", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_lista)
        finally:
            conn.close()

    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    cursor.execute(
        "SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    conn.close()
    if not aluno:
        flash("Aluno nao encontrado!", "erro")
        return redirect(url_for("alunos"))
    return render_template("editar_aluno.html",
        aluno=aluno, turmas=turmas_lista)


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se o aluno tem matrículas
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE aluno_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir o aluno, pois ele possui matrículas ativas.", "erro")
            return redirect(url_for("alunos"))

        cursor.execute("DELETE FROM alunos WHERE id=?", (id,))
        conn.commit()
        flash("Aluno excluído!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir aluno: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("alunos"))


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
        except sqlite3.IntegrityError as e:
            if "professores.email" in str(e):
                flash("Este e-mail já está cadastrado para outro professor!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar professor: {e}", "erro")
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
        except sqlite3.IntegrityError as e:
            if "professores.email" in str(e):
                flash("Este e-mail já está cadastrado para outro professor!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar professor: {e}", "erro")
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


@app.route("/professores/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_professor(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se o professor está associado a alguma disciplina
        cursor.execute("SELECT COUNT(*) FROM disciplinas WHERE professor_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir o professor, pois ele está associado a uma ou mais disciplinas.", "erro")
            return redirect(url_for("professores"))

        cursor.execute("DELETE FROM professores WHERE id=?", (id,))
        conn.commit()
        flash("Professor excluído!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir professor: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("professores"))


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
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        duracao_semanas = request.form.get("duracao_semanas", type=int) or 4
        nota_minima     = request.form.get("nota_minima", type=float) or 6.0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        professor_id    = request.form.get("professor_id") or None
        ativa           = 1 if request.form.get("ativa") else 0

        if not nome:
            flash("Nome e obrigatorio!", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=professores_lista)

        try:
            cursor.execute("""
                INSERT INTO disciplinas
                    (nome,descricao,duracao_semanas,nota_minima,
                     frequencia_minima,tem_atividades,professor_id,ativa)
                VALUES (?,?,?,?,?,?,?,?)
            """, (nome, descricao, duracao_semanas, nota_minima,
                  frequencia_minima, tem_atividades, professor_id, ativa))
            conn.commit()
            flash(f"Disciplina '{nome}' criada!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "disciplinas.nome" in str(e):
                flash("Já existe uma disciplina com este nome!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar disciplina: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar disciplina: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("disciplinas"))

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
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        duracao_semanas = request.form.get("duracao_semanas", type=int) or 4
        nota_minima     = request.form.get("nota_minima", type=float) or 6.0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        professor_id    = request.form.get("professor_id") or None
        ativa           = 1 if request.form.get("ativa") else 0

        if not nome:
            flash("Nome e obrigatorio!", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores_lista)

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
        except sqlite3.IntegrityError as e:
            if "disciplinas.nome" in str(e):
                flash("Já existe outra disciplina com este nome!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar disciplina: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao atualizar disciplina: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("disciplinas"))

    cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
    disciplina = cursor.fetchone()
    cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
    professores_lista = cursor.fetchall()
    conn.close()
    if not disciplina:
        flash("Disciplina nao encontrada!", "erro")
        return redirect(url_for("disciplinas"))
    return render_template("editar_disciplina.html",
        disciplina=disciplina, professores=professores_lista)


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se a disciplina tem matrículas
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir a disciplina, pois ela possui matrículas ativas.", "erro")
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
# MATRICULAS
# ══════════════════════════════════════
@app.route("/matriculas")
@login_required
def matriculas():
    conn = conectar()
    cursor = conn.cursor()

    # Obter parâmetros de filtro da URL
    turma_id_filtro = request.args.get("turma_id", type=int)
    disciplina_id_filtro = request.args.get("disciplina_id", type=int)

    query = """
        SELECT
            m.id as matricula_id,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.nome as turma_nome,
            m.data_inicio,
            m.data_conclusao,
            m.nota1, m.nota2, m.nota_final,
            m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE 1=1
    """
    params = []

    if turma_id_filtro:
        query += " AND a.turma_id = ?"
        params.append(turma_id_filtro)

    if disciplina_id_filtro:
        query += " AND d.id = ?"
        params.append(disciplina_id_filtro)

    query += " ORDER BY t.nome, a.nome, d.nome"

    cursor.execute(query, tuple(params))
    lista = cursor.fetchall()

    # Obter turmas e disciplinas para os filtros
    cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()

    conn.close()
    return render_template("matriculas.html",
        matriculas=lista,
        turmas=turmas_lista,
        disciplinas=disciplinas_lista,
        selected_turma=turma_id_filtro,
        selected_disciplina=disciplina_id_filtro)


@app.route("/matriculas/nova", methods=["GET", "POST"])
@login_required
def nova_matricula():
    conn = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        turma_id      = request.form.get("turma_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio")
        data_conclusao = request.form.get("data_conclusao")

        if not turma_id or not disciplina_id or not data_inicio:
            flash("Selecione uma turma, uma disciplina e a data de início.", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   turmas=turmas_lista,
                                   disciplinas=disciplinas_lista,
                                   selected_turma=turma_id,
                                   selected_disciplina=disciplina_id,
                                   data_inicio=data_inicio,
                                   data_conclusao=data_conclusao)

        # Obter todos os alunos da turma selecionada
        cursor.execute("SELECT id, nome FROM alunos WHERE turma_id = ? ORDER BY nome", (turma_id,))
        alunos_da_turma = cursor.fetchall()

        if not alunos_da_turma:
            flash("Não há alunos nesta turma para matricular.", "aviso")
            # Recarregar dados para o template
            cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   turmas=turmas_lista,
                                   disciplinas=disciplinas_lista,
                                   selected_turma=turma_id,
                                   selected_disciplina=disciplina_id,
                                   data_inicio=data_inicio,
                                   data_conclusao=data_conclusao)

        matriculas_criadas = 0
        matriculas_existentes = 0
        erros = []

        for aluno in alunos_da_turma:
            aluno_id = aluno['id']
            try:
                cursor.execute("""
                    INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, data_conclusao, status)
                    VALUES (?, ?, ?, ?, 'cursando')
                """, (aluno_id, disciplina_id, data_inicio, data_conclusao))
                matriculas_criadas += 1
            except sqlite3.IntegrityError as e:
                if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                    matriculas_existentes += 1
                else:
                    erros.append(f"Erro de integridade para {aluno['nome']}: {e}")
            except Exception as e:
                erros.append(f"Erro inesperado para {aluno['nome']}: {e}")

        conn.commit()

        if matriculas_criadas > 0:
            flash(f"{matriculas_criadas} matrículas criadas com sucesso para a turma!", "sucesso")
        if matriculas_existentes > 0:
            flash(f"{matriculas_existentes} alunos já estavam matriculados nesta disciplina.", "aviso")
        if erros:
            for erro in erros:
                flash(erro, "erro")

        conn.close()
        return redirect(url_for("matriculas"))

    # Método GET
    cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html",
                           turmas=turmas_lista,
                           disciplinas=disciplinas_lista,
                           hoje=date.today().isoformat())


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        nota1_str        = request.form.get("nota1", "").strip()
        nota2_str        = request.form.get("nota2", "").strip()
        nota_final_str   = request.form.get("nota_final", "").strip()
        data_inicio      = request.form.get("data_inicio")
        data_conclusao   = request.form.get("data_conclusao")

        # Converter notas para float, ou None se vazio
        nota1 = float(nota1_str) if nota1_str else None
        nota2 = float(nota2_str) if nota2_str else None
        nota_final = float(nota_final_str) if nota_final_str else None

        # Obter dados da matrícula e disciplina para cálculo
        cursor.execute("""
            SELECT
                m.aluno_id, m.disciplina_id,
                d.nota_minima, d.frequencia_minima
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.id = ?
        """, (id,))
        matricula_info = cursor.fetchone()

        if not matricula_info:
            flash("Matrícula não encontrada.", "erro")
            conn.close()
            return redirect(url_for("matriculas"))

        disciplina_id = matricula_info['disciplina_id']
        aluno_id = matricula_info['aluno_id']
        nota_minima = matricula_info['nota_minima']
        frequencia_minima = matricula_info['frequencia_minima']

        # Calcular nota_final se N1 e N2 estiverem preenchidas e nota_final não foi sobrescrita
        if nota1 is not None and nota2 is not None and not nota_final_str:
            nota_final = (nota1 + nota2) / 2

        # Calcular frequência
        cursor.execute("""
            SELECT
                COUNT(p.id) as total_aulas,
                SUM(p.presente) as presencas
            FROM presencas p
            JOIN matriculas m ON p.matricula_id = m.id
            WHERE m.aluno_id = ? AND m.disciplina_id = ?
        """, (aluno_id, disciplina_id))
        frequencia_data = cursor.fetchone()
        total_aulas = frequencia_data['total_aulas'] or 0
        presencas = frequencia_data['presencas'] or 0

        frequencia_percentual = 0.0
        if total_aulas > 0:
            frequencia_percentual = (presencas / total_aulas) * 100

        # Determinar status
        status = "cursando"
        if nota_final is not None and frequencia_percentual is not None:
            if nota_final >= nota_minima and frequencia_percentual >= frequencia_minima:
                status = "aprovado"
            elif nota_final < nota_minima or frequencia_percentual < frequencia_minima:
                status = "reprovado"

        try:
            cursor.execute("""
                UPDATE matriculas
                SET nota1=?, nota2=?, nota_final=?, status=?,
                    data_inicio=?, data_conclusao=?
                WHERE id=?
            """, (nota1, nota2, nota_final, status,
                  data_inicio, data_conclusao, id))
            conn.commit()
            flash("Matrícula atualizada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    # Método GET
    cursor.execute("""
        SELECT
            m.id, m.aluno_id, m.disciplina_id,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            d.nota_minima, d.frequencia_minima,
            m.data_inicio, m.data_conclusao,
            m.nota1, m.nota2, m.nota_final, m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE m.id = ?
    """, (id,))
    matricula = cursor.fetchone()

    if not matricula:
        flash("Matrícula não encontrada!", "erro")
        conn.close()
        return redirect(url_for("matriculas"))

    # Calcular frequência para exibição
    cursor.execute("""
        SELECT
            COUNT(p.id) as total_aulas,
            SUM(p.presente) as presencas
        FROM presencas p
        JOIN matriculas m ON p.matricula_id = m.id
        WHERE m.aluno_id = ? AND m.disciplina_id = ?
    """, (matricula['aluno_id'], matricula['disciplina_id']))
    frequencia_data = cursor.fetchone()
    total_aulas = frequencia_data['total_aulas'] or 0
    presencas = frequencia_data['presencas'] or 0

    frequencia_percentual = 0.0
    if total_aulas > 0:
        frequencia_percentual = (presencas / total_aulas) * 100

    conn.close()
    return render_template("editar_matricula.html",
        matricula=matricula,
        total_aulas=total_aulas,
        presencas=presencas,
        frequencia_percentual=frequencia_percentual)


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
        flash("Matrícula excluída!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# PRESENCA
# ══════════════════════════════════════
@app.route("/presenca")
@login_required
def presenca():
    conn = conectar()
    cursor = conn.cursor()

    disciplinas_cursando = []
    disciplinas_concluidas = []

    if current_user.is_aluno:
        aluno_id = current_user.aluno_id
        if aluno_id:
            cursor.execute("""
                SELECT d.id, d.nome, m.status
                FROM disciplinas d
                JOIN matriculas m ON d.id = m.disciplina_id
                WHERE m.aluno_id = ? AND d.ativa = 1
                ORDER BY d.nome
            """, (aluno_id,))
            disciplinas_aluno = cursor.fetchall()
            for disc in disciplinas_aluno:
                if disc['status'] == 'cursando':
                    disciplinas_cursando.append(disc)
                else:
                    disciplinas_concluidas.append(disc)
        else:
            flash("Seu perfil de aluno não está associado a um registro de aluno.", "erro")
    else: # Admin ou Professor veem todas as disciplinas ativas
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
        disciplinas_cursando = cursor.fetchall() # Todas são tratadas como "cursando" para seleção

    conn.close()
    return render_template("presenca.html",
                           disciplinas_cursando=disciplinas_cursando,
                           disciplinas_concluidas=disciplinas_concluidas,
                           hoje=date.today().isoformat())


@app.route("/presenca/chamada", methods=["GET"])
@login_required
def chamada():
    disciplina_id = request.args.get("disciplina_id", type=int)
    data_aula     = request.args.get("data_aula")

    if not disciplina_id or not data_aula:
        flash("Selecione uma disciplina e uma data para ver a chamada.", "erro")
        return redirect(url_for("presenca"))

    conn = conectar()
    cursor = conn.cursor()

    # Obter informações da disciplina
    cursor.execute("SELECT nome, tem_atividades FROM disciplinas WHERE id = ?", (disciplina_id,))
    disciplina = cursor.fetchone()
    if not disciplina:
        flash("Disciplina não encontrada.", "erro")
        conn.close()
        return redirect(url_for("presenca"))

    # Obter alunos matriculados na disciplina e suas presenças para a data
    cursor.execute("""
        SELECT
            a.id as aluno_id,
            a.nome as aluno_nome,
            m.id as matricula_id,
            p.presente,
            p.fez_atividade
        FROM alunos a
        JOIN matriculas m ON a.id = m.aluno_id
        LEFT JOIN presencas p ON m.id = p.matricula_id AND p.data_aula = ?
        WHERE m.disciplina_id = ?
        ORDER BY a.nome
    """, (data_aula, disciplina_id))
    alunos = cursor.fetchall()

    conn.close()
    return render_template("chamada.html",
                           disciplina=disciplina,
                           data_aula=data_aula,
                           alunos=alunos,
                           tem_atividades=disciplina['tem_atividades'])


@app.route("/presenca/salvar", methods=["POST"])
@login_required
def salvar_chamada():
    conn = conectar()
    cursor = conn.cursor()
    disciplina_id = request.form.get("disciplina_id", type=int)
    data_aula     = request.form.get("data_aula")

    if not disciplina_id or not data_aula:
        flash("Dados incompletos para salvar a chamada.", "erro")
        conn.close()
        return redirect(url_for("presenca"))

    try:
        # Obter se a disciplina tem atividades para saber se deve salvar 'fez_atividade'
        disciplina_info = cursor.execute("SELECT tem_atividades FROM disciplinas WHERE id = ?", (disciplina_id,)).fetchone()
        tem_atividades = disciplina_info['tem_atividades'] if disciplina_info else 0

        for key, value in request.form.items():
            if key.startswith("presente_"):
                matricula_id = int(key.replace("presente_", ""))
                presente = 1 if value == "on" else 0
                fez_atividade = 1 if tem_atividades and request.form.get(f"atividade_{matricula_id}") == "on" else 0

                # Verificar se já existe um registro de presença para esta matrícula e data
                cursor.execute("""
                    SELECT id FROM presencas
                    WHERE matricula_id = ? AND data_aula = ?
                """, (matricula_id, data_aula))
                presenca_id = cursor.fetchone()

                if presenca_id:
                    # Atualizar
                    cursor.execute("""
                        UPDATE presencas SET presente=?, fez_atividade=?
                        WHERE id=?
                    """, (presente, fez_atividade, presenca_id['id']))
                else:
                    # Inserir
                    cursor.execute("""
                        INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                        VALUES (?, ?, ?, ?)
                    """, (matricula_id, data_aula, presente, fez_atividade))

        conn.commit()
        flash("Chamada salva com sucesso!", "sucesso")
    except sqlite3.IntegrityError as e:
        flash(f"Erro de integridade ao salvar chamada: {e}", "erro")
    except Exception as e:
        flash(f"Erro inesperado ao salvar chamada: {e}", "erro")
    finally:
        conn.close()

    return redirect(url_for("chamada", disciplina_id=disciplina_id, data_aula=data_aula))


# ══════════════════════════════════════
# RELATORIOS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro):
    """Função auxiliar para obter os dados do relatório com base nos filtros."""
    conn = conectar()
    cursor = conn.cursor()

    query = """
        SELECT
            a.nome  as aluno,
            d.nome  as disciplina,
            d.nota_minima,
            d.frequencia_minima,
            m.nota1, m.nota2, m.nota_final,
            m.status,
            m.data_inicio,
            m.data_conclusao,
            -- Contar presenças e total de aulas APENAS para a matrícula específica
            (SELECT COUNT(p_sub.id) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as total_aulas,
            (SELECT SUM(p_sub.presente) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as presencas,
            (SELECT SUM(p_sub.fez_atividade) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as atividades
        FROM matriculas m
        JOIN alunos      a ON m.aluno_id      = a.id
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

    if status_filtro and status_filtro != 'todos':
        query += " AND m.status = ?"
        params.append(status_filtro)

    query += """
        ORDER BY a.nome, d.nome
    """

    cursor.execute(query, tuple(params))
    dados = cursor.fetchall()
    conn.close()
    return dados

@app.route("/relatorios")
@login_required
def relatorios():
    # Obter parâmetros de filtro da URL
    disciplina_id = request.args.get("disciplina_id")
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro)

    conn = conectar()
    cursor = conn.cursor()
    # Obter todas as disciplinas para o filtro
    cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome")
    disciplinas_filtro = cursor.fetchall()
    conn.close()

    return render_template("relatorios.html",
        dados=dados,
        disciplinas=disciplinas_filtro,
        # Passar os filtros atuais para manter a seleção no formulário
        selected_disciplina=disciplina_id,
        selected_data_inicio=data_inicio,
        selected_data_fim=data_fim,
        selected_status=status_filtro)


@app.route("/relatorios/download/pdf")
@login_required
def download_relatorio_pdf():
    disciplina_id = request.args.get("disciplina_id")
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro)

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=30, leftMargin=30,
                            topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    elements = []

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
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')), # Dark header
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Aluno à esquerda
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#dee2e6')), # Light gray grid
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
    disciplina_id = request.args.get("disciplina_id")
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro)

    document = Document()
    # Definir idioma para português (forma correta para python-docx)
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
        table.style = 'Table Grid' # Estilo de tabela com bordas
        hdr_cells = table.rows[0].cells
        for i, header in enumerate(headers):
            p = hdr_cells[i].paragraphs[0]
            p.text = header
            p.runs[0].bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            hdr_cells[i].width = Inches(1.0) # Largura padrão, pode ajustar

        # Linhas de dados
        for item in dados:
            row_cells = table.add_row().cells

            freq_val = "—"
            if item['total_aulas'] is not None and item['total_aulas'] > 0:
                freq = ((item['presencas'] or 0) / item['total_aulas'] * 100)
                freq_val = f"{freq:.1f}% ({item['presencas'] or 0}/{item['total_aulas']})"

            media_val = "—"
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
            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT # Aluno à esquerda

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