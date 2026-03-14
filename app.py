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
            return render_template("nova_turma.html",
                                   selected_faixa_etaria=faixa_etaria) # Manter seleção
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
            return render_template("nova_turma.html",
                                   selected_faixa_etaria=faixa_etaria) # Manter seleção
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar turma: {e}", "erro")
            return render_template("nova_turma.html",
                                   selected_faixa_etaria=faixa_etaria) # Manter seleção
        finally:
            conn.close()
        return redirect(url_for("turmas"))
    return render_template("nova_turma.html", selected_faixa_etaria='adultos') # Default


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
            # Recarregar dados para o template em caso de erro
            cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
            turma = cursor.fetchone()
            cursor.execute(
                "SELECT * FROM alunos WHERE turma_id=? ORDER BY nome", (id,))
            alunos_turma = cursor.fetchall()
            conn.close()
            return render_template("editar_turma.html",
                                   turma=turma, alunos_turma=alunos_turma)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar turma: {e}", "erro")
            # Recarregar dados para o template em caso de erro
            cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
            turma = cursor.fetchone()
            cursor.execute(
                "SELECT * FROM alunos WHERE turma_id=? ORDER BY nome", (id,))
            alunos_turma = cursor.fetchall()
            conn.close()
            return render_template("editar_turma.html",
                                   turma=turma, alunos_turma=alunos_turma)
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
            flash(f"Erro de integridade ao atualizar aluno: {e}", "erro")
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_lista)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar aluno: {e}", "erro")
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
            return render_template("novo_professor.html")
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO professores (nome,telefone,email,especialidade)
                VALUES (?,?,?,?)
            """, (nome, telefone, email, especialidade))
            conn.commit()
            flash(f"Professor '{nome}' cadastrado!", "sucesso")
            return redirect(url_for("professores"))
        except sqlite3.IntegrityError as e:
            # Email não é mais UNIQUE, então este erro deve ser para outras restrições
            flash(f"Erro de integridade ao cadastrar professor: {e}", "erro")
            return render_template("novo_professor.html")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar professor: {e}", "erro")
            return render_template("novo_professor.html")
        finally:
            conn.close()
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
        if not nome:
            flash("Nome e obrigatorio!", "erro")
            cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
            professor = cursor.fetchone()
            conn.close()
            return render_template("editar_professor.html", professor=professor)
        try:
            cursor.execute("""
                UPDATE professores
                SET nome=?,telefone=?,email=?,especialidade=?
                WHERE id=?
            """, (nome, telefone, email, especialidade, id))
            conn.commit()
            flash("Professor atualizado!", "sucesso")
            return redirect(url_for("professores"))
        except sqlite3.IntegrityError as e:
            # Email não é mais UNIQUE, então este erro deve ser para outras restrições
            flash(f"Erro de integridade ao atualizar professor: {e}", "erro")
            cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
            professor = cursor.fetchone()
            conn.close()
            return render_template("editar_professor.html", professor=professor)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar professor: {e}", "erro")
            cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
            professor = cursor.fetchone()
            conn.close()
            return render_template("editar_professor.html", professor=professor)
        finally:
            conn.close()
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
        duracao_semanas = request.form.get("duracao_semanas", type=int)
        nota_minima     = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        professor_id    = request.form.get("professor_id") or None
        ativa           = 1 if request.form.get("ativa") else 0

        if not nome or not duracao_semanas or nota_minima is None or frequencia_minima is None:
            flash("Campos obrigatórios não preenchidos!", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html",
                                   professores=professores_lista,
                                   form_data=request.form) # Manter dados do formulário

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
            return redirect(url_for("disciplinas"))
        except sqlite3.IntegrityError as e:
            if "disciplinas.nome" in str(e):
                flash("Já existe uma disciplina com este nome!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar disciplina: {e}", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html",
                                   professores=professores_lista,
                                   form_data=request.form)
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar disciplina: {e}", "erro")
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html",
                                   professores=professores_lista,
                                   form_data=request.form)
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
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        duracao_semanas = request.form.get("duracao_semanas", type=int)
        nota_minima     = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        professor_id    = request.form.get("professor_id") or None
        ativa           = 1 if request.form.get("ativa") else 0

        if not nome or not duracao_semanas or nota_minima is None or frequencia_minima is None:
            flash("Campos obrigatórios não preenchidos!", "erro")
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html",
                                   disciplina=disciplina,
                                   professores=professores_lista)

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
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html",
                                   disciplina=disciplina,
                                   professores=professores_lista)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar disciplina: {e}", "erro")
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html",
                                   disciplina=disciplina,
                                   professores=professores_lista)
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
    conn   = conectar()
    cursor = conn.cursor()

    selected_turma      = request.args.get("turma_id", type=int)
    selected_disciplina = request.args.get("disciplina_id", type=int)

    query = """
        SELECT
            m.id as matricula_id, a.nome as aluno_nome, t.nome as turma_nome,
            d.nome as disciplina_nome, m.data_inicio, m.data_conclusao,
            m.nota1, m.nota2, m.nota_final, m.status,
            d.nota_minima, d.frequencia_minima, t.faixa_etaria,
            m.participacao, m.desafio, m.prova -- NOVAS COLUNAS
        FROM matriculas m
        JOIN alunos      a ON m.aluno_id      = a.id
        LEFT JOIN turmas t ON a.turma_id      = t.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE 1=1
    """
    params = []

    if selected_turma:
        query += " AND t.id = ?"
        params.append(selected_turma)
    if selected_disciplina:
        query += " AND d.id = ?"
        params.append(selected_disciplina)

    query += " ORDER BY t.nome, a.nome, d.nome"

    cursor.execute(query, tuple(params))
    lista = cursor.fetchall()

    cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()

    return render_template("matriculas.html",
        matriculas=lista,
        turmas=turmas_lista,
        disciplinas=disciplinas_lista,
        selected_turma=selected_turma,
        selected_disciplina=selected_disciplina)


@app.route("/matriculas/nova", methods=["GET", "POST"])
@login_required
def nova_matricula():
    conn   = conectar()
    cursor = conn.cursor()
    hoje = date.today().isoformat()

    if request.method == "POST":
        turma_id      = request.form.get("turma_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio")
        data_conclusao = request.form.get("data_conclusao")

        if not turma_id or not disciplina_id or not data_inicio:
            flash("Todos os campos obrigatórios devem ser preenchidos!", "erro")
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
                                   data_conclusao=data_conclusao,
                                   hoje=hoje)

        try:
            # Obter todos os alunos da turma selecionada
            cursor.execute("SELECT id, nome FROM alunos WHERE turma_id = ?", (turma_id,))
            alunos_da_turma = cursor.fetchall()

            if not alunos_da_turma:
                flash("Nenhum aluno encontrado nesta turma para matricular.", "aviso")
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
                                       data_conclusao=data_conclusao,
                                       hoje=hoje)

            matriculas_criadas = 0
            matriculas_existentes = 0

            for aluno in alunos_da_turma:
                aluno_id = aluno['id']
                # Verificar se a matrícula já existe
                cursor.execute("""
                    SELECT id FROM matriculas
                    WHERE aluno_id = ? AND disciplina_id = ?
                """, (aluno_id, disciplina_id))
                if cursor.fetchone():
                    matriculas_existentes += 1
                else:
                    cursor.execute("""
                        INSERT INTO matriculas
                            (aluno_id, disciplina_id, data_inicio, data_conclusao, status)
                        VALUES (?, ?, ?, ?, 'cursando')
                    """, (aluno_id, disciplina_id, data_inicio, data_conclusao))
                    matriculas_criadas += 1

            conn.commit()

            if matriculas_criadas > 0:
                flash(f"{matriculas_criadas} matrícula(s) criada(s) com sucesso!", "sucesso")
            if matriculas_existentes > 0:
                flash(f"{matriculas_existentes} aluno(s) já estavam matriculado(s) nesta disciplina.", "aviso")

            return redirect(url_for("matriculas"))

        except Exception as e:
            flash(f"Erro inesperado ao matricular turma: {e}", "erro")
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
                                   data_conclusao=data_conclusao,
                                   hoje=hoje)
        finally:
            conn.close()

    cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_lista = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html",
                           turmas=turmas_lista,
                           disciplinas=disciplinas_lista,
                           hoje=hoje)


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        data_inicio   = request.form.get("data_inicio")
        data_conclusao = request.form.get("data_conclusao")

        # Obter a faixa etária da turma do aluno para aplicar a lógica de notas
        cursor.execute("""
            SELECT t.faixa_etaria, d.nota_minima, d.frequencia_minima
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN turmas t ON a.turma_id = t.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.id = ?
        """, (id,))
        matricula_info = cursor.fetchone()

        if not matricula_info:
            flash("Matrícula não encontrada.", "erro")
            conn.close()
            return redirect(url_for("matriculas"))

        faixa_etaria = matricula_info['faixa_etaria']
        nota_minima = matricula_info['nota_minima']
        frequencia_minima = matricula_info['frequencia_minima']

        nota1 = None
        nota2 = None
        nota_final = None
        status = 'cursando'
        participacao = None
        desafio = None
        prova = None

        if faixa_etaria == 'criancas':
            # Crianças não têm notas, status baseado apenas na frequência
            nota1 = None
            nota2 = None
            nota_final = None
            participacao = None
            desafio = None
            prova = None
        elif faixa_etaria == 'adolescentes_jovens':
            # Adolescentes/Jovens: N1 = Participação (1) + Desafio (1) + Prova (8)
            participacao = request.form.get("participacao", type=float)
            desafio      = request.form.get("desafio", type=float)
            prova        = request.form.get("prova", type=float)

            if participacao is not None and desafio is not None and prova is not None:
                nota1 = (participacao or 0) + (desafio or 0) + (prova or 0)
                nota_final = nota1
            else:
                nota1 = None
                nota_final = None
            nota2 = None # Não usa N2
        else: # Adultos
            nota1 = request.form.get("nota1", type=float)
            nota2 = request.form.get("nota2", type=float)

            if nota1 is not None and nota2 is not None:
                nota_final = (nota1 + nota2) / 2
            elif nota1 is not None:
                nota_final = nota1 # Se só tem N1, usa N1 como final
            elif nota2 is not None:
                nota_final = nota2 # Se só tem N2, usa N2 como final
            else:
                nota_final = None

        # Sobrescrever nota_final se fornecida (para adultos e adolescentes/jovens)
        if faixa_etaria != 'criancas':
            if request.form.get("nota_final_override"):
                nota_final = request.form.get("nota_final_override", type=float)

        # Calcular frequência para determinar status
        cursor.execute("""
            SELECT
                COUNT(p.id) as total_aulas,
                SUM(p.presente) as presencas
            FROM presencas p
            WHERE p.matricula_id = ?
        """, (id,))
        frequencia_data = cursor.fetchone()
        total_aulas = frequencia_data['total_aulas'] or 0
        presencas = frequencia_data['presencas'] or 0

        frequencia_percentual = 0.0
        if total_aulas > 0:
            frequencia_percentual = (presencas / total_aulas) * 100

        # Lógica para determinar o status
        if data_conclusao: # Se a disciplina foi concluída
            if faixa_etaria == 'criancas':
                if frequencia_percentual >= frequencia_minima:
                    status = 'aprovado'
                else:
                    status = 'reprovado'
            elif faixa_etaria == 'adolescentes_jovens':
                if nota_final is not None and nota_final >= nota_minima and \
                   frequencia_percentual >= frequencia_minima:
                    status = 'aprovado'
                else:
                    status = 'reprovado'
            else: # Adultos
                if nota_final is not None and nota_final >= nota_minima and \
                   frequencia_percentual >= frequencia_minima:
                    status = 'aprovado'
                else:
                    status = 'reprovado'
        else:
            status = 'cursando' # Ainda não concluída

        try:
            cursor.execute("""
                UPDATE matriculas
                SET data_inicio=?, data_conclusao=?,
                    nota1=?, nota2=?, nota_final=?, status=?,
                    participacao=?, desafio=?, prova=?
                WHERE id=?
            """, (data_inicio, data_conclusao,
                  nota1, nota2, nota_final, status,
                  participacao, desafio, prova, id))
            conn.commit()
            flash("Matrícula atualizada!", "sucesso")
            return redirect(url_for("matriculas"))
        except Exception as e:
            flash(f"Erro inesperado ao atualizar matrícula: {e}", "erro")
            # Recarregar dados para o template em caso de erro
            cursor.execute("""
                SELECT
                    m.*, a.nome as aluno_nome, d.nome as disciplina_nome,
                    d.nota_minima, d.frequencia_minima, t.faixa_etaria, t.nome as turma_nome
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.id = ?
            """, (id,))
            matricula = cursor.fetchone()
            conn.close()
            return render_template("editar_matricula.html",
                                   matricula=matricula,
                                   total_aulas=total_aulas,
                                   presencas=presencas,
                                   frequencia_percentual=frequencia_percentual)
        finally:
            conn.close()

    # GET request
    cursor.execute("""
        SELECT
            m.*, a.nome as aluno_nome, d.nome as disciplina_nome,
            d.nota_minima, d.frequencia_minima, t.faixa_etaria, t.nome as turma_nome
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN turmas t ON a.turma_id = t.id
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
        WHERE p.matricula_id = ?
    """, (id,))
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
    conn   = conectar()
    cursor = conn.cursor()

    # Obter todas as disciplinas ativas
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_ativas = cursor.fetchall()

    # Obter disciplinas que o usuário logado está cursando (se for aluno)
    disciplinas_aluno = []
    if current_user.is_aluno:
        aluno_id = current_user.aluno_id
        if aluno_id:
            cursor.execute("""
                SELECT d.id, d.nome, m.status
                FROM matriculas m
                JOIN disciplinas d ON m.disciplina_id = d.id
                WHERE m.aluno_id = ?
                ORDER BY d.nome
            """, (aluno_id,))
            disciplinas_aluno = cursor.fetchall()

    conn.close()
    return render_template("presenca.html",
                           disciplinas_ativas=disciplinas_ativas,
                           disciplinas_aluno=disciplinas_aluno)


@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def presenca_chamada():
    conn   = conectar()
    cursor = conn.cursor()

    disciplina_id = request.args.get("disciplina_id", type=int)
    data_aula     = request.args.get("data_aula")

    if not disciplina_id or not data_aula:
        flash("Disciplina e data da aula são obrigatórias.", "erro")
        conn.close()
        return redirect(url_for("presenca"))

    # Obter informações da disciplina e da turma (faixa etária)
    cursor.execute("""
        SELECT
            d.nome, d.tem_atividades, d.nota_minima, d.frequencia_minima,
            t.faixa_etaria
        FROM disciplinas d
        LEFT JOIN matriculas m ON d.id = m.disciplina_id
        LEFT JOIN alunos a ON m.aluno_id = a.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE d.id = ?
        LIMIT 1
    """, (disciplina_id,))
    disciplina_info = cursor.fetchone()

    if not disciplina_info:
        flash("Disciplina não encontrada.", "erro")
        conn.close()
        return redirect(url_for("presenca"))

    disciplina_nome = disciplina_info['nome']
    tem_atividades  = disciplina_info['tem_atividades']
    faixa_etaria    = disciplina_info['faixa_etaria']
    nota_minima_disc = disciplina_info['nota_minima']
    frequencia_minima_disc = disciplina_info['frequencia_minima']

    if request.method == "POST":
        try:
            # Primeiro, verificar se já existe um registro de presença para esta disciplina e data
            cursor.execute("""
                SELECT p.id, p.matricula_id
                FROM presencas p
                JOIN matriculas m ON p.matricula_id = m.id
                WHERE m.disciplina_id = ? AND p.data_aula = ?
            """, (disciplina_id, data_aula))
            presencas_existentes = cursor.fetchall()

            if presencas_existentes:
                # Atualizar presenças existentes
                for p_existente in presencas_existentes:
                    matricula_id = p_existente['matricula_id']
                    presente_val = 1 if request.form.get(f"presente_{matricula_id}") else 0
                    atividade_val = 1 if tem_atividades and request.form.get(f"atividade_{matricula_id}") else 0

                    cursor.execute("""
                        UPDATE presencas
                        SET presente = ?, fez_atividade = ?
                        WHERE id = ?
                    """, (presente_val, atividade_val, p_existente['id']))
            else:
                # Inserir novas presenças
                # Obter todas as matrículas ativas para esta disciplina
                cursor.execute("""
                    SELECT m.id as matricula_id
                    FROM matriculas m
                    WHERE m.disciplina_id = ? AND m.status = 'cursando'
                """, (disciplina_id,))
                matriculas_ativas = cursor.fetchall()

                for mat in matriculas_ativas:
                    matricula_id = mat['matricula_id']
                    presente_val = 1 if request.form.get(f"presente_{matricula_id}") else 0
                    atividade_val = 1 if tem_atividades and request.form.get(f"atividade_{matricula_id}") else 0

                    cursor.execute("""
                        INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                        VALUES (?, ?, ?, ?)
                    """, (matricula_id, data_aula, presente_val, atividade_val))

            conn.commit()
            flash("Chamada salva com sucesso!", "sucesso")

            # Após salvar a chamada, re-calcular status das matrículas afetadas
            cursor.execute("""
                SELECT m.id
                FROM matriculas m
                WHERE m.disciplina_id = ? AND m.status = 'cursando'
            """, (disciplina_id,))
            matriculas_para_atualizar = cursor.fetchall()

            for mat_id_row in matriculas_para_atualizar:
                matricula_id = mat_id_row['id']
                # Chamar a lógica de atualização de status para cada matrícula
                _atualizar_status_matricula(matricula_id, conn)

            conn.commit() # Commit final após todas as atualizações de status

            return redirect(url_for("presenca_chamada",
                                   disciplina_id=disciplina_id,
                                   data_aula=data_aula))
        except Exception as e:
            flash(f"Erro ao salvar chamada: {e}", "erro")
            conn.close()
            return redirect(url_for("presenca_chamada",
                                   disciplina_id=disciplina_id,
                                   data_aula=data_aula))
        finally:
            conn.close()

    # GET request: Carregar dados para exibir a chamada
    query_alunos = """
        SELECT
            a.id as aluno_id, a.nome as aluno_nome,
            m.id as matricula_id, m.nota1, m.nota2, m.nota_final, m.status,
            m.participacao, m.desafio, m.prova, -- NOVAS COLUNAS
            d.nota_minima, d.frequencia_minima,
            (SELECT COUNT(p_sub.id) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as total_aulas,
            (SELECT SUM(p_sub.presente) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as total_presencas,
            (SELECT p_hoje.presente FROM presencas p_hoje WHERE p_hoje.matricula_id = m.id AND p_hoje.data_aula = ?) as presente_hoje,
            (SELECT p_hoje.fez_atividade FROM presencas p_hoje WHERE p_hoje.matricula_id = m.id AND p_hoje.data_aula = ?) as fez_atividade_hoje
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE m.disciplina_id = ? AND m.status = 'cursando'
        ORDER BY a.nome
    """
    cursor.execute(query_alunos, (data_aula, data_aula, disciplina_id))
    alunos_matriculados = cursor.fetchall()
    conn.close()

    return render_template("chamada.html",
                           disciplina_nome=disciplina_nome,
                           data_aula=data_aula,
                           alunos=alunos_matriculados,
                           tem_atividades=tem_atividades,
                           faixa_etaria=faixa_etaria) # Passar faixa_etaria


def _atualizar_status_matricula(matricula_id, conn):
    """Função auxiliar para recalcular e atualizar o status de uma matrícula."""
    cursor = conn.cursor()

    cursor.execute("""
        SELECT
            m.nota1, m.nota2, m.nota_final, m.data_conclusao,
            m.participacao, m.desafio, m.prova, -- NOVAS COLUNAS
            d.nota_minima, d.frequencia_minima,
            t.faixa_etaria
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN turmas t ON a.turma_id = t.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE m.id = ?
    """, (matricula_id,))
    matricula_data = cursor.fetchone()

    if not matricula_data:
        return # Matrícula não encontrada

    faixa_etaria = matricula_data['faixa_etaria']
    nota_minima = matricula_data['nota_minima']
    frequencia_minima = matricula_data['frequencia_minima']
    data_conclusao = matricula_data['data_conclusao']

    nota1 = matricula_data['nota1']
    nota2 = matricula_data['nota2']
    nota_final = matricula_data['nota_final']
    participacao = matricula_data['participacao']
    desafio = matricula_data['desafio']
    prova = matricula_data['prova']

    # Recalcular nota_final e nota1 (para adolescentes/jovens)
    if faixa_etaria == 'adolescentes_jovens':
        if participacao is not None and desafio is not None and prova is not None:
            nota1 = (participacao or 0) + (desafio or 0) + (prova or 0)
            nota_final = nota1
        else:
            nota1 = None
            nota_final = None
        nota2 = None # Garante que N2 é nulo para adolescentes/jovens
    elif faixa_etaria == 'adultos':
        if nota1 is not None and nota2 is not None:
            nota_final = (nota1 + nota2) / 2
        elif nota1 is not None:
            nota_final = nota1
        elif nota2 is not None:
            nota_final = nota2
        else:
            nota_final = None
        participacao = None # Garante que são nulos para adultos
        desafio = None
        prova = None
    else: # Criancas
        nota1 = None
        nota2 = None
        nota_final = None
        participacao = None
        desafio = None
        prova = None

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
        if faixa_etaria == 'criancas':
            if frequencia_percentual >= frequencia_minima:
                new_status = 'aprovado'
            else:
                new_status = 'reprovado'
        elif faixa_etaria == 'adolescentes_jovens':
            if nota_final is not None and nota_final >= nota_minima and \
               frequencia_percentual >= frequencia_minima:
                new_status = 'aprovado'
            else:
                new_status = 'reprovado'
        else: # Adultos
            if nota_final is not None and nota_final >= nota_minima and \
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
    """, (nota1, nota2, nota_final, new_status,
          participacao, desafio, prova, matricula_id))

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
            t.faixa_etaria, -- Adicionado faixa_etaria
            m.nota1, m.nota2, m.nota_final,
            m.participacao, m.desafio, m.prova, -- NOVAS COLUNAS
            m.status,
            m.data_inicio,
            m.data_conclusao,
            (SELECT COUNT(p_sub.id) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as total_aulas,
            (SELECT SUM(p_sub.presente) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as presencas,
            (SELECT SUM(p_sub.fez_atividade) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as atividades
        FROM matriculas m
        JOIN alunos      a ON m.aluno_id      = a.id
        LEFT JOIN turmas t ON a.turma_id      = t.id -- JOIN com turmas
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
        conn = conectar() # Nova conexão para buscar nome da disciplina
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
            if item['faixa_etaria'] == 'criancas':
                media_val = "N/A"
            elif item['faixa_etaria'] == 'adolescentes_jovens':
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
        conn = conectar() # Nova conexão para buscar nome da disciplina
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
            if item['faixa_etaria'] == 'criancas':
                media_val = "N/A"
            elif item['faixa_etaria'] == 'adolescentes_jovens':
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

