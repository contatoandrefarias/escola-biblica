import os
from datetime import date
from flask import (Flask, render_template, request,
                   redirect, url_for, flash, send_file)
from flask_login import (LoginManager, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco

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
import sqlite3 # Importar sqlite3 para IntegrityError

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
        cursor.execute(
            "INSERT INTO turmas (nome,descricao,faixa_etaria) VALUES (?,?,?)",
            (nome, descricao, faixa_etaria))
        conn.commit()
        conn.close()
        flash(f"Turma '{nome}' criada!", "sucesso")
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
        cursor.execute("""
            UPDATE turmas
            SET nome=?,descricao=?,faixa_etaria=?,ativa=?
            WHERE id=?
        """, (nome, descricao, faixa_etaria, ativa, id))
        conn.commit()
        conn.close()
        flash("Turma atualizada!", "sucesso")
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

        # A verificação de email duplicado foi movida para o bloco try-except da inserção

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
            # Capturar especificamente o erro de UNIQUE constraint
            if "alunos.email" in str(e):
                flash("Este e-mail já está cadastrado para outro aluno!", "erro")
            else:
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

        # A verificação de email duplicado foi movida para o bloco try-except da atualização

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
            # Capturar especificamente o erro de UNIQUE constraint
            if "alunos.email" in str(e):
                flash("Este e-mail já está cadastrado para outro aluno!", "erro")
            else:
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


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha(id):
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    if not aluno:
        flash("Aluno nao encontrado!", "erro")
        conn.close()
        return redirect(url_for("alunos"))
    cursor.execute("""
        SELECT
            d.nome            as disciplina,
            d.duracao_semanas,
            d.nota_minima,
            d.frequencia_minima,
            m.id              as mat_id,
            m.nota1, m.nota2, m.nota_final,
            m.status,
            m.data_inicio,
            m.data_conclusao,
            -- Contar presenças e total de aulas APENAS para a matrícula específica
            (SELECT COUNT(p_sub.id) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as total_aulas,
            (SELECT SUM(p_sub.presente) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as presencas,
            (SELECT SUM(p_sub.fez_atividade) FROM presencas p_sub WHERE p_sub.matricula_id = m.id) as atividades,
            pr.nome           as professor
        FROM matriculas m
        JOIN disciplinas d   ON m.disciplina_id = d.id
        LEFT JOIN professores pr ON d.professor_id = pr.id
        LEFT JOIN presencas p    ON p.matricula_id = m.id
        WHERE m.aluno_id = ?
        GROUP BY m.id
        ORDER BY CASE m.status
            WHEN 'aprovado'  THEN 1
            WHEN 'cursando'  THEN 2
            WHEN 'reprovado' THEN 3
        END, d.nome
    """, (id,))
    trilha_dados = cursor.fetchall()
    conn.close()
    aprovadas  = sum(1 for t in trilha_dados if t["status"] == "aprovado")
    reprovadas = sum(1 for t in trilha_dados if t["status"] == "reprovado")
    em_curso   = sum(1 for t in trilha_dados if t["status"] == "cursando")
    return render_template("trilha.html",
        aluno=aluno,
        trilha=trilha_dados,
        aprovadas=aprovadas,
        reprovadas=reprovadas,
        em_curso=em_curso)


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
        nome          = request.form.get("nome", "").strip()
        telefone      = request.form.get("telefone", "").strip()
        email         = request.form.get("email", "").strip()
        especialidade = request.form.get("especialidade", "").strip()
        if not nome:
            flash("Nome e obrigatorio!", "erro")
            return redirect(url_for("novo_professor"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO professores
                    (nome,telefone,email,especialidade)
                VALUES (?,?,?,?)
            """, (nome, telefone, email, especialidade))
            conn.commit()
            flash(f"Professor '{nome}' cadastrado!", "sucesso")
            return redirect(url_for("professores"))
        except sqlite3.IntegrityError as e:
            if "professores.email" in str(e):
                flash("Este e-mail já está cadastrado para outro professor!", "erro")
            else:
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
        nome          = request.form.get("nome", "").strip()
        telefone      = request.form.get("telefone", "").strip()
        email         = request.form.get("email", "").strip()
        especialidade = request.form.get("especialidade","").strip()

        if not nome:
            flash("Nome e obrigatorio!", "erro")
            prof = cursor.fetchone() # Recarregar prof
            conn.close()
            return render_template("editar_professor.html", prof=prof)

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
            if "professores.email" in str(e):
                flash("Este e-mail já está cadastrado para outro professor!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar professor: {e}", "erro")
            prof = cursor.fetchone() # Recarregar prof
            conn.close()
            return render_template("editar_professor.html", prof=prof)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar professor: {e}", "erro")
            prof = cursor.fetchone() # Recarregar prof
            conn.close()
            return render_template("editar_professor.html", prof=prof)
        finally:
            conn.close()

    cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
    prof = cursor.fetchone()
    conn.close()
    if not prof:
        flash("Professor nao encontrado!", "erro")
        return redirect(url_for("professores"))
    return render_template("editar_professor.html", prof=prof)


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
        nome              = request.form.get("nome", "").strip()
        descricao         = request.form.get("descricao", "").strip()
        duracao_semanas   = request.form.get("duracao_semanas", type=int)
        nota_minima       = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades    = 1 if request.form.get("tem_atividades") else 0
        professor_id      = request.form.get("professor_id") or None
        ativa             = 1 if request.form.get("ativa") else 0

        if not nome or duracao_semanas is None or nota_minima is None or frequencia_minima is None:
            flash("Nome, duracao, nota minima e frequencia minima sao obrigatorios!", "erro")
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
        nome              = request.form.get("nome", "").strip()
        descricao         = request.form.get("descricao", "").strip()
        duracao_semanas   = request.form.get("duracao_semanas", type=int)
        nota_minima       = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades    = 1 if request.form.get("tem_atividades") else 0
        professor_id      = request.form.get("professor_id") or None
        ativa             = 1 if request.form.get("ativa") else 0

        if not nome or duracao_semanas is None or nota_minima is None or frequencia_minima is None:
            flash("Nome, duracao, nota minima e frequencia minima sao obrigatorios!", "erro")
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
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores_lista)
        except Exception as e:
            flash(f"Erro inesperado ao atualizar disciplina: {e}", "erro")
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            cursor.execute("SELECT id,nome FROM professores ORDER BY nome")
            professores_lista = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores_lista)
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


# ══════════════════════════════════════
# MATRICULAS
# ══════════════════════════════════════
@app.route("/matriculas")
@login_required
def matriculas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.id, a.nome as aluno_nome, d.nome as disciplina_nome,
               m.data_inicio, m.data_conclusao, m.nota_final, m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
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
        data_inicio   = request.form.get("data_inicio", "").strip()

        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Aluno, disciplina e data de inicio sao obrigatorios!", "erro")
            cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html", alunos=alunos_lista, disciplinas=disciplinas_lista)

        try:
            cursor.execute("""
                INSERT INTO matriculas
                    (aluno_id,disciplina_id,data_inicio,status)
                VALUES (?,?,?,?)
            """, (aluno_id, disciplina_id, data_inicio, 'cursando'))
            conn.commit()
            flash("Matricula criada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        except sqlite3.IntegrityError as e:
            if "matriculas.aluno_id, disciplina_id" in str(e):
                flash("Este aluno já está matriculado nesta disciplina!", "erro")
            else:
                flash(f"Erro de integridade ao criar matrícula: {e}", "erro")
            cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html", alunos=alunos_lista, disciplinas=disciplinas_lista)
        except Exception as e:
            flash(f"Erro inesperado ao criar matrícula: {e}", "erro")
            cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html", alunos=alunos_lista, disciplinas=disciplinas_lista)
        finally:
            conn.close()

    cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
    alunos_lista = cursor.fetchall()
    cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html", alunos=alunos_lista, disciplinas=disciplinas_lista)


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        nota1_str         = request.form.get("nota1", "").strip()
        nota2_str         = request.form.get("nota2", "").strip()
        nota_final_str    = request.form.get("nota_final", "").strip()
        data_inicio       = request.form.get("data_inicio", "").strip()
        data_conclusao    = request.form.get("data_conclusao", "").strip()

        nota1 = float(nota1_str) if nota1_str else None
        nota2 = float(nota2_str) if nota2_str else None
        nota_final = float(nota_final_str) if nota_final_str else None

        try:
            # Obter dados da matrícula e disciplina para cálculos
            cursor.execute("""
                SELECT m.*, d.nota_minima, d.frequencia_minima
                FROM matriculas m
                JOIN disciplinas d ON m.disciplina_id = d.id
                WHERE m.id = ?
            """, (id,))
            matricula_db = cursor.fetchone()

            if not matricula_db:
                flash("Matrícula não encontrada!", "erro")
                return redirect(url_for("matriculas"))

            disciplina_nota_minima = matricula_db['nota_minima']
            disciplina_frequencia_minima = matricula_db['frequencia_minima']

            # Calcular nota_final se não foi sobrescrita
            if nota_final is None and nota1 is not None and nota2 is not None:
                nota_final = (nota1 + nota2) / 2.0

            # Calcular frequência
            cursor.execute("""
                SELECT COUNT(id) as total_aulas, SUM(presente) as presencas
                FROM presencas
                WHERE matricula_id = ?
            """, (id,))
            presenca_data = cursor.fetchone()
            total_aulas = presenca_data['total_aulas'] or 0
            presencas   = presenca_data['presencas'] or 0

            frequencia_percentual = 0.0
            if total_aulas > 0:
                frequencia_percentual = (presencas / total_aulas) * 100.0

            # Determinar status
            novo_status = 'cursando'
            if data_conclusao: # Se a data de conclusão foi preenchida, tentar finalizar
                if nota_final is not None and frequencia_percentual is not None:
                    if nota_final >= disciplina_nota_minima and \
                       frequencia_percentual >= disciplina_frequencia_minima:
                        novo_status = 'aprovado'
                    else:
                        novo_status = 'reprovado'
                else:
                    flash("Para finalizar a matrícula, notas e frequência devem ser suficientes.", "aviso")
                    novo_status = 'cursando' # Mantém cursando se não houver dados para avaliar
            else:
                novo_status = 'cursando' # Se não tem data de conclusão, continua cursando

            # Atualizar matrícula
            cursor.execute("""
                UPDATE matriculas
                SET nota1=?, nota2=?, nota_final=?, status=?,
                    data_inicio=?, data_conclusao=?
                WHERE id=?
            """, (nota1, nota2, nota_final, novo_status,
                  data_inicio, data_conclusao, id))
            conn.commit()
            flash("Matrícula atualizada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
            conn.rollback() # Reverter em caso de erro
            # Recarregar dados para o template
            cursor.execute("""
                SELECT m.*, a.nome as aluno_nome, d.nome as disciplina_nome,
                       d.nota_minima, d.frequencia_minima
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                WHERE m.id = ?
            """, (id,))
            matricula = cursor.fetchone()

            cursor.execute("""
                SELECT COUNT(id) as total_aulas, SUM(presente) as presencas
                FROM presencas
                WHERE matricula_id = ?
            """, (id,))
            presenca_data = cursor.fetchone()
            total_aulas = presenca_data['total_aulas'] or 0
            presencas   = presenca_data['presencas'] or 0
            frequencia_percentual = (presencas / total_aulas) * 100.0 if total_aulas > 0 else 0.0

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
        SELECT m.*, a.nome as aluno_nome, d.nome as disciplina_nome,
               d.nota_minima, d.frequencia_minima
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
        SELECT COUNT(id) as total_aulas, SUM(presente) as presencas
        FROM presencas
        WHERE matricula_id = ?
    """, (id,))
    presenca_data = cursor.fetchone()
    total_aulas = presenca_data['total_aulas'] or 0
    presencas   = presenca_data['presencas'] or 0
    frequencia_percentual = (presencas / total_aulas) * 100.0 if total_aulas > 0 else 0.0

    conn.close()
    return render_template("editar_matricula.html",
        matricula=matricula,
        total_aulas=total_aulas,
        presencas=presencas,
        frequencia_percentual=frequencia_percentual)


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

    if current_user.is_aluno and current_user.aluno_id:
        aluno_id = current_user.aluno_id
        # Disciplinas que o aluno está cursando
        cursor.execute("""
            SELECT d.id, d.nome
            FROM disciplinas d
            JOIN matriculas m ON d.id = m.disciplina_id
            WHERE m.aluno_id = ? AND m.status = 'cursando' AND d.ativa = 1
            ORDER BY d.nome
        """, (aluno_id,))
        disciplinas_cursando = cursor.fetchall()

        # Disciplinas que o aluno já concluiu (aprovado/reprovado)
        cursor.execute("""
            SELECT d.id, d.nome
            FROM disciplinas d
            JOIN matriculas m ON d.id = m.disciplina_id
            WHERE m.aluno_id = ? AND (m.status = 'aprovado' OR m.status = 'reprovado') AND d.ativa = 1
            ORDER BY d.nome
        """, (aluno_id,))
        disciplinas_concluidas = cursor.fetchall()
    else:
        # Para admins/professores, todas as disciplinas ativas são listadas como 'cursando'
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa = 1 ORDER BY nome")
        disciplinas_cursando = cursor.fetchall()

    conn.close()
    return render_template("presenca.html",
        hoje=date.today().isoformat(),
        disciplinas_cursando=disciplinas_cursando,
        disciplinas_concluidas=disciplinas_concluidas)


@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = conectar()
    cursor = conn.cursor()

    if request.method == "GET":
        disciplina_id = request.args.get("disciplina_id", type=int)
        data_aula     = request.args.get("data_aula", "").strip()

        if not disciplina_id or not data_aula:
            flash("Selecione uma disciplina e uma data para a chamada.", "erro")
            conn.close()
            return redirect(url_for("presenca"))

        # Obter nome da disciplina
        cursor.execute("SELECT nome FROM disciplinas WHERE id = ?", (disciplina_id,))
        disciplina_nome = cursor.fetchone()['nome']

        # Obter alunos matriculados na disciplina
        cursor.execute("""
            SELECT a.id as aluno_id, a.nome as aluno_nome, m.id as matricula_id
            FROM alunos a
            JOIN matriculas m ON a.id = m.aluno_id
            WHERE m.disciplina_id = ? AND m.status = 'cursando'
            ORDER BY a.nome
        """, (disciplina_id,))
        alunos_matriculados = cursor.fetchall()

        # Obter presenças existentes para a data e disciplina
        presencas_existentes = {}
        if alunos_matriculados:
            matricula_ids = [aluno['matricula_id'] for aluno in alunos_matriculados]
            placeholders = ','.join('?' for _ in matricula_ids)
            cursor.execute(f"""
                SELECT matricula_id, presente, fez_atividade
                FROM presencas
                WHERE matricula_id IN ({placeholders}) AND data_aula = ?
            """, (*matricula_ids, data_aula))
            for p in cursor.fetchall():
                presencas_existentes[p['matricula_id']] = {
                    'presente': p['presente'],
                    'fez_atividade': p['fez_atividade']
                }

        conn.close()
        return render_template("chamada.html",
            disciplina_id=disciplina_id,
            disciplina_nome=disciplina_nome,
            data_aula=data_aula,
            alunos=alunos_matriculados,
            presencas_existentes=presencas_existentes)

    elif request.method == "POST":
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_aula     = request.form.get("data_aula", "").strip()

        if not disciplina_id or not data_aula:
            flash("Dados da chamada incompletos.", "erro")
            conn.close()
            return redirect(url_for("presenca"))

        try:
            # Obter todas as matrículas ativas para a disciplina para processar a chamada
            cursor.execute("""
                SELECT id FROM matriculas
                WHERE disciplina_id = ? AND status = 'cursando'
            """, (disciplina_id,))
            matriculas_ativas = cursor.fetchall()

            for mat in matriculas_ativas:
                mat_id = mat['id']
                presente = 1 if request.form.get(f"presente_{mat_id}") else 0
                fez_ativ = 1 if request.form.get(f"atividade_{mat_id}") else 0

                cursor.execute("""
                    SELECT id FROM presencas
                    WHERE matricula_id=? AND data_aula=?
                """, (mat_id, data_aula))
                existe = cursor.fetchone()

                if existe:
                    cursor.execute("""
                        UPDATE presencas
                        SET presente=?,fez_atividade=?
                        WHERE id=?
                    """, (presente, fez_ativ, existe["id"]))
                else:
                    cursor.execute("""
                        INSERT INTO presencas
                            (matricula_id,data_aula,presente,fez_atividade)
                        VALUES (?,?,?,?)
                    """, (mat_id, data_aula, presente, fez_ativ))

            conn.commit()
            flash("Chamada salva com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao salvar chamada: {e}", "erro")
            conn.rollback() # Reverter transação em caso de erro
        finally:
            conn.close()

        return redirect(url_for("presenca"))


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