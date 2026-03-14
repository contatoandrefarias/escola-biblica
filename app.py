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
login_manager.login_message = "Faça login para continuar."
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
    flash(f"Até logo, {nome}!", "sucesso")
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
        flash("Turma não encontrada!", "erro")
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
            flash("Nome é obrigatório!", "erro")
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
            flash("Nome é obrigatório!", "erro")
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
        flash("Aluno não encontrado!", "erro")
        return redirect(url_for("alunos"))
    return render_template("editar_aluno.html",
        aluno=aluno, turmas=turmas_lista)


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
            flash("Nome é obrigatório!", "erro")
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
        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("editar_professor", id=id))
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
        flash("Professor não encontrado!", "erro")
        return redirect(url_for("professores"))
    return render_template("editar_professor.html", professor=professor)


@app.route("/professores/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_professor(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM professores WHERE id=?", (id,))
        conn.commit()
        flash("Professor excluído com sucesso!", "sucesso")
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
        nome             = request.form.get("nome", "").strip()
        descricao        = request.form.get("descricao", "").strip()
        duracao_semanas  = request.form.get("duracao_semanas", type=int)
        nota_minima      = request.form.get("nota_minima", type=float)
        frequencia_minima = request.form.get("frequencia_minima", type=float)
        tem_atividades   = 1 if request.form.get("tem_atividades") else 0
        professor_id     = request.form.get("professor_id") or None

        if not nome or not duracao_semanas or nota_minima is None or frequencia_minima is None:
            flash("Todos os campos obrigatórios devem ser preenchidos!", "erro")
            professores_lista = cursor.execute("SELECT id, nome FROM professores ORDER BY nome").fetchall()
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

    professores_lista = cursor.execute("SELECT id, nome FROM professores ORDER BY nome").fetchall()
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
            flash("Todos os campos obrigatórios devem ser preenchidos!", "erro")
            professores_lista = cursor.execute("SELECT id, nome FROM professores ORDER BY nome").fetchall()
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
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
    professores_lista = cursor.execute("SELECT id, nome FROM professores ORDER BY nome").fetchall()
    conn.close()
    if not disciplina:
        flash("Disciplina não encontrada!", "erro")
        return redirect(url_for("disciplinas"))
    return render_template("editar_disciplina.html",
        disciplina=disciplina, professores=professores_lista)


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM disciplinas WHERE id=?", (id,))
        conn.commit()
        flash("Disciplina excluída com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir disciplina: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("disciplinas"))


# ══════════════════════════════════════
# MATRICULAS
# ══════════════════════════════════════
def _atualizar_status_matricula(matricula_id, conn):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT
            m.aluno_id,
            m.disciplina_id,
            m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
            m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
            d.nota_minima, d.frequencia_minima, d.duracao_semanas,
            t.faixa_etaria
        FROM matriculas m
        JOIN disciplinas d ON m.disciplina_id = d.id
        JOIN alunos a ON m.aluno_id = a.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE m.id = ?
    """, (matricula_id,))
    matricula = cursor.fetchone()

    if not matricula:
        return

    faixa_etaria = matricula['faixa_etaria']
    nota_minima = matricula['nota_minima']
    frequencia_minima = matricula['frequencia_minima']
    duracao_semanas = matricula['duracao_semanas']

    # Calcular frequência
    cursor.execute("SELECT COUNT(*) FROM presencas WHERE matricula_id = ? AND presente = 1", (matricula_id,))
    presencas = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM presencas WHERE matricula_id = ?", (matricula_id,))
    total_aulas = cursor.fetchone()[0]

    frequencia_percentual = (presencas / total_aulas * 100) if total_aulas > 0 else 0

    status = 'cursando'
    nota_final = None

    if faixa_etaria and faixa_etaria.startswith('criancas'):
        # Crianças não têm notas, apenas frequência
        if frequencia_percentual >= frequencia_minima:
            status = 'aprovado'
        else:
            status = 'reprovado'
        nota_final = None # N/A para crianças
    elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
        # Adolescentes e Jovens: Meditação (4), Versículos (4), Desafio (1), Visitante (1)
        meditacao = matricula['meditacao'] if matricula['meditacao'] is not None else 0
        versiculos = matricula['versiculos'] if matricula['versiculos'] is not None else 0
        desafio_nota = matricula['desafio_nota'] if matricula['desafio_nota'] is not None else 0
        visitante = matricula['visitante'] if matricula['visitante'] is not None else 0

        # A soma é a Média Final
        nota_final = meditacao + versiculos + desafio_nota + visitante

        # Atualiza a nota1 com a média final para fins de exibição consistente
        cursor.execute("UPDATE matriculas SET nota1 = ? WHERE id = ?", (nota_final, matricula_id))

        if nota_final >= nota_minima and frequencia_percentual >= frequencia_minima:
            status = 'aprovado'
        elif nota_final is not None and (nota_final < nota_minima or frequencia_percentual < frequencia_minima):
            status = 'reprovado'
        else:
            status = 'cursando' # Se ainda não tem todas as notas ou frequência não calculada
    else: # Adultos
        nota1 = matricula['nota1']
        nota2 = matricula['nota2']

        if nota1 is not None and nota2 is not None:
            nota_final = (nota1 + nota2) / 2
            if nota_final >= nota_minima and frequencia_percentual >= frequencia_minima:
                status = 'aprovado'
            else:
                status = 'reprovado'
        else:
            status = 'cursando' # Se ainda não tem as duas notas

    cursor.execute("""
        UPDATE matriculas
        SET nota_final = ?, status = ?
        WHERE id = ?
    """, (nota_final, status, matricula_id))
    conn.commit()


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
            m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN turmas t ON a.turma_id = t.id
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
        aluno_id      = request.form.get("aluno_id", type=int) # Na verdade é turma_id
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio")

        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Todos os campos são obrigatórios!", "erro")
            turmas_lista = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
            disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   turmas=turmas_lista,
                                   disciplinas=disciplinas_lista,
                                   now=date.today())

        # Obter todos os alunos da turma selecionada
        cursor.execute("SELECT id FROM alunos WHERE turma_id = ?", (aluno_id,))
        alunos_da_turma = cursor.fetchall()

        if not alunos_da_turma:
            flash("Não há alunos nesta turma para matricular.", "erro")
            turmas_lista = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
            disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   turmas=turmas_lista,
                                   disciplinas=disciplinas_lista,
                                   now=date.today())

        matriculas_criadas = 0
        for aluno in alunos_da_turma:
            try:
                cursor.execute("""
                    INSERT INTO matriculas
                        (aluno_id, disciplina_id, data_inicio, status)
                    VALUES (?, ?, ?, 'cursando')
                """, (aluno['id'], disciplina_id, data_inicio))
                matriculas_criadas += 1
            except sqlite3.IntegrityError:
                # Ignora se a matrícula já existe para este aluno e disciplina
                pass
            except Exception as e:
                flash(f"Erro ao matricular aluno {aluno['id']}: {e}", "erro")
                conn.rollback()
                conn.close()
                return redirect(url_for("matriculas"))

        conn.commit()
        flash(f"{matriculas_criadas} matrículas criadas para a turma!", "sucesso")
        conn.close()
        return redirect(url_for("matriculas"))

    turmas_lista = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome").fetchall()
    disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
    conn.close()
    return render_template("nova_matricula.html",
                           turmas=turmas_lista,
                           disciplinas=disciplinas_lista,
                           now=date.today())


@app.route("/matriculas/novo_aluno_disciplina", methods=["GET", "POST"])
@login_required
def novo_aluno_disciplina():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        aluno_id      = request.form.get("aluno_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio")

        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Todos os campos são obrigatórios!", "erro")
            alunos_lista = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
            disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
            conn.close()
            return render_template("novo_aluno_disciplina.html",
                                   alunos=alunos_lista,
                                   disciplinas=disciplinas_lista,
                                   now=date.today())

        try:
            cursor.execute("""
                INSERT INTO matriculas
                    (aluno_id, disciplina_id, data_inicio, status)
                VALUES (?, ?, ?, 'cursando')
            """, (aluno_id, disciplina_id, data_inicio))
            conn.commit()
            flash("Matrícula criada com sucesso!", "sucesso")
        except sqlite3.IntegrityError:
            flash("Este aluno já está matriculado nesta disciplina!", "erro")
        except Exception as e:
            flash(f"Erro ao matricular aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    alunos_lista = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
    disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()
    conn.close()
    return render_template("novo_aluno_disciplina.html",
                           alunos=alunos_lista,
                           disciplinas=disciplinas_lista,
                           now=date.today())


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        data_conclusao = request.form.get("data_conclusao") or None

        # Notas para Adultos
        nota1 = request.form.get("nota1", type=float)
        nota2 = request.form.get("nota2", type=float)

        # Novas notas para Adolescentes/Jovens
        meditacao = request.form.get("meditacao", type=float)
        versiculos = request.form.get("versiculos", type=float)
        desafio_nota = request.form.get("desafio_nota", type=float)
        visitante = request.form.get("visitante", type=float)

        # Notas antigas (para adultos, se ainda usar)
        participacao = request.form.get("participacao", type=float)
        desafio = request.form.get("desafio", type=float)
        prova = request.form.get("prova", type=float)

        try:
            # Primeiro, obtenha a faixa etária para determinar quais campos atualizar
            cursor.execute("""
                SELECT t.faixa_etaria
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.id = ?
            """, (id,))
            matricula_info = cursor.fetchone()
            faixa_etaria = matricula_info['faixa_etaria'] if matricula_info else None

            update_query = """
                UPDATE matriculas
                SET data_conclusao = ?
            """
            update_params = [data_conclusao]

            if faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
                update_query += ", meditacao = ?, versiculos = ?, desafio_nota = ?, visitante = ?"
                update_params.extend([meditacao, versiculos, desafio_nota, visitante])
                # Para manter a compatibilidade com a coluna nota1 que agora armazena a média final
                # e para garantir que o _atualizar_status_matricula tenha os dados
                # nota1 será calculado na função _atualizar_status_matricula
            elif faixa_etaria and faixa_etaria.startswith('criancas'):
                # Não faz nada com notas para crianças
                pass
            else: # Adultos
                update_query += ", nota1 = ?, nota2 = ?, participacao = ?, desafio = ?, prova = ?"
                update_params.extend([nota1, nota2, participacao, desafio, prova])

            update_query += " WHERE id = ?"
            update_params.append(id)

            cursor.execute(update_query, tuple(update_params))
            conn.commit()

            _atualizar_status_matricula(id, conn) # Recalcula status após atualização
            flash("Matrícula atualizada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
            conn.rollback()
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    # Lógica para GET
    cursor.execute("""
        SELECT
            m.*,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.faixa_etaria,
            d.nota_minima,
            d.frequencia_minima
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

    conn.close()
    return render_template("editar_matricula.html", matricula=matricula)


@app.route("/matriculas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula excluída com sucesso!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# PRESENCA
# ══════════════════════════════════════
@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = conectar()
    cursor = conn.cursor()

    selected_disciplina = request.args.get("disciplina_id", type=int)
    selected_data_aula = request.args.get("data_aula", date.today().strftime('%Y-%m-%d'))
    alunos_chamada = []
    tem_atividades = False

    if request.method == "POST":
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_aula = request.form.get("data_aula")

        # Obter a lista de alunos para esta disciplina e data
        cursor.execute("""
            SELECT
                m.id as matricula_id
            FROM matriculas m
            WHERE m.disciplina_id = ? AND m.status = 'cursando'
        """, (disciplina_id,))
        matriculas_ids = [row['matricula_id'] for row in cursor.fetchall()]

        for matricula_id in matriculas_ids:
            presente = 1 if request.form.get(f"presente_{matricula_id}") else 0
            fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") else 0

            try:
                cursor.execute("""
                    INSERT OR REPLACE INTO presencas
                        (matricula_id, data_aula, presente, fez_atividade)
                    VALUES (?, ?, ?, ?)
                """, (matricula_id, data_aula, presente, fez_atividade))

                # Após salvar a presença, recalcular o status da matrícula
                _atualizar_status_matricula(matricula_id, conn)

            except Exception as e:
                flash(f"Erro ao registrar presença para matrícula {matricula_id}: {e}", "erro")
                conn.rollback()
                conn.close()
                return redirect(url_for("chamada", disciplina_id=disciplina_id, data_aula=data_aula))

        conn.commit()
        flash("Chamada salva com sucesso!", "sucesso")
        conn.close()
        return redirect(url_for("chamada", disciplina_id=disciplina_id, data_aula=data_aula))

    # Lógica para GET
    disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome").fetchall()

    if selected_disciplina:
        cursor.execute("SELECT tem_atividades FROM disciplinas WHERE id = ?", (selected_disciplina,))
        disc_info = cursor.fetchone()
        if disc_info:
            tem_atividades = disc_info['tem_atividades'] == 1

            cursor.execute("""
                SELECT
                    a.id as aluno_id,
                    a.nome as aluno_nome,
                    m.id as matricula_id,
                    t.faixa_etaria,
                    p.presente,
                    p.fez_atividade,
                    p.id as presenca_id
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                LEFT JOIN presencas p ON m.id = p.matricula_id AND p.data_aula = ?
                WHERE m.disciplina_id = ? AND m.status = 'cursando'
                ORDER BY a.nome
            """, (selected_data_aula, selected_disciplina))
            alunos_chamada = cursor.fetchall()
        else:
            flash("Disciplina selecionada não encontrada ou inativa.", "erro")
            selected_disciplina = None # Reset para não exibir a tabela vazia
            alunos_chamada = [] # Garante que não é None para o template

    conn.close()
    return render_template("chamada.html",
                           disciplinas=disciplinas_lista,
                           selected_disciplina=selected_disciplina,
                           selected_data_aula=selected_data_aula,
                           alunos_chamada=alunos_chamada,
                           tem_atividades=tem_atividades)


# ══════════════════════════════════════
# RELATORIOS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro, conn):
    cursor = conn.cursor()
    query = """
        SELECT
            a.nome as aluno,
            d.nome as disciplina,
            m.data_inicio,
            m.data_conclusao,
            m.nota1,
            m.nota2,
            m.nota_final,
            m.participacao,
            m.desafio,
            m.prova,
            m.meditacao,
            m.versiculos,
            m.desafio_nota,
            m.visitante,
            m.status,
            t.faixa_etaria,
            d.nota_minima,
            d.frequencia_minima,
            d.duracao_semanas,
            (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND presente = 1) as presencas,
            (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id) as total_aulas,
            (SELECT SUM(CASE WHEN fez_atividade = 1 THEN 1 ELSE 0 END) FROM presencas WHERE matricula_id = m.id) as atividades
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
    if data_inicio:
        query += " AND m.data_inicio >= ?"
        params.append(data_inicio)
    if data_fim:
        query += " AND m.data_inicio <= ?"
        params.append(data_fim)
    if status_filtro and status_filtro != "todos":
        query += " AND m.status = ?"
        params.append(status_filtro)

    query += " ORDER BY a.nome, d.nome"
    cursor.execute(query, tuple(params))
    return cursor.fetchall()


@app.route("/relatorios")
@login_required
def relatorios():
    conn   = conectar()
    cursor = conn.cursor()

    disciplinas_lista = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()

    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio   = request.args.get("data_inicio")
    data_fim      = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro", "todos")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro, conn)
    conn.close()

    return render_template("relatorios.html",
                           disciplinas=disciplinas_lista,
                           dados=dados,
                           selected_disciplina=disciplina_id,
                           selected_data_inicio=data_inicio,
                           selected_data_fim=data_fim,
                           selected_status_filtro=status_filtro)


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
                # Para adolescentes/jovens, a nota final é a soma dos componentes
                meditacao_val = item['meditacao'] if item['meditacao'] is not None else 0
                versiculos_val = item['versiculos'] if item['versiculos'] is not None else 0
                desafio_nota_val = item['desafio_nota'] if item['desafio_nota'] is not None else 0
                visitante_val = item['visitante'] if item['visitante'] is not None else 0

                # A nota1 é a Média Final para Adolescentes/Jovens
                nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
                media_val = f"{nota_final_calc:.1f}"
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
        # A query aqui deve ser mais simples, apenas para pegar o nome da disciplina
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
                # Para adolescentes/jovens, a nota final é a soma dos componentes
                meditacao_val = item['meditacao'] if item['meditacao'] is not None else 0
                versiculos_val = item['versiculos'] if item['versiculos'] is not None else 0
                desafio_nota_val = item['desafio_nota'] if item['desafio_nota'] is not None else 0
                visitante_val = item['visitante'] if item['visitante'] is not None else 0

                # A nota1 é a Média Final para Adolescentes/Jovens
                nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
                media_val = f"{nota_final_calc:.1f}"
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
            flash("Todos os campos são obrigatórios!", "erro")
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
            flash(f"Usuário '{nome}' criado!", "sucesso")
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)