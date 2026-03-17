import os
from datetime import date, datetime
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
        SELECT a.*, t.nome as turma_nome, t.faixa_etaria
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
        nome          = request.form.get("nome", "").strip()
        telefone      = request.form.get("telefone", "").strip()
        email         = request.form.get("email", "").strip()
        data_nascimento = request.form.get("data_nascimento", "").strip()
        membro_igreja = 1 if request.form.get("membro_igreja") else 0
        turma_id      = request.form.get("turma_id")
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT id, nome FROM turmas ORDER BY nome")
            turmas_disponiveis = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_disponiveis)
        try:
            cursor.execute("""
                INSERT INTO alunos (nome,telefone,email,data_nascimento,membro_igreja,turma_id)
                VALUES (?,?,?,?,?,?)
            """, (nome, telefone, email, data_nascimento, membro_igreja, turma_id if turma_id else None))
            conn.commit()
            flash(f"Aluno '{nome}' cadastrado!", "sucesso")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("alunos"))
    cursor.execute("SELECT id, nome FROM turmas ORDER BY nome")
    turmas_disponiveis = cursor.fetchall()
    conn.close()
    return render_template("novo_aluno.html", turmas=turmas_disponiveis)


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome          = request.form.get("nome", "").strip()
        telefone      = request.form.get("telefone", "").strip()
        email         = request.form.get("email", "").strip()
        data_nascimento = request.form.get("data_nascimento", "").strip()
        membro_igreja = 1 if request.form.get("membro_igreja") else 0
        turma_id      = request.form.get("turma_id")
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT id, nome FROM turmas ORDER BY nome")
            turmas_disponiveis = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=None, turmas=turmas_disponiveis) # Retorna com erro
        try:
            cursor.execute("""
                UPDATE alunos
                SET nome=?,telefone=?,email=?,data_nascimento=?,membro_igreja=?,turma_id=?
                WHERE id=?
            """, (nome, telefone, email, data_nascimento, membro_igreja, turma_id if turma_id else None, id))
            conn.commit()
            flash("Aluno atualizado!", "sucesso")
        except Exception as e:
            flash(f"Erro inesperado ao atualizar aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("alunos"))
    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    cursor.execute("SELECT id, nome FROM turmas ORDER BY nome")
    turmas_disponiveis = cursor.fetchall()
    conn.close()
    if not aluno:
        flash("Aluno não encontrado!", "erro")
        return redirect(url_for("alunos"))
    return render_template("editar_aluno.html",
        aluno=aluno, turmas=turmas_disponiveis)


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Excluir matrículas do aluno primeiro
        cursor.execute("DELETE FROM matriculas WHERE aluno_id=?", (id,))
        # Excluir presenças do aluno (se houver)
        cursor.execute("""
            DELETE FROM presencas WHERE matricula_id IN (
                SELECT id FROM matriculas WHERE aluno_id=?
            )
        """, (id,))
        # Excluir o aluno
        cursor.execute("DELETE FROM alunos WHERE id=?", (id,))
        conn.commit()
        flash("Aluno e suas matrículas excluídos!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir aluno: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("alunos"))


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha_aluno(id):
    conn = None
    aluno = None
    matriculas_processadas = []
    try:
        conn = conectar()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT a.id, a.nome, a.data_nascimento, t.nome AS turma_nome, t.faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno = cursor.fetchone()

        if not aluno:
            flash("Aluno não encontrado!", "erro")
            return redirect(url_for("alunos"))

        # Converter aluno para dict para poder adicionar chaves
        aluno_dict = dict(aluno)

        cursor.execute("""
            SELECT
                m.id AS matricula_id,
                m.disciplina_id,
                m.data_inicio,
                m.data_conclusao,
                m.status,
                m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                d.nome AS disciplina_nome,
                d.tem_atividades,
                t.faixa_etaria AS turma_faixa_etaria,
                (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND presente = 1) AS presencas,
                (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND fez_atividade = 1) AS atividades_feitas,
                (SELECT COUNT(DISTINCT data_aula) FROM presencas WHERE matricula_id = m.id) AS total_aulas_registradas,
                (SELECT COUNT(*) FROM aulas_disciplinas ad WHERE ad.disciplina_id = m.disciplina_id AND ad.data_aula BETWEEN m.data_inicio AND COALESCE(m.data_conclusao, DATE('now'))) AS total_aulas_previstas
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE m.aluno_id = ?
            ORDER BY d.nome
        """, (id,))
        matriculas = cursor.fetchall()

        for mat in matriculas:
            matricula_dict = dict(mat) # CONVERSÃO CRUCIAL AQUI!

            faixa_etaria = matricula_dict.get('turma_faixa_etaria', 'adultos')
            presencas = matricula_dict.get('presencas', 0)
            atividades_feitas = matricula_dict.get('atividades_feitas', 0)

            # Calcular total de aulas (usar o máximo entre registradas e previstas)
            total_aulas_registradas = matricula_dict.get('total_aulas_registradas', 0)
            total_aulas_previstas = matricula_dict.get('total_aulas_previstas', 0)
            total_aulas = max(total_aulas_registradas, total_aulas_previstas)
            matricula_dict['total_aulas'] = total_aulas

            # Frequência
            frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
            matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

            frequencia_minima = 75.0 # Padrão
            if faixa_etaria and faixa_etaria.startswith('criancas'):
                frequencia_minima = 60.0 # Exemplo: Crianças podem ter frequência mínima menor
            matricula_dict['frequencia_minima'] = frequencia_minima

            # Status de Frequência (para crianças)
            if faixa_etaria and faixa_etaria.startswith('criancas'):
                if frequencia_porcentagem >= frequencia_minima:
                    matricula_dict['status_frequencia'] = 'Aprovado'
                else:
                    matricula_dict['status_frequencia'] = 'Reprovado'
            else:
                matricula_dict['status_frequencia'] = None # Não aplicável para outras faixas

            # Média Final e Status
            media_final = None
            status_display = matricula_dict['status'].capitalize() if matricula_dict['status'] else "—"

            if faixa_etaria and faixa_etaria.startswith('criancas'):
                media_final = None # Crianças não têm notas
                status_display = matricula_dict['status_frequencia'] # Status baseado na frequência
            elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
                meditacao = matricula_dict.get('meditacao', 0) or 0
                versiculos = matricula_dict.get('versiculos', 0) or 0
                desafio_nota = matricula_dict.get('desafio_nota', 0) or 0
                visitante = matricula_dict.get('visitante', 0) or 0
                media_final = meditacao + versiculos + desafio_nota + visitante
                if media_final >= 7.0 and frequencia_porcentagem >= frequencia_minima:
                    status_display = "Aprovado"
                elif media_final >= 5.0 and frequencia_porcentagem >= frequencia_minima:
                    status_display = "Aprovado (Provisório)"
                else:
                    status_display = "Reprovado"
            else: # Adultos
                nota1 = matricula_dict.get('nota1', 0) or 0
                nota2 = matricula_dict.get('nota2', 0) or 0
                if nota1 is not None and nota2 is not None:
                    media_final = (nota1 + nota2) / 2
                elif nota1 is not None:
                    media_final = nota1
                elif nota2 is not None:
                    media_final = nota2

                if media_final is not None and media_final >= 7.0 and frequencia_porcentagem >= frequencia_minima:
                    status_display = "Aprovado"
                elif media_final is not None and media_final >= 5.0 and frequencia_porcentagem >= frequencia_minima:
                    status_display = "Aprovado (Provisório)"
                else:
                    status_display = "Reprovado"

            matricula_dict['media_final'] = media_final
            matricula_dict['media_display'] = f"{media_final:.1f}" if media_final is not None else "N/A"
            matricula_dict['status_display'] = status_display

            matriculas_processadas.append(matricula_dict)

    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        print(f"ERRO NA TRILHA DO ALUNO: {e}") # Para debug no console/logs
        if conn:
            conn.close()
        return redirect(url_for("alunos"))
    finally:
        if conn:
            conn.close()

    return render_template("trilha_aluno.html",
                           aluno=aluno_dict, # Passa o dicionário mutável
                           matriculas=matriculas_processadas)


# ══════════════════════════════════════
# DISCIPLINAS
# ══════════════════════════════════════
@app.route("/disciplinas")
@login_required
def disciplinas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT d.*, p.nome AS professor_nome
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
        nome          = request.form.get("nome", "").strip()
        descricao     = request.form.get("descricao", "").strip()
        professor_id  = request.form.get("professor_id")
        tem_atividades = 1 if request.form.get("tem_atividades") else 0
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT id, nome FROM professores ORDER BY nome")
            professores_disponiveis = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=professores_disponiveis)
        try:
            cursor.execute("""
                INSERT INTO disciplinas (nome,descricao,professor_id,tem_atividades)
                VALUES (?,?,?,?)
            """, (nome, descricao, professor_id if professor_id else None, tem_atividades))
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
    cursor.execute("SELECT id, nome FROM professores ORDER BY nome")
    professores_disponiveis = cursor.fetchall()
    conn.close()
    return render_template("nova_disciplina.html", professores=professores_disponiveis)


@app.route("/disciplinas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome          = request.form.get("nome", "").strip()
        descricao     = request.form.get("descricao", "").strip()
        professor_id  = request.form.get("professor_id")
        ativa         = 1 if request.form.get("ativa") else 0
        tem_atividades = 1 if request.form.get("tem_atividades") else 0
        try:
            cursor.execute("""
                UPDATE disciplinas
                SET nome=?,descricao=?,professor_id=?,ativa=?,tem_atividades=?
                WHERE id=?
            """, (nome, descricao, professor_id if professor_id else None, ativa, tem_atividades, id))
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
    cursor.execute("SELECT id, nome FROM professores ORDER BY nome")
    professores_disponiveis = cursor.fetchall()
    conn.close()
    if not disciplina:
        flash("Disciplina não encontrada!", "erro")
        return redirect(url_for("disciplinas"))
    return render_template("editar_disciplina.html",
        disciplina=disciplina, professores=professores_disponiveis)


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há matrículas ativas para esta disciplina
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id = ? AND status = 'cursando'", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir uma disciplina com alunos matriculados e cursando. Altere o status das matrículas primeiro.", "erro")
            return redirect(url_for("disciplinas"))

        # Excluir aulas_disciplinas associadas
        cursor.execute("DELETE FROM aulas_disciplinas WHERE disciplina_id=?", (id,))
        # Excluir presenças associadas a matrículas desta disciplina
        cursor.execute("""
            DELETE FROM presencas WHERE matricula_id IN (
                SELECT id FROM matriculas WHERE disciplina_id=?
            )
        """, (id,))
        # Excluir matrículas associadas
        cursor.execute("DELETE FROM matriculas WHERE disciplina_id=?", (id,))
        # Excluir a disciplina
        cursor.execute("DELETE FROM disciplinas WHERE id=?", (id,))
        conn.commit()
        flash("Disciplina excluída!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir disciplina: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("disciplinas"))


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
        flash("Professor não encontrado!", "erro")
        return redirect(url_for("professores"))
    return render_template("editar_professor.html", professor=professor)


@app.route("/professores/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_professor(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há disciplinas associadas a este professor
        cursor.execute("SELECT COUNT(*) FROM disciplinas WHERE professor_id = ?", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir um professor que possui disciplinas associadas. Realoque as disciplinas primeiro.", "erro")
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
# PRESENCA
# ══════════════════════════════════════
@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = conectar()
    cursor = conn.cursor()

    disciplinas_disponiveis = cursor.execute("SELECT id, nome, tem_atividades FROM disciplinas WHERE ativa = 1 ORDER BY nome").fetchall()

    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    data_aula_str = request.args.get("data_aula", type=str)

    selected_disciplina = None
    alunos_chamada = []
    tem_atividades = False

    if selected_disciplina_id and data_aula_str:
        selected_disciplina = cursor.execute("SELECT id, nome, tem_atividades FROM disciplinas WHERE id = ? AND ativa = 1", (selected_disciplina_id,)).fetchone()

        if selected_disciplina:
            tem_atividades = selected_disciplina['tem_atividades']

            # Buscar alunos matriculados na disciplina com status 'cursando'
            # E também buscar o registro de presença para a data e disciplina selecionadas
            cursor.execute("""
                SELECT
                    m.id AS matricula_id,
                    a.id AS aluno_id,
                    a.nome AS aluno_nome,
                    t.faixa_etaria,
                    p.presente,
                    p.fez_atividade
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                LEFT JOIN presencas p ON p.matricula_id = m.id AND p.data_aula = ?
                WHERE m.disciplina_id = ? AND m.status = 'cursando'
                ORDER BY a.nome
            """, (data_aula_str, selected_disciplina_id))
            alunos_chamada = cursor.fetchall()

            # Se a requisição for POST, processar a chamada
            if request.method == "POST":
                try:
                    # Registrar a aula na tabela aulas_disciplinas se ainda não existir
                    cursor.execute("INSERT OR IGNORE INTO aulas_disciplinas (disciplina_id, data_aula) VALUES (?, ?)",
                                   (selected_disciplina_id, data_aula_str))
                    conn.commit()

                    for aluno_chamada in alunos_chamada:
                        matricula_id = aluno_chamada['matricula_id']
                        presente = 1 if request.form.get(f"presente_{matricula_id}") else 0
                        fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") else 0

                        # Inserir ou atualizar a presença
                        cursor.execute("""
                            INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                            VALUES (?, ?, ?, ?)
                            ON CONFLICT(matricula_id, data_aula) DO UPDATE SET
                                presente = EXCLUDED.presente,
                                fez_atividade = EXCLUDED.fez_atividade
                        """, (matricula_id, data_aula_str, presente, fez_atividade))

                        # Atualizar o status da matrícula após cada registro de presença
                        _atualizar_status_matricula(matricula_id, conn)

                    conn.commit()
                    flash("Chamada salva com sucesso!", "sucesso")
                    # Redirecionar para a mesma página com os filtros para recarregar a tabela
                    return redirect(url_for("chamada", disciplina_id=selected_disciplina_id, data_aula=data_aula_str))
                except Exception as e:
                    conn.rollback()
                    flash(f"Erro ao salvar chamada: {e}", "erro")
                    print(f"ERRO AO SALVAR CHAMADA: {e}")
        else:
            flash("Disciplina não encontrada ou inativa!", "erro")

    conn.close()
    return render_template("chamada.html",
                           disciplinas=disciplinas_disponiveis,
                           selected_disciplina=selected_disciplina,
                           data_aula=data_aula_str,
                           alunos_chamada=alunos_chamada,
                           tem_atividades=tem_atividades)


def _atualizar_status_matricula(matricula_id, conn):
    cursor = conn.cursor()

    # Obter dados da matrícula e do aluno/turma
    cursor.execute("""
        SELECT
            m.id, m.aluno_id, m.disciplina_id, m.data_inicio, m.data_conclusao,
            m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
            m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
            d.tem_atividades,
            t.faixa_etaria
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE m.id = ?
    """, (matricula_id,))
    matricula_data = cursor.fetchone()

    if not matricula_data:
        return # Matrícula não encontrada

    faixa_etaria = matricula_data['faixa_etaria']
    tem_atividades = matricula_data['tem_atividades']

    # Calcular frequência
    cursor.execute("""
        SELECT
            COUNT(*) FILTER (WHERE presente = 1) AS presencas,
            COUNT(*) FILTER (WHERE fez_atividade = 1) AS atividades_feitas,
            COUNT(DISTINCT data_aula) AS total_aulas_registradas
        FROM presencas
        WHERE matricula_id = ?
    """, (matricula_id,))
    frequencia_data = cursor.fetchone()

    presencas = frequencia_data['presencas'] if frequencia_data else 0
    atividades_feitas = frequencia_data['atividades_feitas'] if frequencia_data else 0
    total_aulas_registradas = frequencia_data['total_aulas_registradas'] if frequencia_data else 0

    # Total de aulas previstas (entre data_inicio e data_conclusao ou hoje)
    data_inicio_obj = datetime.strptime(matricula_data['data_inicio'], '%Y-%m-%d').date()
    data_fim_obj = datetime.strptime(matricula_data['data_conclusao'], '%Y-%m-%d').date() if matricula_data['data_conclusao'] else date.today()

    cursor.execute("""
        SELECT COUNT(*) FROM aulas_disciplinas
        WHERE disciplina_id = ? AND data_aula BETWEEN ? AND ?
    """, (matricula_data['disciplina_id'], data_inicio_obj, data_fim_obj))
    total_aulas_previstas = cursor.fetchone()[0]

    total_aulas = max(total_aulas_registradas, total_aulas_previstas)

    frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

    frequencia_minima = 75.0
    if faixa_etaria and faixa_etaria.startswith('criancas'):
        frequencia_minima = 60.0

    # Lógica de Status
    novo_status = matricula_data['status'] # Manter o status atual por padrão

    # Para crianças, o status é baseado apenas na frequência
    if faixa_etaria and faixa_etaria.startswith('criancas'):
        if frequencia_porcentagem >= frequencia_minima:
            novo_status = 'aprovado'
        else:
            novo_status = 'reprovado'
    else: # Adolescentes, Jovens e Adultos
        media_final = None
        if faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
            meditacao = matricula_data.get('meditacao', 0) or 0
            versiculos = matricula_data.get('versiculos', 0) or 0
            desafio_nota = matricula_data.get('desafio_nota', 0) or 0
            visitante = matricula_data.get('visitante', 0) or 0
            media_final = meditacao + versiculos + desafio_nota + visitante
        else: # Adultos
            nota1 = matricula_data.get('nota1', 0) or 0
            nota2 = matricula_data.get('nota2', 0) or 0
            if nota1 is not None and nota2 is not None:
                media_final = (nota1 + nota2) / 2
            elif nota1 is not None:
                media_final = nota1
            elif nota2 is not None:
                media_final = nota2

        if media_final is not None:
            if media_final >= 7.0 and frequencia_porcentagem >= frequencia_minima:
                novo_status = 'aprovado'
            elif media_final >= 5.0 and frequencia_porcentagem >= frequencia_minima:
                novo_status = 'cursando' # Provisório, mas ainda cursando
            else:
                novo_status = 'reprovado'
        else:
            # Se não há notas ainda, e a frequência está boa, mantém cursando
            if frequencia_porcentagem >= frequencia_minima:
                novo_status = 'cursando'
            else:
                novo_status = 'reprovado' # Reprova por frequência mesmo sem notas

    # Atualizar o status da matrícula no banco de dados
    cursor.execute("UPDATE matriculas SET status = ? WHERE id = ?", (novo_status, matricula_data['id']))
    conn.commit()


# ══════════════════════════════════════
# AULAS
# ══════════════════════════════════════
@app.route("/aulas")
@login_required
def aulas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT ad.*, d.nome AS disciplina_nome
        FROM aulas_disciplinas ad
        JOIN disciplinas d ON ad.disciplina_id = d.id
        ORDER BY ad.data_aula DESC, d.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("aulas.html", aulas=lista)


@app.route("/aulas/nova", methods=["GET", "POST"])
@login_required
def nova_aula():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_aula_str = request.form.get("data_aula", "").strip()
        if not disciplina_id or not data_aula_str:
            flash("Disciplina e Data da Aula são obrigatórios!", "erro")
            disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa = 1 ORDER BY nome").fetchall()
            conn.close()
            return render_template("nova_aula.html", disciplinas=disciplinas_disponiveis, now=date.today())
        try:
            cursor.execute("""
                INSERT INTO aulas_disciplinas (disciplina_id, data_aula)
                VALUES (?, ?)
            """, (disciplina_id, data_aula_str))
            conn.commit()
            flash("Aula cadastrada com sucesso!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed: aulas_disciplinas.disciplina_id, aulas_disciplinas.data_aula" in str(e):
                flash("Já existe uma aula para esta disciplina nesta data!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar aula: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar aula: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("aulas"))

    disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa = 1 ORDER BY nome").fetchall()
    conn.close()
    return render_template("nova_aula.html", disciplinas=disciplinas_disponiveis, now=date.today())


@app.route("/aulas/<int:disciplina_id>/<string:data_aula>/excluir", methods=["POST"])
@login_required
def excluir_aula(disciplina_id, data_aula):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Excluir presenças associadas a esta aula
        cursor.execute("DELETE FROM presencas WHERE disciplina_id = ? AND data_aula = ?", (disciplina_id, data_aula))
        # Excluir a aula
        cursor.execute("DELETE FROM aulas_disciplinas WHERE disciplina_id = ? AND data_aula = ?", (disciplina_id, data_aula))
        conn.commit()
        flash("Aula e suas presenças excluídas!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir aula: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("aulas"))


# ══════════════════════════════════════
# RELATORIOS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id=None, data_inicio=None, data_fim=None, status_filtro=None):
    conn = conectar()
    cursor = conn.cursor()
    query = """
        SELECT
            m.id AS matricula_id,
            a.nome AS aluno,
            d.nome AS disciplina,
            m.data_inicio,
            m.data_conclusao,
            m.status,
            m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
            m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
            t.faixa_etaria,
            d.tem_atividades,
            (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND presente = 1) AS presencas,
            (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND fez_atividade = 1) AS atividades,
            (SELECT COUNT(DISTINCT data_aula) FROM presencas WHERE matricula_id = m.id) AS total_aulas_registradas,
            (SELECT COUNT(*) FROM aulas_disciplinas ad WHERE ad.disciplina_id = m.disciplina_id AND ad.data_aula BETWEEN m.data_inicio AND COALESCE(m.data_conclusao, DATE('now'))) AS total_aulas_previstas
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
    if data_inicio:
        query += " AND m.data_inicio >= ?"
        params.append(data_inicio)
    if data_fim:
        query += " AND COALESCE(m.data_conclusao, DATE('now')) <= ?"
        params.append(data_fim)
    if status_filtro and status_filtro != 'todos':
        query += " AND m.status = ?"
        params.append(status_filtro)

    query += " ORDER BY a.nome, d.nome"

    cursor.execute(query, params)
    raw_data = cursor.fetchall()

    processed_data = []
    for item in raw_data:
        item_dict = dict(item) # CONVERSÃO CRUCIAL AQUI!

        faixa_etaria = item_dict.get('faixa_etaria', 'adultos')
        presencas = item_dict.get('presencas', 0)
        atividades_feitas = item_dict.get('atividades', 0) # Renomeado para 'atividades' na query

        total_aulas_registradas = item_dict.get('total_aulas_registradas', 0)
        total_aulas_previstas = item_dict.get('total_aulas_previstas', 0)
        total_aulas = max(total_aulas_registradas, total_aulas_previstas)
        item_dict['total_aulas'] = total_aulas

        frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
        item_dict['frequencia_porcentagem'] = frequencia_porcentagem

        frequencia_minima = 75.0
        if faixa_etaria and faixa_etaria.startswith('criancas'):
            frequencia_minima = 60.0
        item_dict['frequencia_minima'] = frequencia_minima

        media_val = "N/A"
        status_val = item_dict['status'].capitalize() if item_dict['status'] else "—"

        if faixa_etaria and faixa_etaria.startswith('criancas'):
            # Status para crianças baseado apenas na frequência
            if frequencia_porcentagem >= frequencia_minima:
                status_val = 'Aprovado'
            else:
                status_val = 'Reprovado'
        elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
            meditacao = item_dict.get('meditacao', 0) or 0
            versiculos = item_dict.get('versiculos', 0) or 0
            desafio_nota = item_dict.get('desafio_nota', 0) or 0
            visitante = item_dict.get('visitante', 0) or 0
            media_final = meditacao + versiculos + desafio_nota + visitante
            media_val = f"{media_final:.1f}"
            if media_final >= 7.0 and frequencia_porcentagem >= frequencia_minima:
                status_val = "Aprovado"
            elif media_final >= 5.0 and frequencia_porcentagem >= frequencia_minima:
                status_val = "Aprovado (Provisório)"
            else:
                status_val = "Reprovado"
        else: # Adultos
            nota1 = item_dict.get('nota1', 0) or 0
            nota2 = item_dict.get('nota2', 0) or 0
            media_final = None
            if nota1 is not None and nota2 is not None:
                media_final = (nota1 + nota2) / 2
            elif nota1 is not None:
                media_final = nota1
            elif nota2 is not None:
                media_final = nota2

            if media_final is not None:
                media_val = f"{media_final:.1f}"
                if media_final >= 7.0 and frequencia_porcentagem >= frequencia_minima:
                    status_val = "Aprovado"
                elif media_final >= 5.0 and frequencia_porcentagem >= frequencia_minima:
                    status_val = "Aprovado (Provisório)"
                else:
                    status_val = "Reprovado"
            else:
                media_val = "N/A"
                # Se não há notas, o status é determinado pela frequência
                if frequencia_porcentagem >= frequencia_minima:
                    status_val = "Cursando"
                else:
                    status_val = "Reprovado"

        item_dict['media_final_display'] = media_val
        item_dict['status_display'] = status_val
        item_dict['frequencia_display'] = f"{frequencia_porcentagem:.1f}%"

        processed_data.append(item_dict)

    conn.close()
    return processed_data


@app.route("/relatorios")
@login_required
def relatorios():
    conn = conectar()
    cursor = conn.cursor()
    disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
    conn.close()

    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro", "todos")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro)

    return render_template("relatorios.html",
                           disciplinas=disciplinas_disponiveis,
                           dados=dados,
                           selected_disciplina_id=disciplina_id,
                           selected_data_inicio=data_inicio,
                           selected_data_fim=data_fim,
                           selected_status_filtro=status_filtro)


@app.route("/relatorios/download/<format>")
@login_required
def download_relatorio(format):
    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro", "todos")

    dados = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro)

    if format == "pdf":
        return gerar_relatorio_pdf(dados)
    elif format == "docx":
        return gerar_relatorio_docx(dados)
    else:
        flash("Formato de download inválido.", "erro")
        return redirect(url_for("relatorios"))


def gerar_relatorio_pdf(dados):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Relatório de Matrículas e Desempenho", styles['h1']))
    story.append(Spacer(1, 0.2 * inch))

    headers = ["Aluno", "Disciplina", "Início", "Conclusão", "Média", "Status", "Frequência", "Atividades"]
    data = [headers]

    for item in dados:
        media_val = item['media_final_display']
        status_val = item['status_display']
        freq_val = item['frequencia_display']
        atividades_val = item.get('atividades', 0) or "—" # Usar 'atividades' da query

        data.append([
            item['aluno'],
            item['disciplina'],
            item['data_inicio'] or "—",
            item['data_conclusao'] or "Em andamento",
            media_val,
            status_val,
            freq_val,
            atividades_val
        ])

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')), # Dark header
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Aluno left-aligned
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.pdf", mimetype="application/pdf")


def gerar_relatorio_docx(dados):
    document = Document()
    document.add_heading('Relatório de Matrículas e Desempenho', level=1)

    if dados:
        table = document.add_table(rows=1, cols=8)
        table.autofit = True
        table.allow_autofit = True
        table.columns[0].width = Inches(1.5) # Aluno
        table.columns[1].width = Inches(1.5) # Disciplina
        table.columns[2].width = Inches(1.0) # Início
        table.columns[3].width = Inches(1.0) # Conclusão
        table.columns[4].width = Inches(0.8) # Média
        table.columns[5].width = Inches(1.2) # Status
        table.columns[6].width = Inches(1.0) # Frequência
        table.columns[7].width = Inches(1.0) # Atividades

        # Set header row
        headers = ["Aluno", "Disciplina", "Início", "Conclusão", "Média", "Status", "Frequência", "Atividades"]
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Set header background color (requires python-docx-oss)
            # shd = OxmlElement('w:shd')
            # shd.set(qn('w:fill'), '212529') # Hex color for dark
            # hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

        for item in dados:
            row_cells = table.add_row().cells

            media_val = item['media_final_display']
            status_val = item['status_display']
            freq_val = item['frequencia_display']
            atividades_val = item.get('atividades', 0) or "—" # Usar 'atividades' da query

            row_cells[0].text = item['aluno']
            row_cells[1].text = item['disciplina']
            row_cells[2].text = item['data_inicio'] or "—"
            row_cells[3].text = item['data_conclusao'] or "Em andamento"
            row_cells[4].text = media_val
            row_cells[5].text = status_val
            row_cells[6].text = freq_val
            row_cells[7].text = atividades_val

            # Centralizar todas as células, exceto a primeira (Aluno)
            for i in range(1, len(headers)):
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    else:
        document.add_paragraph("Nenhum relatório encontrado com os filtros aplicados.")

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/relatorios/frequencia")
@login_required
def relatorios_frequencia():
    conn = conectar()
    cursor = conn.cursor()

    disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa = 1 ORDER BY nome").fetchall()
    turmas_disponiveis = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa = 1 ORDER BY nome").fetchall()
    alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()

    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    selected_turma_id = request.args.get("turma_id", type=int)
    selected_aluno_id = request.args.get("aluno_id", type=int)
    data_inicio_str = request.args.get("data_inicio")
    data_fim_str = request.args.get("data_fim")

    dados_frequencia = []

    if selected_disciplina_id or selected_turma_id or selected_aluno_id or (data_inicio_str and data_fim_str):
        query = """
            SELECT
                a.nome AS aluno_nome,
                d.nome AS disciplina_nome,
                t.nome AS turma_nome,
                t.faixa_etaria,
                m.id AS matricula_id,
                m.data_inicio,
                m.data_conclusao,
                (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND presente = 1 AND data_aula BETWEEN COALESCE(?, m.data_inicio) AND COALESCE(?, m.data_conclusao, DATE('now'))) AS presencas,
                (SELECT COUNT(DISTINCT data_aula) FROM aulas_disciplinas ad WHERE ad.disciplina_id = d.id AND ad.data_aula BETWEEN COALESCE(?, m.data_inicio) AND COALESCE(?, m.data_conclusao, DATE('now'))) AS total_aulas_previstas_periodo,
                (SELECT COUNT(DISTINCT data_aula) FROM presencas WHERE matricula_id = m.id AND data_aula BETWEEN COALESCE(?, m.data_inicio) AND COALESCE(?, m.data_conclusao, DATE('now'))) AS total_aulas_registradas_periodo
            FROM matriculas m
            JOIN alunos a ON m.aluno_id = a.id
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE 1=1
        """
        params = [data_inicio_str, data_fim_str, data_inicio_str, data_fim_str, data_inicio_str, data_fim_str]

        if selected_disciplina_id:
            query += " AND m.disciplina_id = ?"
            params.append(selected_disciplina_id)
        if selected_turma_id:
            query += " AND a.turma_id = ?"
            params.append(selected_turma_id)
        if selected_aluno_id:
            query += " AND m.aluno_id = ?"
            params.append(selected_aluno_id)

        query += " ORDER BY a.nome, d.nome"

        cursor.execute(query, params)
        raw_frequencia_data = cursor.fetchall()

        for item in raw_frequencia_data:
            item_dict = dict(item) # CONVERSÃO CRUCIAL AQUI!

            presencas = item_dict.get('presencas', 0)
            total_aulas_registradas_periodo = item_dict.get('total_aulas_registradas_periodo', 0)
            total_aulas_previstas_periodo = item_dict.get('total_aulas_previstas_periodo', 0)

            # Usar o máximo entre aulas registradas e previstas para o total de aulas no período
            total_aulas_no_periodo = max(total_aulas_registradas_periodo, total_aulas_previstas_periodo)
            item_dict['total_aulas_no_periodo'] = total_aulas_no_periodo

            frequencia_porcentagem = (presencas / total_aulas_no_periodo * 100) if total_aulas_no_periodo > 0 else 0
            item_dict['frequencia_porcentagem'] = frequencia_porcentagem
            item_dict['frequencia_display'] = f"{frequencia_porcentagem:.1f}%"

            # Determinar status de frequência
            faixa_etaria = item_dict.get('faixa_etaria', 'adultos')
            frequencia_minima = 75.0
            if faixa_etaria and faixa_etaria.startswith('criancas'):
                frequencia_minima = 60.0

            if frequencia_porcentagem >= frequencia_minima:
                item_dict['status_frequencia'] = 'Aprovado'
            else:
                item_dict['status_frequencia'] = 'Reprovado'

            dados_frequencia.append(item_dict)

    conn.close()
    return render_template("relatorio_frequencia.html",
                           disciplinas=disciplinas_disponiveis,
                           turmas=turmas_disponiveis,
                           alunos=alunos_disponiveis,
                           dados_frequencia=dados_frequencia,
                           selected_disciplina_id=selected_disciplina_id,
                           selected_turma_id=selected_turma_id,
                           selected_aluno_id=selected_aluno_id,
                           selected_data_inicio=data_inicio_str,
                           selected_data_fim=data_fim_str,
                           now=date.today())


@app.route("/relatorios/frequencia/download/<format>")
@login_required
def download_relatorio_frequencia(format):
    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    selected_turma_id = request.args.get("turma_id", type=int)
    selected_aluno_id = request.args.get("aluno_id", type=int)
    data_inicio_str = request.args.get("data_inicio")
    data_fim_str = request.args.get("data_fim")

    conn = conectar()
    cursor = conn.cursor()
    query = """
        SELECT
            a.nome AS aluno_nome,
            d.nome AS disciplina_nome,
            t.nome AS turma_nome,
            t.faixa_etaria,
            m.id AS matricula_id,
            m.data_inicio,
            m.data_conclusao,
            (SELECT COUNT(*) FROM presencas WHERE matricula_id = m.id AND presente = 1 AND data_aula BETWEEN COALESCE(?, m.data_inicio) AND COALESCE(?, m.data_conclusao, DATE('now'))) AS presencas,
            (SELECT COUNT(DISTINCT data_aula) FROM aulas_disciplinas ad WHERE ad.disciplina_id = d.id AND ad.data_aula BETWEEN COALESCE(?, m.data_inicio) AND COALESCE(?, m.data_conclusao, DATE('now'))) AS total_aulas_previstas_periodo,
            (SELECT COUNT(DISTINCT data_aula) FROM presencas WHERE matricula_id = m.id AND data_aula BETWEEN COALESCE(?, m.data_inicio) AND COALESCE(?, m.data_conclusao, DATE('now'))) AS total_aulas_registradas_periodo
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE 1=1
    """
    params = [data_inicio_str, data_fim_str, data_inicio_str, data_fim_str, data_inicio_str, data_fim_str]

    if selected_disciplina_id:
        query += " AND m.disciplina_id = ?"
        params.append(selected_disciplina_id)
    if selected_turma_id:
        query += " AND a.turma_id = ?"
        params.append(selected_turma_id)
    if selected_aluno_id:
        query += " AND m.aluno_id = ?"
        params.append(selected_aluno_id)

    query += " ORDER BY a.nome, d.nome"

    cursor.execute(query, params)
    raw_frequencia_data = cursor.fetchall()
    conn.close()

    dados_frequencia_processados = []
    for item in raw_frequencia_data:
        item_dict = dict(item) # CONVERSÃO CRUCIAL AQUI!

        presencas = item_dict.get('presencas', 0)
        total_aulas_registradas_periodo = item_dict.get('total_aulas_registradas_periodo', 0)
        total_aulas_previstas_periodo = item_dict.get('total_aulas_previstas_periodo', 0)

        total_aulas_no_periodo = max(total_aulas_registradas_periodo, total_aulas_previstas_periodo)
        item_dict['total_aulas_no_periodo'] = total_aulas_no_periodo

        frequencia_porcentagem = (presencas / total_aulas_no_periodo * 100) if total_aulas_no_periodo > 0 else 0
        item_dict['frequencia_porcentagem'] = frequencia_porcentagem
        item_dict['frequencia_display'] = f"{frequencia_porcentagem:.1f}%"

        faixa_etaria = item_dict.get('faixa_etaria', 'adultos')
        frequencia_minima = 75.0
        if faixa_etaria and faixa_etaria.startswith('criancas'):
            frequencia_minima = 60.0

        if frequencia_porcentagem >= frequencia_minima:
            item_dict['status_frequencia'] = 'Aprovado'
        else:
            item_dict['status_frequencia'] = 'Reprovado'

        dados_frequencia_processados.append(item_dict)

    if format == "pdf":
        return gerar_relatorio_frequencia_pdf(dados_frequencia_processados)
    elif format == "docx":
        return gerar_relatorio_frequencia_docx(dados_frequencia_processados)
    else:
        flash("Formato de download inválido.", "erro")
        return redirect(url_for("relatorios_frequencia"))


def gerar_relatorio_frequencia_pdf(dados):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Relatório de Frequência de Alunos", styles['h1']))
    story.append(Spacer(1, 0.2 * inch))

    headers = ["Aluno", "Disciplina", "Turma", "Presenças", "Total Aulas", "Frequência (%)", "Status"]
    data = [headers]

    for item in dados:
        data.append([
            item['aluno_nome'],
            item['disciplina_nome'],
            item['turma_nome'] or "N/A",
            item['presencas'],
            item['total_aulas_no_periodo'],
            item['frequencia_display'],
            item['status_frequencia']
        ])

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.pdf", mimetype="application/pdf")


def gerar_relatorio_frequencia_docx(dados):
    document = Document()
    document.add_heading('Relatório de Frequência de Alunos', level=1)

    if dados:
        table = document.add_table(rows=1, cols=7)
        table.autofit = True
        table.allow_autofit = True

        headers = ["Aluno", "Disciplina", "Turma", "Presenças", "Total Aulas", "Frequência (%)", "Status"]
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        for item in dados:
            row_cells = table.add_row().cells
            row_cells[0].text = item['aluno_nome']
            row_cells[1].text = item['disciplina_nome']
            row_cells[2].text = item['turma_nome'] or "N/A"
            row_cells[3].text = str(item['presencas'])
            row_cells[4].text = str(item['total_aulas_no_periodo'])
            row_cells[5].text = item['frequencia_display']
            row_cells[6].text = item['status_frequencia']

            for i in range(1, len(headers)):
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT # Aluno left-aligned

    else:
        document.add_paragraph("Nenhum relatório de frequência encontrado com os filtros aplicados.")

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


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


@app.route("/usuarios/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_usuario(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome   = request.form.get("nome", "").strip()
        email  = request.form.get("email", "").strip()
        perfil = request.form.get("perfil", "usuario")
        if not nome or not email:
            flash("Nome e E-mail são obrigatórios!", "erro")
            cursor.execute("SELECT * FROM usuarios WHERE id=?", (id,))
            usuario = cursor.fetchone()
            conn.close()
            return render_template("editar_usuario.html", usuario=usuario)
        try:
            cursor.execute("""
                UPDATE usuarios
                SET nome=?,email=?,perfil=?
                WHERE id=?
            """, (nome, email, perfil, id))
            conn.commit()
            flash("Usuário atualizado!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "usuarios.email" in str(e):
                flash("Este e-mail já está cadastrado para outro usuário!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar usuário: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao atualizar usuário: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("usuarios"))
    cursor.execute("SELECT * FROM usuarios WHERE id=?", (id,))
    usuario = cursor.fetchone()
    conn.close()
    if not usuario:
        flash("Usuário não encontrado!", "erro")
        return redirect(url_for("usuarios"))
    return render_template("editar_usuario.html", usuario=usuario)


@app.route("/usuarios/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_usuario(id):
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)