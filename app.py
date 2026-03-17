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
        data_nascimento = request.form.get("data_nascimento", "").strip() or None
        membro_igreja = 1 if request.form.get("membro_igreja") else 0
        turma_id      = request.form.get("turma_id", type=int) or None
        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("novo_aluno"))
        try:
            cursor.execute("""
                INSERT INTO alunos (nome,telefone,email,data_nascimento,membro_igreja,turma_id)
                VALUES (?,?,?,?,?,?)
            """, (nome, telefone, email, data_nascimento, membro_igreja, turma_id))
            conn.commit()
            flash(f"Aluno '{nome}' cadastrado!", "sucesso")
        except Exception as e:
            flash(f"Erro ao cadastrar aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("alunos"))
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_ativas = cursor.fetchall()
    conn.close()
    return render_template("novo_aluno.html", turmas_ativas=turmas_ativas)


@app.route("/alunos/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome          = request.form.get("nome", "").strip()
        telefone      = request.form.get("telefone", "").strip()
        email         = request.form.get("email", "").strip()
        data_nascimento = request.form.get("data_nascimento", "").strip() or None
        membro_igreja = 1 if request.form.get("membro_igreja") else 0
        turma_id      = request.form.get("turma_id", type=int) or None
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_ativas = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=aluno, turmas_ativas=turmas_ativas)
        try:
            cursor.execute("""
                UPDATE alunos
                SET nome=?,telefone=?,email=?,data_nascimento=?,membro_igreja=?,turma_id=?
                WHERE id=?
            """, (nome, telefone, email, data_nascimento, membro_igreja, turma_id, id))
            conn.commit()
            flash("Aluno atualizado!", "sucesso")
        except Exception as e:
            flash(f"Erro ao atualizar aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("alunos"))
    cursor.execute("""
        SELECT a.*, t.nome as turma_nome, t.faixa_etaria
        FROM alunos a
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE a.id=?
    """, (id,))
    aluno = cursor.fetchone()
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_ativas = cursor.fetchall()
    conn.close()
    if not aluno:
        flash("Aluno não encontrado!", "erro")
        return redirect(url_for("alunos"))
    return render_template("editar_aluno.html", aluno=aluno, turmas_ativas=turmas_ativas)


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
            DELETE FROM presencas WHERE matricula_id IN (SELECT id FROM matriculas WHERE aluno_id=?)
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
    conn = conectar()
    cursor = conn.cursor()
    aluno = None
    matriculas = []

    try:
        cursor.execute("""
            SELECT a.*, t.nome as turma_nome, t.faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno_data = cursor.fetchone()
        if aluno_data:
            aluno = dict(aluno_data) # Converter para dict mutável

            cursor.execute("""
                SELECT
                    m.id as matricula_id,
                    m.data_inicio,
                    m.data_conclusao,
                    m.status,
                    m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                    m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                    d.nome as disciplina_nome,
                    d.tem_atividades,
                    t.faixa_etaria as turma_faixa_etaria,
                    d.frequencia_minima
                FROM matriculas m
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.aluno_id = ?
                ORDER BY d.nome
            """, (id,))
            matriculas_raw = cursor.fetchall()

            for mat in matriculas_raw:
                matricula_dict = dict(mat) # Converter para dict mutável

                # --- Cálculo de Frequência ---
                cursor.execute("""
                    SELECT COUNT(DISTINCT data_aula) FROM presencas
                    WHERE matricula_id = ? AND presente = 1
                """, (matricula_dict['matricula_id'],))
                presencas = cursor.fetchone()[0]

                cursor.execute("""
                    SELECT COUNT(DISTINCT data_aula) FROM presencas
                    WHERE matricula_id = ?
                """, (matricula_dict['matricula_id'],))
                total_aulas = cursor.fetchone()[0]

                matricula_dict['presencas'] = presencas
                matricula_dict['total_aulas'] = total_aulas
                matricula_dict['frequencia_porcentagem'] = (presencas / total_aulas * 100) if total_aulas > 0 else 0

                # --- Cálculo de Atividades ---
                if matricula_dict['tem_atividades']:
                    cursor.execute("""
                        SELECT COUNT(DISTINCT data_aula) FROM presencas
                        WHERE matricula_id = ? AND fez_atividade = 1
                    """, (matricula_dict['matricula_id'],))
                    atividades_feitas = cursor.fetchone()[0]
                    matricula_dict['atividades_feitas'] = atividades_feitas
                else:
                    matricula_dict['atividades_feitas'] = 0 # Garante que a chave existe

                # --- Cálculo de Média Final e Status ---
                faixa_etaria = matricula_dict['turma_faixa_etaria']

                if faixa_etaria and faixa_etaria.startswith('criancas'):
                    matricula_dict['media_final'] = None
                    matricula_dict['media_display'] = 'N/A'
                    # Status para crianças baseado apenas na frequência
                    if matricula_dict['frequencia_porcentagem'] >= matricula_dict['frequencia_minima']:
                        matricula_dict['status_frequencia'] = 'Aprovado'
                    else:
                        matricula_dict['status_frequencia'] = 'Reprovado'
                    matricula_dict['status_display'] = f"{matricula_dict['status_frequencia']} (Frequência)"

                elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
                    meditacao_val = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                    versiculos_val = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                    desafio_nota_val = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                    visitante_val = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0

                    nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
                    matricula_dict['media_final'] = nota_final_calc
                    matricula_dict['media_display'] = f"{nota_final_calc:.1f}"

                    # Status para Adolescentes/Jovens
                    if matricula_dict['status'] == 'cursando':
                        if matricula_dict['frequencia_porcentagem'] >= matricula_dict['frequencia_minima'] and nota_final_calc >= 6.0: # Exemplo de critério de aprovação
                            matricula_dict['status_display'] = 'Aprovado (Provisório)'
                        elif matricula_dict['frequencia_porcentagem'] < matricula_dict['frequencia_minima'] or nota_final_calc < 6.0:
                            matricula_dict['status_display'] = 'Reprovado (Provisório)'
                        else:
                            matricula_dict['status_display'] = 'Cursando'
                    else:
                        matricula_dict['status_display'] = matricula_dict['status'].replace('aprovado', 'Aprovado').replace('reprovado', 'Reprovado').title()

                else: # Adultos
                    nota1_val = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                    nota2_val = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0

                    if nota1_val > 0 or nota2_val > 0: # Evita divisão por zero se ambas forem 0
                        media_final_calc = (nota1_val + nota2_val) / 2
                    else:
                        media_final_calc = 0 # Ou None, dependendo da regra de negócio

                    matricula_dict['media_final'] = media_final_calc
                    matricula_dict['media_display'] = f"{media_final_calc:.1f}"

                    # Status para Adultos
                    if matricula_dict['status'] == 'cursando':
                        if matricula_dict['frequencia_porcentagem'] >= matricula_dict['frequencia_minima'] and media_final_calc >= 7.0: # Exemplo de critério de aprovação
                            matricula_dict['status_display'] = 'Aprovado (Provisório)'
                        elif matricula_dict['frequencia_porcentagem'] < matricula_dict['frequencia_minima'] or media_final_calc < 7.0:
                            matricula_dict['status_display'] = 'Reprovado (Provisório)'
                        else:
                            matricula_dict['status_display'] = 'Cursando'
                    else:
                        matricula_dict['status_display'] = matricula_dict['status'].replace('aprovado', 'Aprovado').replace('reprovado', 'Reprovado').title()

                # --- Histórico de Chamadas para esta Matrícula ---
                cursor.execute("""
                    SELECT data_aula, presente, fez_atividade
                    FROM presencas
                    WHERE matricula_id = ?
                    ORDER BY data_aula DESC
                """, (matricula_dict['matricula_id'],))
                matricula_dict['historico_chamadas'] = [dict(row) for row in cursor.fetchall()] # Converter para lista de dicts

                matriculas.append(matricula_dict)

    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        app.logger.error(f"Erro na trilha do aluno {id}: {e}", exc_info=True)
        return redirect(url_for("alunos"))
    finally:
        conn.close()

    return render_template("trilha_aluno.html", aluno=aluno, matriculas=matriculas)


# ══════════════════════════════════════
# DISCIPLINAS
# ══════════════════════════════════════
@app.route("/disciplinas")
@login_required
def disciplinas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM disciplinas ORDER BY nome")
    lista = cursor.fetchall()
    conn.close()
    return render_template("disciplinas.html", disciplinas=lista)


@app.route("/disciplinas/nova", methods=["GET", "POST"])
@login_required
def nova_disciplina():
    if request.method == "POST":
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        professor_id    = request.form.get("professor_id", type=int) or None
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0 # Default 75%
        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("nova_disciplina"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO disciplinas (nome,descricao,professor_id,tem_atividades,frequencia_minima)
                VALUES (?,?,?,?,?)
            """, (nome, descricao, professor_id, tem_atividades, frequencia_minima))
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
    conn   = conectar()
    cursor = conn.cursor()
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
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        professor_id    = request.form.get("professor_id", type=int) or None
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        ativa           = 1 if request.form.get("ativa") else 0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            cursor.execute("SELECT id, nome FROM professores ORDER BY nome")
            professores_disponiveis = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores_disponiveis)
        try:
            cursor.execute("""
                UPDATE disciplinas
                SET nome=?,descricao=?,professor_id=?,tem_atividades=?,ativa=?,frequencia_minima=?
                WHERE id=?
            """, (nome, descricao, professor_id, tem_atividades, ativa, frequencia_minima, id))
            conn.commit()
            flash("Disciplina atualizada!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "disciplinas.nome" in str(e):
                flash("Já existe uma disciplina com este nome!", "erro")
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
    return render_template("editar_disciplina.html", disciplina=disciplina, professores=professores_disponiveis)


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há matrículas ativas para esta disciplina
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id = ? AND status = 'cursando'", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir uma disciplina com alunos 'cursando'. Altere o status das matrículas primeiro.", "erro")
            return redirect(url_for("disciplinas"))

        # Excluir presenças relacionadas às matrículas desta disciplina
        cursor.execute("""
            DELETE FROM presencas WHERE matricula_id IN (SELECT id FROM matriculas WHERE disciplina_id = ?)
        """, (id,))
        # Excluir matrículas relacionadas a esta disciplina
        cursor.execute("DELETE FROM matriculas WHERE disciplina_id = ?", (id,))
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
    cursor.execute("""
        SELECT p.*, COUNT(d.id) as total_disciplinas
        FROM professores p
        LEFT JOIN disciplinas d ON d.professor_id = p.id
        GROUP BY p.id ORDER BY p.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("professores.html", professores=lista)


@app.route("/professores/novo", methods=["GET", "POST"])
@login_required
def novo_professor():
    if request.method == "POST":
        nome     = request.form.get("nome", "").strip()
        telefone = request.form.get("telefone", "").strip()
        email    = request.form.get("email", "").strip()
        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("novo_professor"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO professores (nome,telefone,email) VALUES (?,?,?)",
                (nome, telefone, email))
            conn.commit()
            flash(f"Professor '{nome}' cadastrado!", "sucesso")
        except Exception as e:
            flash(f"Erro ao cadastrar professor: {e}", "erro")
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
        nome     = request.form.get("nome", "").strip()
        telefone = request.form.get("telefone", "").strip()
        email    = request.form.get("email", "").strip()
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
            professor = cursor.fetchone()
            conn.close()
            return render_template("editar_professor.html", professor=professor)
        try:
            cursor.execute("""
                UPDATE professores
                SET nome=?,telefone=?,email=?
                WHERE id=?
            """, (nome, telefone, email, id))
            conn.commit()
            flash("Professor atualizado!", "sucesso")
        except Exception as e:
            flash(f"Erro ao atualizar professor: {e}", "erro")
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
        # Verificar se o professor está associado a alguma disciplina ativa
        cursor.execute("SELECT COUNT(*) FROM disciplinas WHERE professor_id = ? AND ativa = 1", (id,))
        if cursor.fetchone()[0] > 0:
            flash("Não é possível excluir um professor associado a disciplinas ativas. Desative ou reatribua as disciplinas primeiro.", "erro")
            return redirect(url_for("professores"))

        # Desassociar o professor de quaisquer disciplinas inativas
        cursor.execute("UPDATE disciplinas SET professor_id = NULL WHERE professor_id = ?", (id,))
        # Excluir o professor
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
    disciplinas_ativas = cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa = 1 ORDER BY nome").fetchall()
    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    data_aula = request.args.get("data_aula", str(date.today()))
    alunos_chamada = []
    selected_disciplina = None
    tem_atividades = False

    if selected_disciplina_id:
        selected_disciplina = cursor.execute("SELECT id, nome, tem_atividades FROM disciplinas WHERE id = ?", (selected_disciplina_id,)).fetchone()
        if selected_disciplina:
            tem_atividades = selected_disciplina['tem_atividades'] == 1

            if request.method == "POST":
                # Processar o formulário de chamada
                cursor.execute("""
                    SELECT m.id as matricula_id
                    FROM matriculas m
                    WHERE m.disciplina_id = ? AND m.status = 'cursando'
                """, (selected_disciplina_id,))
                matriculas_ids = [row['matricula_id'] for row in cursor.fetchall()]

                for matricula_id in matriculas_ids:
                    presente = 1 if request.form.get(f"presente_{matricula_id}") else 0
                    fez_atividade = 1 if request.form.get(f"atividade_{matricula_id}") else 0

                    # Verificar se já existe um registro para esta matrícula e data
                    cursor.execute("""
                        SELECT id FROM presencas
                        WHERE matricula_id = ? AND data_aula = ?
                    """, (matricula_id, data_aula))
                    presenca_existente = cursor.fetchone()

                    if presenca_existente:
                        cursor.execute("""
                            UPDATE presencas SET presente = ?, fez_atividade = ?
                            WHERE id = ?
                        """, (presente, fez_atividade, presenca_existente['id']))
                    else:
                        cursor.execute("""
                            INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                            VALUES (?, ?, ?, ?)
                        """, (matricula_id, data_aula, presente, fez_atividade))
                conn.commit()
                flash("Chamada salva com sucesso!", "sucesso")

                # Após salvar, atualizar o status das matrículas afetadas
                for matricula_id in matriculas_ids:
                    _atualizar_status_matricula(matricula_id)

                return redirect(url_for("chamada", disciplina_id=selected_disciplina_id, data_aula=data_aula))

            # Carregar alunos para exibição (GET ou após POST)
            cursor.execute("""
                SELECT
                    a.id as aluno_id,
                    a.nome as aluno_nome,
                    m.id as matricula_id,
                    t.faixa_etaria
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.disciplina_id = ? AND m.status = 'cursando'
                ORDER BY a.nome
            """, (selected_disciplina_id,))
            alunos_raw = cursor.fetchall()

            for aluno_raw in alunos_raw:
                aluno_dict = dict(aluno_raw)
                # Verificar presença e atividade para a data selecionada
                cursor.execute("""
                    SELECT presente, fez_atividade FROM presencas
                    WHERE matricula_id = ? AND data_aula = ?
                """, (aluno_dict['matricula_id'], data_aula))
                presenca_info = cursor.fetchone()
                aluno_dict['presente'] = presenca_info['presente'] if presenca_info else 0
                aluno_dict['fez_atividade'] = presenca_info['fez_atividade'] if presenca_info else 0
                alunos_chamada.append(aluno_dict)

    conn.close()
    return render_template("chamada.html",
                           disciplinas_ativas=disciplinas_ativas,
                           selected_disciplina=selected_disciplina,
                           data_aula=data_aula,
                           alunos_chamada=alunos_chamada,
                           tem_atividades=tem_atividades)


def _atualizar_status_matricula(matricula_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT
                m.id, m.aluno_id, m.disciplina_id, m.status,
                m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                d.tem_atividades, d.frequencia_minima,
                t.faixa_etaria
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            LEFT JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE m.id = ?
        """, (matricula_id,))
        matricula = cursor.fetchone()

        if not matricula:
            return # Matrícula não encontrada

        matricula_dict = dict(matricula) # Converter para dict mutável

        # --- Cálculo de Frequência ---
        cursor.execute("""
            SELECT COUNT(DISTINCT data_aula) FROM presencas
            WHERE matricula_id = ? AND presente = 1
        """, (matricula_id,))
        presencas = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(DISTINCT data_aula) FROM presencas
            WHERE matricula_id = ?
        """, (matricula_id,))
        total_aulas = cursor.fetchone()[0]

        frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

        # --- Cálculo de Média Final ---
        faixa_etaria = matricula_dict['faixa_etaria']
        media_final = 0.0
        status_final = matricula_dict['status'] # Mantém o status atual por padrão

        if faixa_etaria and faixa_etaria.startswith('criancas'):
            # Crianças: Status baseado apenas na frequência
            if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                status_final = 'aprovado'
            else:
                status_final = 'reprovado'
            media_final = None # Não há média para crianças

        elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
            meditacao_val = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
            versiculos_val = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
            desafio_nota_val = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
            visitante_val = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
            media_final = meditacao_val + versiculos_val + desafio_nota_val + visitante_val

            if matricula_dict['status'] == 'cursando': # Só atualiza se ainda estiver cursando
                if frequencia_porcentagem >= matricula_dict['frequencia_minima'] and media_final >= 6.0:
                    status_final = 'aprovado'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima'] or media_final < 6.0:
                    status_final = 'reprovado'
                else:
                    status_final = 'cursando' # Continua cursando se não atingiu os critérios

        else: # Adultos
            nota1_val = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
            nota2_val = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
            media_final = (nota1_val + nota2_val) / 2 if (nota1_val > 0 or nota2_val > 0) else 0

            if matricula_dict['status'] == 'cursando': # Só atualiza se ainda estiver cursando
                if frequencia_porcentagem >= matricula_dict['frequencia_minima'] and media_final >= 7.0:
                    status_final = 'aprovado'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima'] or media_final < 7.0:
                    status_final = 'reprovado'
                else:
                    status_final = 'cursando' # Continua cursando se não atingiu os critérios

        # Atualizar o status da matrícula no banco de dados
        if status_final != matricula_dict['status']:
            cursor.execute("UPDATE matriculas SET status = ? WHERE id = ?", (status_final, matricula_id))
            conn.commit()

    except Exception as e:
        app.logger.error(f"Erro ao atualizar status da matrícula {matricula_id}: {e}", exc_info=True)
    finally:
        conn.close()


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
            m.id, m.data_inicio, m.data_conclusao, m.status,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.nome as turma_nome, t.faixa_etaria
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
        turma_id      = request.form.get("turma_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio", str(date.today()))
        if not turma_id or not disciplina_id:
            flash("Turma e Disciplina são obrigatórios!", "erro")
            return redirect(url_for("nova_matricula"))
        try:
            # Selecionar todos os alunos da turma
            cursor.execute("SELECT id FROM alunos WHERE turma_id = ?", (turma_id,))
            alunos_da_turma = cursor.fetchall()

            if not alunos_da_turma:
                flash("Não há alunos nesta turma para matricular!", "erro")
                return redirect(url_for("nova_matricula"))

            for aluno_row in alunos_da_turma:
                aluno_id = aluno_row['id']
                # Verificar se o aluno já está matriculado nesta disciplina
                cursor.execute("""
                    SELECT COUNT(*) FROM matriculas
                    WHERE aluno_id = ? AND disciplina_id = ?
                """, (aluno_id, disciplina_id))
                if cursor.fetchone()[0] == 0: # Se não estiver matriculado, insere
                    cursor.execute("""
                        INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, status)
                        VALUES (?, ?, ?, ?)
                    """, (aluno_id, disciplina_id, data_inicio, 'cursando'))
            conn.commit()
            flash("Matrícula da turma realizada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao matricular turma: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_disponiveis = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_disponiveis = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html",
                           turmas=turmas_disponiveis,
                           disciplinas=disciplinas_disponiveis,
                           now=date.today())


@app.route("/matriculas/novo_aluno_disciplina", methods=["GET", "POST"])
@login_required
def novo_aluno_disciplina():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        aluno_id      = request.form.get("aluno_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio   = request.form.get("data_inicio", str(date.today()))
        if not aluno_id or not disciplina_id:
            flash("Aluno e Disciplina são obrigatórios!", "erro")
            return redirect(url_for("novo_aluno_disciplina"))
        try:
            cursor.execute("""
                INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, status)
                VALUES (?, ?, ?, ?)
            """, (aluno_id, disciplina_id, data_inicio, 'cursando'))
            conn.commit()
            flash("Matrícula realizada com sucesso!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "UNIQUE constraint failed: matriculas.aluno_id, matriculas.disciplina_id" in str(e):
                flash("Este aluno já está matriculado nesta disciplina!", "erro")
            else:
                flash(f"Erro de integridade ao matricular: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao matricular: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos_disponiveis = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome")
    disciplinas_disponiveis = cursor.fetchall()
    conn.close()
    return render_template("novo_aluno_disciplina.html",
                           alunos=alunos_disponiveis,
                           disciplinas=disciplinas_disponiveis,
                           now=date.today())


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        data_conclusao = request.form.get("data_conclusao", "").strip() or None
        status         = request.form.get("status", "").strip()

        # Campos de notas para Adultos
        nota1_adulto = request.form.get("nota1_adulto", type=float)
        nota2_adulto = request.form.get("nota2_adulto", type=float)
        participacao_adulto = request.form.get("participacao_adulto", type=float)
        desafio_adulto = request.form.get("desafio_adulto", type=float)
        prova_adulto = request.form.get("prova_adulto", type=float)

        # Campos de notas para Adolescentes/Jovens
        meditacao_aj = request.form.get("meditacao_aj", type=float)
        versiculos_aj = request.form.get("versiculos_aj", type=float)
        desafio_nota_aj = request.form.get("desafio_nota_aj", type=float)
        visitante_aj = request.form.get("visitante_aj", type=float)

        try:
            # Primeiro, obter a faixa etária da turma do aluno para saber qual lógica de nota aplicar
            cursor.execute("""
                SELECT t.faixa_etaria
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.id = ?
            """, (id,))
            faixa_etaria_matricula = cursor.fetchone()
            faixa_etaria_matricula = faixa_etaria_matricula['faixa_etaria'] if faixa_etaria_matricula else 'adultos' # Default para adultos

            # Resetar todas as notas para NULL antes de aplicar
            # Isso evita que notas de uma faixa etária sejam persistidas para outra
            cursor.execute("""
                UPDATE matriculas SET
                    nota1 = NULL, nota2 = NULL, participacao = NULL, desafio = NULL, prova = NULL,
                    meditacao = NULL, versiculos = NULL, desafio_nota = NULL, visitante = NULL
                WHERE id = ?
            """, (id,))

            if faixa_etaria_matricula and (faixa_etaria_matricula.startswith('adolescentes') or faixa_etaria_matricula.startswith('jovens')):
                cursor.execute("""
                    UPDATE matriculas SET
                        data_conclusao = ?, status = ?,
                        meditacao = ?, versiculos = ?, desafio_nota = ?, visitante = ?
                    WHERE id = ?
                """, (data_conclusao, status,
                      meditacao_aj, versiculos_aj, desafio_nota_aj, visitante_aj, id))
            elif faixa_etaria_matricula and faixa_etaria_matricula.startswith('criancas'):
                # Crianças não têm notas, apenas data_conclusao e status
                cursor.execute("""
                    UPDATE matriculas SET
                        data_conclusao = ?, status = ?
                    WHERE id = ?
                """, (data_conclusao, status, id))
            else: # Adultos
                cursor.execute("""
                    UPDATE matriculas SET
                        data_conclusao = ?, status = ?,
                        nota1 = ?, nota2 = ?, participacao = ?, desafio = ?, prova = ?
                    WHERE id = ?
                """, (data_conclusao, status,
                      nota1_adulto, nota2_adulto, participacao_adulto, desafio_adulto, prova_adulto, id))
            conn.commit()
            flash("Matrícula atualizada!", "sucesso")

            # Após atualizar a matrícula, recalcular o status
            _atualizar_status_matricula(id)

        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    cursor.execute("""
        SELECT
            m.*,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.nome as turma_nome, t.faixa_etaria
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN turmas t ON a.turma_id = t.id
        WHERE m.id = ?
    """, (id,))
    matricula = cursor.fetchone()
    conn.close()
    if not matricula:
        flash("Matrícula não encontrada!", "erro")
        return redirect(url_for("matriculas"))
    return render_template("editar_matricula.html", matricula=matricula)


@app.route("/matriculas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Excluir presenças relacionadas a esta matrícula
        cursor.execute("DELETE FROM presencas WHERE matricula_id = ?", (id,))
        # Excluir a matrícula
        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula excluída!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# RELATORIOS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id=None, data_inicio=None, data_fim=None, status_filtro=None):
    conn = conectar()
    cursor = conn.cursor()
    query = """
        SELECT
            m.id as matricula_id,
            m.data_inicio,
            m.data_conclusao,
            m.status,
            m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
            m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            d.tem_atividades,
            t.faixa_etaria as turma_faixa_etaria,
            d.frequencia_minima
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
        query += " AND (m.data_conclusao <= ? OR m.data_conclusao IS NULL)"
        params.append(data_fim)
    if status_filtro and status_filtro != 'todos':
        query += " AND m.status = ?"
        params.append(status_filtro)

    query += " ORDER BY d.nome, a.nome"

    cursor.execute(query, params)
    relatorio_raw = cursor.fetchall()
    relatorio_data = []

    for item_raw in relatorio_raw:
        item_dict = dict(item_raw) # Converter para dict mutável

        # --- Cálculo de Frequência ---
        cursor.execute("""
            SELECT COUNT(DISTINCT data_aula) FROM presencas
            WHERE matricula_id = ? AND presente = 1
        """, (item_dict['matricula_id'],))
        presencas = cursor.fetchone()[0]

        cursor.execute("""
            SELECT COUNT(DISTINCT data_aula) FROM presencas
            WHERE matricula_id = ?
        """, (item_dict['matricula_id'],))
        total_aulas = cursor.fetchone()[0]

        item_dict['presencas'] = presencas
        item_dict['total_aulas'] = total_aulas
        item_dict['frequencia_porcentagem'] = (presencas / total_aulas * 100) if total_aulas > 0 else 0

        # --- Cálculo de Atividades ---
        if item_dict['tem_atividades']:
            cursor.execute("""
                SELECT COUNT(DISTINCT data_aula) FROM presencas
                WHERE matricula_id = ? AND fez_atividade = 1
            """, (item_dict['matricula_id'],))
            atividades_feitas = cursor.fetchone()[0]
            item_dict['atividades_feitas'] = atividades_feitas
        else:
            item_dict['atividades_feitas'] = 0

        # --- Cálculo de Média Final e Status ---
        faixa_etaria = item_dict['turma_faixa_etaria']

        if faixa_etaria and faixa_etaria.startswith('criancas'):
            item_dict['media_final'] = None
            item_dict['media_display'] = 'N/A'
            if item_dict['frequencia_porcentagem'] >= item_dict['frequencia_minima']:
                item_dict['status_display'] = 'Aprovado (Frequência)'
            else:
                item_dict['status_display'] = 'Reprovado (Frequência)'

        elif faixa_etaria and (faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens')):
            meditacao_val = item_dict['meditacao'] if item_dict['meditacao'] is not None else 0
            versiculos_val = item_dict['versiculos'] if item_dict['versiculos'] is not None else 0
            desafio_nota_val = item_dict['desafio_nota'] if item_dict['desafio_nota'] is not None else 0
            visitante_val = item_dict['visitante'] if item_dict['visitante'] is not None else 0

            nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
            item_dict['media_final'] = nota_final_calc
            item_dict['media_display'] = f"{nota_final_calc:.1f}"

            if item_dict['status'] == 'cursando':
                if item_dict['frequencia_porcentagem'] >= item_dict['frequencia_minima'] and nota_final_calc >= 6.0:
                    item_dict['status_display'] = 'Aprovado (Provisório)'
                elif item_dict['frequencia_porcentagem'] < item_dict['frequencia_minima'] or nota_final_calc < 6.0:
                    item_dict['status_display'] = 'Reprovado (Provisório)'
                else:
                    item_dict['status_display'] = 'Cursando'
            else:
                item_dict['status_display'] = item_dict['status'].replace('aprovado', 'Aprovado').replace('reprovado', 'Reprovado').title()

        else: # Adultos
            nota1_val = item_dict['nota1'] if item_dict['nota1'] is not None else 0
            nota2_val = item_dict['nota2'] if item_dict['nota2'] is not None else 0

            if nota1_val > 0 or nota2_val > 0:
                media_final_calc = (nota1_val + nota2_val) / 2
            else:
                media_final_calc = 0

            item_dict['media_final'] = media_final_calc
            item_dict['media_display'] = f"{media_final_calc:.1f}"

            if item_dict['status'] == 'cursando':
                if item_dict['frequencia_porcentagem'] >= item_dict['frequencia_minima'] and media_final_calc >= 7.0:
                    item_dict['status_display'] = 'Aprovado (Provisório)'
                elif item_dict['frequencia_porcentagem'] < item_dict['frequencia_minima'] or media_final_calc < 7.0:
                    item_dict['status_display'] = 'Reprovado (Provisório)'
                else:
                    item_dict['status_display'] = 'Cursando'
            else:
                item_dict['status_display'] = item_dict['status'].replace('aprovado', 'Aprovado').replace('reprovado', 'Reprovado').title()

        relatorio_data.append(item_dict)

    conn.close()
    return relatorio_data


@app.route("/relatorios")
@login_required
def relatorios():
    conn = conectar()
    cursor = conn.cursor()
    disciplinas = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
    conn.close()
    return render_template("relatorios.html", disciplinas=disciplinas)


@app.route("/relatorios/download/<format>")
@login_required
def download_relatorio(format):
    disciplina_id = request.args.get("disciplina_id", type=int)
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")
    status_filtro = request.args.get("status_filtro")

    relatorio_data = get_relatorio_data(disciplina_id, data_inicio, data_fim, status_filtro)

    if format == "pdf":
        return gerar_pdf_relatorio(relatorio_data)
    elif format == "docx":
        return gerar_docx_relatorio(relatorio_data)
    else:
        flash("Formato de download inválido!", "erro")
        return redirect(url_for("relatorios"))


def gerar_pdf_relatorio(data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Relatório de Matrículas e Notas", styles['h1']))
    elements.append(Spacer(1, 0.2 * inch))

    if data:
        # Headers
        headers = ["Aluno", "Disciplina", "Faixa Etária", "Início", "Conclusão", "Média", "Freq. (%)", "Status"]
        table_data = [headers]

        for item in data:
            media_display = item['media_display']
            frequencia_porcentagem = f"{item['frequencia_porcentagem']:.1f}%"
            status_display = item['status_display']

            table_data.append([
                item['aluno_nome'],
                item['disciplina_nome'],
                item['turma_faixa_etaria'].replace('_', ' ').title() if item['turma_faixa_etaria'] else 'N/A',
                item['data_inicio'].replace('-', '/') if item['data_inicio'] else '—',
                item['data_conclusao'].replace('-', '/') if item['data_conclusao'] else 'Em andamento',
                media_display,
                frequencia_porcentagem,
                status_display
            ])

        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')), # Dark header
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Aluno left-aligned
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(table)
    else:
        elements.append(Paragraph("Nenhum relatório de matrículas encontrado com os filtros aplicados.", styles['Normal']))

    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.pdf", mimetype="application/pdf")


def gerar_docx_relatorio(data):
    document = Document()
    document.add_heading('Relatório de Matrículas e Notas', level=1)

    if data:
        table = document.add_table(rows=1, cols=8)
        table.style = 'Table Grid' # Adiciona bordas

        # Header
        hdr_cells = table.rows[0].cells
        headers = ["Aluno", "Disciplina", "Faixa Etária", "Início", "Conclusão", "Média", "Freq. (%)", "Status"]
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT # Aluno left-aligned

        for item in data:
            row_cells = table.add_row().cells
            media_display = item['media_display']
            frequencia_porcentagem = f"{item['frequencia_porcentagem']:.1f}%"
            status_display = item['status_display']

            row_cells[0].text = item['aluno_nome']
            row_cells[1].text = item['disciplina_nome']
            row_cells[2].text = item['turma_faixa_etaria'].replace('_', ' ').title() if item['turma_faixa_etaria'] else 'N/A'
            row_cells[3].text = item['data_inicio'].replace('-', '/') if item['data_inicio'] else '—'
            row_cells[4].text = item['data_conclusao'].replace('-', '/') if item['data_conclusao'] else 'Em andamento'
            row_cells[5].text = str(media_display)
            row_cells[6].text = frequencia_porcentagem
            row_cells[7].text = status_display

            for i in range(1, 8): # Alinha colunas 1 a 7 ao centro
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT # Aluno left-aligned

    else:
        document.add_paragraph("Nenhum relatório de matrículas encontrado com os filtros aplicados.")

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/relatorios/frequencia")
@login_required
def relatorios_frequencia():
    conn = conectar()
    cursor = conn.cursor()
    disciplinas = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
    turmas = cursor.execute("SELECT id, nome, faixa_etaria FROM turmas ORDER BY nome").fetchall()
    alunos = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
    conn.close()

    # Parâmetros de filtro
    disciplina_id = request.args.get("disciplina_id", type=int)
    turma_id = request.args.get("turma_id", type=int)
    aluno_id = request.args.get("aluno_id", type=int)
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")

    frequencia_data = []
    if request.args: # Só busca dados se houver filtros aplicados
        query = """
            SELECT
                a.nome as aluno_nome,
                d.nome as disciplina_nome,
                t.nome as turma_nome,
                t.faixa_etaria,
                m.id as matricula_id,
                COUNT(DISTINCT p.data_aula) as total_aulas_registradas,
                SUM(CASE WHEN p.presente = 1 THEN 1 ELSE 0 END) as total_presencas,
                SUM(CASE WHEN p.fez_atividade = 1 THEN 1 ELSE 0 END) as total_atividades_feitas,
                d.tem_atividades,
                d.frequencia_minima
            FROM presencas p
            JOIN matriculas m ON p.matricula_id = m.id
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
            query += " AND a.turma_id = ?"
            params.append(turma_id)
        if aluno_id:
            query += " AND m.aluno_id = ?"
            params.append(aluno_id)
        if data_inicio:
            query += " AND p.data_aula >= ?"
            params.append(data_inicio)
        if data_fim:
            query += " AND p.data_aula <= ?"
            params.append(data_fim)

        query += """
            GROUP BY m.id, a.nome, d.nome, t.nome, t.faixa_etaria, d.tem_atividades, d.frequencia_minima
            ORDER BY a.nome, d.nome
        """

        conn = conectar()
        cursor = conn.cursor()
        cursor.execute(query, params)
        frequencia_raw = cursor.fetchall()

        for item_raw in frequencia_raw:
            item_dict = dict(item_raw) # Converter para dict mutável
            total_aulas = item_dict['total_aulas_registradas']
            total_presencas = item_dict['total_presencas']

            item_dict['frequencia_porcentagem'] = (total_presencas / total_aulas * 100) if total_aulas > 0 else 0

            # Determinar status de frequência
            if item_dict['frequencia_porcentagem'] >= item_dict['frequencia_minima']:
                item_dict['status_frequencia'] = 'Aprovado'
                item_dict['status_frequencia_badge'] = 'bg-success'
            else:
                item_dict['status_frequencia'] = 'Reprovado'
                item_dict['status_frequencia_badge'] = 'bg-danger'

            frequencia_data.append(item_dict)
        conn.close()

    return render_template("relatorio_frequencia.html",
                           disciplinas=disciplinas,
                           turmas=turmas,
                           alunos=alunos,
                           frequencia_data=frequencia_data,
                           selected_disciplina_id=disciplina_id,
                           selected_turma_id=turma_id,
                           selected_aluno_id=aluno_id,
                           selected_data_inicio=data_inicio,
                           selected_data_fim=data_fim)


@app.route("/relatorios/frequencia/download/<format>")
@login_required
def download_relatorio_frequencia(format):
    disciplina_id = request.args.get("disciplina_id", type=int)
    turma_id = request.args.get("turma_id", type=int)
    aluno_id = request.args.get("aluno_id", type=int)
    data_inicio = request.args.get("data_inicio")
    data_fim = request.args.get("data_fim")

    conn = conectar()
    cursor = conn.cursor()
    query = """
        SELECT
            a.nome as aluno_nome,
            d.nome as disciplina_nome,
            t.nome as turma_nome,
            t.faixa_etaria,
            m.id as matricula_id,
            COUNT(DISTINCT p.data_aula) as total_aulas_registradas,
            SUM(CASE WHEN p.presente = 1 THEN 1 ELSE 0 END) as total_presencas,
            SUM(CASE WHEN p.fez_atividade = 1 THEN 1 ELSE 0 END) as total_atividades_feitas,
            d.tem_atividades,
            d.frequencia_minima
        FROM presencas p
        JOIN matriculas m ON p.matricula_id = m.id
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
        query += " AND a.turma_id = ?"
        params.append(turma_id)
    if aluno_id:
        query += " AND m.aluno_id = ?"
        params.append(aluno_id)
    if data_inicio:
        query += " AND p.data_aula >= ?"
        params.append(data_inicio)
    if data_fim:
        query += " AND p.data_aula <= ?"
        params.append(data_fim)

    query += """
        GROUP BY m.id, a.nome, d.nome, t.nome, t.faixa_etaria, d.tem_atividades, d.frequencia_minima
        ORDER BY a.nome, d.nome
    """

    cursor.execute(query, params)
    frequencia_raw = cursor.fetchall()
    frequencia_data = []

    for item_raw in frequencia_raw:
        item_dict = dict(item_raw) # Converter para dict mutável
        total_aulas = item_dict['total_aulas_registradas']
        total_presencas = item_dict['total_presencas']

        item_dict['frequencia_porcentagem'] = (total_presencas / total_aulas * 100) if total_aulas > 0 else 0

        # Determinar status de frequência
        if item_dict['frequencia_porcentagem'] >= item_dict['frequencia_minima']:
            item_dict['status_frequencia'] = 'Aprovado'
        else:
            item_dict['status_frequencia'] = 'Reprovado'

        frequencia_data.append(item_dict)
    conn.close()

    if format == "pdf":
        return gerar_pdf_relatorio_frequencia(frequencia_data)
    elif format == "docx":
        return gerar_docx_relatorio_frequencia(frequencia_data)
    else:
        flash("Formato de download inválido!", "erro")
        return redirect(url_for("relatorios_frequencia"))


def gerar_pdf_relatorio_frequencia(data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph("Relatório de Frequência de Alunos", styles['h1']))
    elements.append(Spacer(1, 0.2 * inch))

    if data:
        headers = ["Aluno", "Disciplina", "Turma", "Faixa Etária", "Aulas Reg.", "Presenças", "Freq. (%)", "Status"]
        table_data = [headers]

        for item in data:
            table_data.append([
                item['aluno_nome'],
                item['disciplina_nome'],
                item['turma_nome'] if item['turma_nome'] else 'N/A',
                item['faixa_etaria'].replace('_', ' ').title() if item['faixa_etaria'] else 'N/A',
                item['total_aulas_registradas'],
                item['total_presencas'],
                f"{item['frequencia_porcentagem']:.1f}%",
                item['status_frequencia']
            ])

        table = Table(table_data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Aluno left-aligned
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(table)
    else:
        elements.append(Paragraph("Nenhum relatório de frequência encontrado com os filtros aplicados.", styles['Normal']))

    doc.build(elements)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.pdf", mimetype="application/pdf")


def gerar_docx_relatorio_frequencia(data):
    document = Document()
    document.add_heading('Relatório de Frequência de Alunos', level=1)

    if data:
        table = document.add_table(rows=1, cols=8)
        table.style = 'Table Grid'

        hdr_cells = table.rows[0].cells
        headers = ["Aluno", "Disciplina", "Turma", "Faixa Etária", "Aulas Reg.", "Presenças", "Freq. (%)", "Status"]
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

        for item in data:
            row_cells = table.add_row().cells
            row_cells[0].text = item['aluno_nome']
            row_cells[1].text = item['disciplina_nome']
            row_cells[2].text = item['turma_nome'] if item['turma_nome'] else 'N/A'
            row_cells[3].text = item['faixa_etaria'].replace('_', ' ').title() if item['faixa_etaria'] else 'N/A'
            row_cells[4].text = str(item['total_aulas_registradas'])
            row_cells[5].text = str(item['total_presencas'])
            row_cells[6].text = f"{item['frequencia_porcentagem']:.1f}%"
            row_cells[7].text = item['status_frequencia']

            for i in range(1, 8):
                row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT

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