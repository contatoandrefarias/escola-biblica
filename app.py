import os
from datetime import date, datetime
from flask import (Flask, render_template, request,
                   redirect, url_for, flash, send_file)
from flask_login import (LoginManager, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco, DATABASE # Importar DATABASE
import sqlite3
import shutil # Para copiar arquivos
from functools import wraps # Para criar o decorador de admin

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
    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
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
            SELECT a.id, a.nome, a.data_nascimento, a.email, a.telefone, a.membro_igreja,
                   t.nome AS turma_nome, t.faixa_etaria AS turma_faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno_raw = cursor.fetchone()

        if aluno_raw:
            aluno = dict(aluno_raw) # Converter para dict mutável

            cursor.execute("""
                SELECT m.id AS matricula_id,
                       d.id AS disciplina_id,
                       d.nome AS disciplina_nome,
                       d.tem_atividades,
                       d.frequencia_minima,
                       m.data_inicio, m.data_conclusao, m.status,
                       m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                       m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                       t.faixa_etaria AS turma_faixa_etaria -- Adicionado para consistência
                FROM matriculas m
                JOIN disciplinas d ON m.disciplina_id = d.id
                LEFT JOIN alunos a ON m.aluno_id = a.id
                LEFT JOIN turmas t ON a.turma_id = t.id
                WHERE m.aluno_id = ?
                ORDER BY d.nome
            """, (id,))
            raw_matriculas = cursor.fetchall()

            for mat in raw_matriculas:
                matricula_dict = dict(mat) # Converter para dict mutável

                # --- Cálculo de Frequência ---
                cursor.execute("""
                    SELECT data_aula, presente, fez_atividade
                    FROM presencas
                    WHERE matricula_id = ?
                    ORDER BY data_aula DESC
                """, (matricula_dict['matricula_id'],))
                historico_chamadas_raw = cursor.fetchall()
                matricula_dict['historico_chamadas'] = [dict(c) for c in historico_chamadas_raw] # Converter para lista de dicts

                presencas = sum(1 for c in historico_chamadas_raw if c['presente'])
                total_aulas = len(historico_chamadas_raw)
                atividades_feitas = sum(1 for c in historico_chamadas_raw if c['fez_atividade'])

                matricula_dict['presencas'] = presencas
                matricula_dict['total_aulas'] = total_aulas
                matricula_dict['atividades_feitas'] = atividades_feitas

                frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
                matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

                # --- Cálculo de Notas e Status ---
                faixa_etaria = matricula_dict.get('turma_faixa_etaria', 'adultos') # Usar get com default

                nota_final_calc = None
                status_display = matricula_dict['status'] # Default para o status do BD

                if faixa_etaria.startswith('criancas'):
                    matricula_dict['media_display'] = 'N/A'
                    # Status para crianças baseado apenas na frequência
                    if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                        matricula_dict['status_frequencia'] = 'Aprovado'
                    else:
                        matricula_dict['status_frequencia'] = 'Reprovado'
                    status_display = matricula_dict['status_frequencia'] # Crianças usam status da frequência
                elif faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens'):
                    meditacao_val = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                    versiculos_val = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                    desafio_nota_val = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                    visitante_val = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
                    nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
                    matricula_dict['media_display'] = f"{nota_final_calc:.1f}"
                    # Lógica de status combinada para Adolescentes/Jovens
                    if matricula_dict['status'] == 'cursando':
                        if nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                            status_display = 'Aprovado (Provisório)'
                        elif nota_final_calc < 7.0 and frequencia_porcentagem < matricula_dict['frequencia_minima']:
                            status_display = 'Reprovado (Provisório)'
                        elif nota_final_calc < 7.0:
                            status_display = 'Reprovado (Provisório - Notas)'
                        elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                            status_display = 'Reprovado (Provisório - Frequência)'
                        else:
                            status_display = 'Cursando' # Caso não se encaixe nas condições acima
                    else: # Aprovado/Reprovado final
                        status_display = matricula_dict['status'].capitalize()

                else: # Adultos
                    nota1_val = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                    nota2_val = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
                    if nota1_val > 0 or nota2_val > 0:
                        nota_final_calc = (nota1_val + nota2_val) / 2
                        matricula_dict['media_display'] = f"{nota_final_calc:.1f}"
                    else:
                        matricula_dict['media_display'] = '—'
                    # Lógica de status combinada para Adultos
                    if matricula_dict['status'] == 'cursando':
                        if nota_final_calc is not None and nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                            status_display = 'Aprovado (Provisório)'
                        elif (nota_final_calc is not None and nota_final_calc < 7.0) and frequencia_porcentagem < matricula_dict['frequencia_minima']:
                            status_display = 'Reprovado (Provisório)'
                        elif nota_final_calc is not None and nota_final_calc < 7.0:
                            status_display = 'Reprovado (Provisório - Notas)'
                        elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                            status_display = 'Reprovado (Provisório - Frequência)'
                        else:
                            status_display = 'Cursando' # Caso não se encaixe nas condições acima
                    else: # Aprovado/Reprovado final
                        status_display = matricula_dict['status'].capitalize()

                matricula_dict['status_display'] = status_display
                matriculas.append(matricula_dict)

    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        print(f"ERRO na trilha do aluno: {e}") # Logar o erro no console
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
        nome              = request.form.get("nome", "").strip()
        descricao         = request.form.get("descricao", "").strip()
        tem_atividades    = 1 if request.form.get("tem_atividades") else 0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0
        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("nova_disciplina"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO disciplinas (nome,descricao,tem_atividades,frequencia_minima)
                VALUES (?,?,?,?)
            """, (nome, descricao, tem_atividades, frequencia_minima))
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
    return render_template("nova_disciplina.html")


@app.route("/disciplinas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome              = request.form.get("nome", "").strip()
        descricao         = request.form.get("descricao", "").strip()
        tem_atividades    = 1 if request.form.get("tem_atividades") else 0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0
        ativa             = 1 if request.form.get("ativa") else 0
        if not nome:
            flash("Nome é obrigatório!", "erro")
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disciplina = cursor.fetchone()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=disciplina)
        try:
            cursor.execute("""
                UPDATE disciplinas
                SET nome=?,descricao=?,tem_atividades=?,frequencia_minima=?,ativa=?
                WHERE id=?
            """, (nome, descricao, tem_atividades, frequencia_minima, ativa, id))
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
    conn.close()
    if not disciplina:
        flash("Disciplina não encontrada!", "erro")
        return redirect(url_for("disciplinas"))
    return render_template("editar_disciplina.html", disciplina=disciplina)


@app.route("/disciplinas/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Verificar se há matrículas ativas para esta disciplina
        cursor.execute("SELECT COUNT(*) FROM matriculas WHERE disciplina_id = ?", (id,))
        num_matriculas = cursor.fetchone()[0]
        if num_matriculas > 0:
            flash(f"Não é possível excluir a disciplina. Existem {num_matriculas} matrículas associadas a ela.", "erro")
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
        cursor.execute("DELETE FROM professores WHERE id=?", (id,))
        conn.commit()
        flash("Professor excluído!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir professor: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("professores"))


# ══════════════════════════════════════
# MATRÍCULAS
# ══════════════════════════════════════
@app.route("/matriculas")
@login_required
def matriculas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.id, a.nome AS aluno_nome, d.nome AS disciplina_nome,
               t.nome AS turma_nome, t.faixa_etaria,
               m.data_inicio, m.data_conclusao, m.status
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
        aluno_id       = request.form.get("aluno_id", type=int)
        disciplina_id  = request.form.get("disciplina_id", type=int)
        data_inicio    = request.form.get("data_inicio", "").strip()
        data_conclusao = request.form.get("data_conclusao", "").strip() or None
        status         = request.form.get("status", "cursando").strip()
        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Aluno, Disciplina e Data de Início são obrigatórios!", "erro")
            return redirect(url_for("nova_matricula"))
        try:
            cursor.execute("""
                INSERT INTO matriculas (aluno_id,disciplina_id,data_inicio,data_conclusao,status)
                VALUES (?,?,?,?,?)
            """, (aluno_id, disciplina_id, data_inicio, data_conclusao, status))
            conn.commit()
            flash("Matrícula criada com sucesso!", "sucesso")
            _atualizar_status_matricula(cursor.lastrowid) # Atualiza o status da nova matrícula
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
    return render_template("nova_matricula.html",
                           alunos=alunos_disponiveis,
                           disciplinas=disciplinas_disponiveis,
                           now=date.today())


@app.route("/matriculas/novo_aluno_disciplina", methods=["GET", "POST"])
@login_required
def novo_aluno_disciplina():
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        aluno_id       = request.form.get("aluno_id", type=int)
        disciplina_id  = request.form.get("disciplina_id", type=int)
        data_inicio    = request.form.get("data_inicio", "").strip()
        data_conclusao = request.form.get("data_conclusao", "").strip() or None
        status         = request.form.get("status", "cursando").strip()
        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Aluno, Disciplina e Data de Início são obrigatórios!", "erro")
            return redirect(url_for("novo_aluno_disciplina"))
        try:
            cursor.execute("""
                INSERT INTO matriculas (aluno_id,disciplina_id,data_inicio,data_conclusao,status)
                VALUES (?,?,?,?,?)
            """, (aluno_id, disciplina_id, data_inicio, data_conclusao, status))
            conn.commit()
            flash("Matrícula criada com sucesso!", "sucesso")
            _atualizar_status_matricula(cursor.lastrowid) # Atualiza o status da nova matrícula
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
            update_query = """
                UPDATE matriculas
                SET data_conclusao=?, status=?,
                    nota1=NULL, nota2=NULL, participacao=NULL, desafio=NULL, prova=NULL,
                    meditacao=NULL, versiculos=NULL, desafio_nota=NULL, visitante=NULL
                WHERE id=?
            """
            cursor.execute(update_query, (data_conclusao, status, id))

            if faixa_etaria_matricula.startswith('adolescentes') or faixa_etaria_matricula.startswith('jovens'):
                cursor.execute("""
                    UPDATE matriculas
                    SET meditacao=?, versiculos=?, desafio_nota=?, visitante=?
                    WHERE id=?
                """, (meditacao_aj, versiculos_aj, desafio_nota_aj, visitante_aj, id))
            elif faixa_etaria_matricula == 'adultos':
                cursor.execute("""
                    UPDATE matriculas
                    SET nota1=?, nota2=?, participacao=?, desafio=?, prova=?
                    WHERE id=?
                """, (nota1_adulto, nota2_adulto, participacao_adulto, desafio_adulto, prova_adulto, id))
            # Crianças não têm notas para atualizar aqui

            conn.commit()
            flash("Matrícula atualizada!", "sucesso")
            _atualizar_status_matricula(id) # Recalcula o status após a atualização
        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    cursor.execute("""
        SELECT m.*, a.nome AS aluno_nome, d.nome AS disciplina_nome,
               t.faixa_etaria AS turma_faixa_etaria, d.tem_atividades
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
        # Excluir registros de presença associados a esta matrícula
        cursor.execute("DELETE FROM presencas WHERE matricula_id=?", (id,))
        # Excluir a matrícula
        cursor.execute("DELETE FROM matriculas WHERE id=?", (id,))
        conn.commit()
        flash("Matrícula e registros de presença excluídos!", "sucesso")
    except Exception as e:
        flash(f"Erro ao excluir matrícula: {e}", "erro")
    finally:
        conn.close()
    return redirect(url_for("matriculas"))


# ══════════════════════════════════════
# PRESENÇA
# ══════════════════════════════════════
@app.route("/presenca/chamada", methods=["GET", "POST"])
@login_required
def chamada():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("SELECT id, nome, tem_atividades FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_ativas = cursor.fetchall()

    alunos_chamada = []
    selected_disciplina = None
    selected_data_aula = None
    tem_atividades = False

    if request.method == "GET":
        disciplina_id_str = request.args.get("disciplina_id")
        data_aula_str = request.args.get("data_aula")

        if disciplina_id_str and data_aula_str:
            disciplina_id = int(disciplina_id_str)
            selected_data_aula = data_aula_str
            try:
                data_aula = datetime.strptime(data_aula_str, "%Y-%m-%d").date()
            except ValueError:
                flash("Formato de data inválido.", "erro")
                return redirect(url_for("chamada"))

            cursor.execute("SELECT id, nome, tem_atividades FROM disciplinas WHERE id = ?", (disciplina_id,))
            selected_disciplina_raw = cursor.fetchone()
            if selected_disciplina_raw:
                selected_disciplina = dict(selected_disciplina_raw)
                tem_atividades = selected_disciplina['tem_atividades'] == 1

                cursor.execute("""
                    SELECT m.id AS matricula_id,
                           a.nome AS aluno_nome,
                           p.presente,
                           p.fez_atividade
                    FROM matriculas m
                    JOIN alunos a ON m.aluno_id = a.id
                    LEFT JOIN presencas p ON p.matricula_id = m.id AND p.data_aula = ?
                    WHERE m.disciplina_id = ? AND m.status = 'cursando'
                    ORDER BY a.nome
                """, (data_aula.isoformat(), disciplina_id))
                alunos_chamada = [dict(row) for row in cursor.fetchall()] # Converter para lista de dicts
            else:
                flash("Disciplina selecionada não encontrada ou inativa.", "erro")

    elif request.method == "POST":
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_aula_str = request.form.get("data_aula")

        if not disciplina_id or not data_aula_str:
            flash("Disciplina e Data da Aula são obrigatórias!", "erro")
            return redirect(url_for("chamada"))

        try:
            data_aula = datetime.strptime(data_aula_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Formato de data inválido.", "erro")
            return redirect(url_for("chamada"))

        cursor.execute("SELECT tem_atividades FROM disciplinas WHERE id = ?", (disciplina_id,))
        tem_atividades_disciplina = cursor.fetchone()['tem_atividades'] == 1

        # Excluir registros de presença existentes para esta disciplina e data
        cursor.execute(
            "DELETE FROM presencas WHERE matricula_id IN (SELECT id FROM matriculas WHERE disciplina_id = ?) AND data_aula = ?",
            (disciplina_id, data_aula.isoformat())
        )

        # Inserir novos registros de presença
        matriculas_afetadas = []
        for key, value in request.form.items():
            if key.startswith("presente_"):
                matricula_id = int(key.split("_")[1])
                presente = 1 if value == "on" else 0
                fez_atividade = 1 if tem_atividades_disciplina and request.form.get(f"atividade_{matricula_id}") == "on" else 0

                cursor.execute("""
                    INSERT INTO presencas (matricula_id, data_aula, presente, fez_atividade)
                    VALUES (?, ?, ?, ?)
                """, (matricula_id, data_aula.isoformat(), presente, fez_atividade))
                matriculas_afetadas.append(matricula_id)

        conn.commit()

        # Atualizar status das matrículas afetadas
        for mat_id in matriculas_afetadas:
            _atualizar_status_matricula(mat_id)

        flash("Chamada salva com sucesso!", "sucesso")
        return redirect(url_for("chamada", disciplina_id=disciplina_id, data_aula=data_aula_str))

    conn.close()
    return render_template("chamada.html",
                           disciplinas=disciplinas_ativas,
                           alunos_chamada=alunos_chamada,
                           selected_disciplina=selected_disciplina,
                           selected_data_aula=selected_data_aula,
                           tem_atividades=tem_atividades,
                           today=date.today())


# ══════════════════════════════════════
# RELATÓRIOS
# ══════════════════════════════════════
def get_relatorio_data(disciplina_id=None, turma_id=None, aluno_id=None, status=None):
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

        # --- Cálculo de Notas e Status ---
        faixa_etaria = matricula_dict.get('turma_faixa_etaria', 'adultos') # Usar get com default

        nota_final_calc = None
        status_display = matricula_dict['status'] # Default para o status do BD

        if faixa_etaria.startswith('criancas'):
            matricula_dict['media_display'] = 'N/A'
            # Status para crianças baseado apenas na frequência
            if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                matricula_dict['status_frequencia'] = 'Aprovado'
            else:
                matricula_dict['status_frequencia'] = 'Reprovado'
            status_display = matricula_dict['status_frequencia'] # Crianças usam status da frequência
        elif faixa_etaria.startswith('adolescentes') or faixa_etaria.startswith('jovens'):
            meditacao_val = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
            versiculos_val = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
            desafio_nota_val = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
            visitante_val = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
            nota_final_calc = meditacao_val + versiculos_val + desafio_nota_val + visitante_val
            matricula_dict['media_display'] = f"{nota_final_calc:.1f}"
            # Lógica de status combinada para Adolescentes/Jovens
            if matricula_dict['status'] == 'cursando':
                if nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                # Se aprovado por nota e frequência, mas ainda cursando, é provisório
                    status_display = 'Aprovado (Provisório)'
                elif nota_final_calc < 7.0 and frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status_display = 'Reprovado (Provisório)'
                elif nota_final_calc < 7.0:
                    status_display = 'Reprovado (Provisório - Notas)'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status_display = 'Reprovado (Provisório - Frequência)'
                else:
                    status_display = 'Cursando' # Caso não se encaixe nas condições acima
            else: # Aprovado/Reprovado final
                status_display = matricula_dict['status'].capitalize()

        else: # Adultos
            nota1_val = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
            nota2_val = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
            if nota1_val > 0 or nota2_val > 0:
                nota_final_calc = (nota1_val + nota2_val) / 2
                matricula_dict['media_display'] = f"{nota_final_calc:.1f}"
            else:
                matricula_dict['media_display'] = '—'
            # Lógica de status combinada para Adultos
            if matricula_dict['status'] == 'cursando':
                if nota_final_calc is not None and nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status_display = 'Aprovado (Provisório)'
                elif (nota_final_calc is not None and nota_final_calc < 7.0) and frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status_display = 'Reprovado (Provisório)'
                elif nota_final_calc is not None and nota_final_calc < 7.0:
                    status_display = 'Reprovado (Provisório - Notas)'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status_display = 'Reprovado (Provisório - Frequência)'
                elif frequencia_porcentagem < matricula_dict['frequencia_minima']:
                    status_display = 'Reprovado (Provisório - Frequência)'
                else:
                    status_display = 'Cursando' # Caso não se encaixe nas condições acima
            else: # Aprovado/Reprovado final
                status_display = matricula_dict['status'].capitalize()

        matricula_dict['status_display'] = status_display
        processed_matriculas.append(matricula_dict)

    conn.close()
    return processed_matriculas


@app.route("/relatorios")
@login_required
def relatorios():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas = cursor.fetchall()
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos = cursor.fetchall()
    conn.close()

    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    selected_turma_id = request.args.get("turma_id", type=int)
    selected_aluno_id = request.args.get("aluno_id", type=int)
    selected_status = request.args.get("status")

    relatorio_data = get_relatorio_data(
        disciplina_id=selected_disciplina_id,
        turma_id=selected_turma_id,
        aluno_id=selected_aluno_id,
        status=selected_status
    )

    return render_template("relatorios.html",
                           disciplinas=disciplinas,
                           turmas=turmas,
                           alunos=alunos,
                           relatorio_data=relatorio_data,
                           selected_disciplina_id=selected_disciplina_id,
                           selected_turma_id=selected_turma_id,
                           selected_aluno_id=selected_aluno_id,
                           selected_status=selected_status)


@app.route("/relatorios/download/<format>")
@login_required
def download_relatorio(format):
    disciplina_id = request.args.get("disciplina_id", type=int)
    turma_id = request.args.get("turma_id", type=int)
    aluno_id = request.args.get("aluno_id", type=int)
    status = request.args.get("status")

    relatorio_data = get_relatorio_data(
        disciplina_id=disciplina_id,
        turma_id=turma_id,
        aluno_id=aluno_id,
        status=status
    )

    if not relatorio_data:
        flash("Nenhum dado para gerar o relatório.", "erro")
        return redirect(url_for("relatorios"))

    if format == "pdf":
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph("Relatório de Matrículas", styles['h2']))
        elements.append(Spacer(1, 0.2 * inch))

        # Cabeçalho da tabela
        data = [
            ["Aluno", "Disciplina", "Turma", "Início", "Conclusão", "Média Final", "Frequência (%)", "Status"]
        ]
        for item in relatorio_data:
            data.append([
                item['aluno_nome'],
                item['disciplina_nome'],
                f"{item['turma_nome']} ({item['turma_faixa_etaria'].replace('_', ' ').title()})" if item['turma_nome'] else 'N/A',
                item['data_inicio'].replace('-', '/'),
                item['data_conclusao'].replace('-', '/') if item['data_conclusao'] else '—',
                item['media_display'],
                f"{item['frequencia_porcentagem']:.1f}%",
                item['status_display']
            ])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')), # Dark header
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'), # Aluno à esquerda
            ('ALIGN', (1, 0), (1, -1), 'LEFT'), # Disciplina à esquerda
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')), # Light row background
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        elements.append(table)
        doc.build(buffer)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.pdf", mimetype="application/pdf")

    elif format == "docx":
        document = Document()
        document.add_heading('Relatório de Matrículas', level=1)

        table = document.add_table(rows=1, cols=8)
        table.autofit = True
        # Set column widths manually for better control
        table.columns[0].width = Inches(1.2) # Aluno
        table.columns[1].width = Inches(1.2) # Disciplina
        table.columns[2].width = Inches(1.0) # Turma
        table.columns[3].width = Inches(0.8) # Início
        table.columns[4].width = Inches(0.9) # Conclusão
        table.columns[5].width = Inches(0.8) # Média Final
        table.columns[6].width = Inches(1.0) # Frequência (%)
        table.columns[7].width = Inches(1.2) # Status

        hdr_cells = table.rows[0].cells
        headers = ["Aluno", "Disciplina", "Turma", "Início", "Conclusão", "Média Final", "Frequência (%)", "Status"]
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            # Make header bold and centered
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Set header background color (requires python-docx-oss)
            # from docx.oxml.ns import qn
            # from docx.oxml import OxmlElement
            # shd = OxmlElement('w:shd')
            # shd.set(qn('w:fill'), '212529') # Hex color for dark gray
            # hdr_cells[i]._tc.get_or_add_tcPr().append(shd)


        for item in relatorio_data:
            row_cells = table.add_row().cells
            row_cells[0].text = item['aluno_nome']
            row_cells[1].text = item['disciplina_nome']
            row_cells[2].text = f"{item['turma_nome']} ({item['turma_faixa_etaria'].replace('_', ' ').title()})" if item['turma_nome'] else 'N/A'
            row_cells[3].text = item['data_inicio'].replace('-', '/')
            row_cells[4].text = item['data_conclusao'].replace('-', '/') if item['data_conclusao'] else '—'
            row_cells[5].text = item['media_display']
            row_cells[6].text = f"{item['frequencia_porcentagem']:.1f}%"
            row_cells[7].text = item['status_display']
            # Center all cells except first two
            for i in range(2, 8):
                for paragraph in row_cells[i].paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        buffer = BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="relatorio_matriculas.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    flash("Formato de download inválido.", "erro")
    return redirect(url_for("relatorios"))


@app.route("/relatorios/frequencia")
@login_required
def relatorios_frequencia():
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas = cursor.fetchall()
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas = cursor.fetchall()
    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos = cursor.fetchall()
    conn.close()

    selected_disciplina_id = request.args.get("disciplina_id", type=int)
    selected_turma_id = request.args.get("turma_id", type=int)
    selected_aluno_id = request.args.get("aluno_id", type=int)
    data_inicio_str = request.args.get("data_inicio")
    data_fim_str = request.args.get("data_fim")

    frequencia_data = []
    if selected_disciplina_id or selected_turma_id or selected_aluno_id or (data_inicio_str and data_fim_str):
        frequencia_data = get_frequencia_data(
            disciplina_id=selected_disciplina_id,
            turma_id=selected_turma_id,
            aluno_id=selected_aluno_id,
            data_inicio_str=data_inicio_str,
            data_fim_str=data_fim_str
        )

    return render_template("relatorio_frequencia.html",
                           disciplinas=disciplinas,
                           turmas=turmas,
                           alunos=alunos,
                           frequencia_data=frequencia_data,
                           selected_disciplina_id=selected_disciplina_id,
                           selected_turma_id=selected_turma_id,
                           selected_aluno_id=selected_aluno_id,
                           selected_data_inicio=data_inicio_str,
                           selected_data_fim=data_fim_str)


def get_frequencia_data(disciplina_id=None, turma_id=None, aluno_id=None, data_inicio_str=None, data_fim_str=None):
    conn = conectar()
    cursor = conn.cursor()

    query = """
        SELECT p.data_aula,
               a.nome AS aluno_nome,
               d.nome AS disciplina_nome,
               t.nome AS turma_nome,
               p.presente,
               p.fez_atividade,
               d.tem_atividades
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
        query += " AND t.id = ?"
        params.append(turma_id)
    if aluno_id:
        query += " AND a.id = ?"
        params.append(aluno_id)
    if data_inicio_str:
        query += " AND p.data_aula >= ?"
        params.append(data_inicio_str)
    if data_fim_str:
        query += " AND p.data_aula <= ?"
        params.append(data_fim_str)

    query += " ORDER BY p.data_aula DESC, a.nome"
    cursor.execute(query, params)
    raw_data = cursor.fetchall()
    conn.close()

    return [dict(row) for row in raw_data]


@app.route("/relatorios/frequencia/download/<format>")
@login_required
def download_relatorio_frequencia(format):
    disciplina_id = request.args.get("disciplina_id", type=int)
    turma_id = request.args.get("turma_id", type=int)
    aluno_id = request.args.get("aluno_id", type=int)
    data_inicio_str = request.args.get("data_inicio")
    data_fim_str = request.args.get("data_fim")

    frequencia_data = get_frequencia_data(
        disciplina_id=disciplina_id,
        turma_id=turma_id,
        aluno_id=aluno_id,
        data_inicio_str=data_inicio_str,
        data_fim_str=data_fim_str
    )

    if not frequencia_data:
        flash("Nenhum dado para gerar o relatório de frequência.", "erro")
        return redirect(url_for("relatorios_frequencia"))

    if format == "pdf":
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph("Relatório de Frequência", styles['h2']))
        elements.append(Spacer(1, 0.2 * inch))

        data = [
            ["Data da Aula", "Aluno", "Disciplina", "Turma", "Presente", "Atividade Feita"]
        ]
        for item in frequencia_data:
            data.append([
                item['data_aula'].replace('-', '/'),
                item['aluno_nome'],
                item['disciplina_nome'],
                item['turma_nome'] if item['turma_nome'] else 'N/A',
                "Sim" if item['presente'] else "Não",
                "Sim" if item['tem_atividades'] and item['fez_atividade'] else ("N/A" if not item['tem_atividades'] else "Não")
            ])

        table = Table(data)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#212529')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (1, -1), 'LEFT'), # Aluno e Disciplina à esquerda
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        elements.append(table)
        doc.build(buffer)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.pdf", mimetype="application/pdf")

    elif format == "docx":
        document = Document()
        document.add_heading('Relatório de Frequência', level=1)

        table = document.add_table(rows=1, cols=6)
        table.autofit = True
        table.columns[0].width = Inches(1.0) # Data
        table.columns[1].width = Inches(1.5) # Aluno
        table.columns[2].width = Inches(1.5) # Disciplina
        table.columns[3].width = Inches(1.0) # Turma
        table.columns[4].width = Inches(0.8) # Presente
        table.columns[5].width = Inches(1.2) # Atividade Feita

        hdr_cells = table.rows[0].cells
        headers = ["Data da Aula", "Aluno", "Disciplina", "Turma", "Presente", "Atividade Feita"]
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            for paragraph in hdr_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for item in frequencia_data:
            row_cells = table.add_row().cells
            row_cells[0].text = item['data_aula'].replace('-', '/')
            row_cells[1].text = item['aluno_nome']
            row_cells[2].text = item['disciplina_nome']
            row_cells[3].text = item['turma_nome'] if item['turma_nome'] else 'N/A'
            row_cells[4].text = "Sim" if item['presente'] else "Não"
            row_cells[5].text = "Sim" if item['tem_atividades'] and item['fez_atividade'] else ("N/A" if not item['tem_atividades'] else "Não")
            for i in range(4, 6): # Center "Presente" and "Atividade Feita"
                for paragraph in row_cells[i].paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        buffer = BytesIO()
        document.save(buffer)
        buffer.seek(0)
        return send_file(buffer, as_attachment=True, download_name="relatorio_frequencia.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    flash("Formato de download inválido.", "erro")
    return redirect(url_for("relatorios_frequencia"))


# ══════════════════════════════════════
# ADMIN - BACKUP E RESTAURAÇÃO
# ══════════════════════════════════════
@app.route("/admin/backup", methods=["GET", "POST"])
@login_required
@admin_required # Apenas administradores podem acessar
def backup_restauracao():
    if request.method == "POST":
        if 'backup_action' in request.form:
            # Ação de backup
            try:
                # O arquivo DATABASE é 'escola.db'
                return send_file(DATABASE, as_attachment=True, download_name=f"escola_backup_{date.today().isoformat()}.db", mimetype="application/x-sqlite3")
            except Exception as e:
                flash(f"Erro ao gerar backup: {e}", "erro")
                return redirect(url_for("backup_restauracao"))
        elif 'restore_file' in request.files:
            # Ação de restauração
            backup_file = request.files['restore_file']
            if backup_file.filename == '':
                flash("Nenhum arquivo selecionado para restauração.", "erro")
                return redirect(url_for("backup_restauracao"))
            if not backup_file.filename.endswith('.db'):
                flash("O arquivo de backup deve ter a extensão .db", "erro")
                return redirect(url_for("backup_restauracao"))

            try:
                # Fechar todas as conexões existentes com o banco de dados antes de substituir o arquivo
                # Isso é crucial para evitar erros de "database is locked"
                # No Flask, as conexões são gerenciadas por requisição, então pode ser necessário
                # garantir que nenhuma conexão esteja aberta no momento da substituição.
                # Uma forma simples é reiniciar o aplicativo após a restauração, mas isso não é ideal.
                # Para SQLite, a melhor prática é garantir que o arquivo não esteja em uso.
                # Como estamos em um ambiente de desenvolvimento/Railway, podemos tentar a substituição direta.

                # Criar um backup temporário do banco de dados atual antes de sobrescrever
                shutil.copy(DATABASE, f"{DATABASE}.pre_restore_backup")

                # Salvar o arquivo de backup enviado
                backup_file.save(DATABASE)

                # Re-inicializar o banco de dados para garantir que as novas tabelas/estrutura sejam carregadas
                # (embora o arquivo já esteja lá, isso pode ajudar a Flask a reconhecer a mudança)
                inicializar_banco() 

                flash("Banco de dados restaurado com sucesso! Recomenda-se reiniciar o servidor.", "sucesso")
            except Exception as e:
                # Se houver um erro na restauração, tentar restaurar o backup pré-restauração
                if os.path.exists(f"{DATABASE}.pre_restore_backup"):
                    shutil.copy(f"{DATABASE}.pre_restore_backup", DATABASE)
                    flash(f"Erro ao restaurar banco de dados: {e}. O backup anterior foi restaurado.", "erro")
                else:
                    flash(f"Erro ao restaurar banco de dados: {e}. Não foi possível restaurar o backup anterior.", "erro")
            finally:
                # Limpar o backup temporário
                if os.path.exists(f"{DATABASE}.pre_restore_backup"):
                    os.remove(f"{DATABASE}.pre_restore_backup")

                # No Railway, para que as mudanças no DB sejam refletidas, o contêiner precisaria ser reiniciado.
                # No ambiente local, você precisaria reiniciar o servidor Flask.
                # Uma mensagem para o usuário é o suficiente aqui.
                pass 
            return redirect(url_for("backup_restauracao"))

    return render_template("backup_restauracao.html")


# ══════════════════════════════════════
# USUÁRIOS
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
@admin_required # Apenas administradores podem acessar
def usuarios():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome, email, perfil FROM usuarios ORDER BY nome")
    lista = cursor.fetchall()
    conn.close()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
@admin_required # Apenas administradores podem acessar
def novo_usuario():
    if request.method == "POST":
        nome   = request.form.get("nome", "").strip()
        email  = request.form.get("email", "").strip()
        senha  = request.form.get("senha", "")
        perfil = request.form.get("perfil", "usuario")
        if not nome or not email or not senha:
            flash("Nome, E-mail e Senha são obrigatórios!", "erro")
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
@admin_required # Apenas administradores podem acessar
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
@admin_required # Apenas administradores podem acessar
def excluir_usuario(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Impedir que o próprio admin logado se exclua
        if current_user.id == id:
            flash("Você não pode excluir sua própria conta de administrador!", "erro")
            return redirect(url_for("usuarios"))

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