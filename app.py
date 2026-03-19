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
        turma_id      = request.form.get("turma_id", type=int) or None # Pode ser None se não selecionado

        if not nome:
            flash("Nome é obrigatório!", "erro")
            # Recarregar turmas para o template em caso de erro
            cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_disponiveis = cursor.fetchall()
            conn.close()
            return render_template("novo_aluno.html", turmas=turmas_disponiveis)

        try:
            cursor.execute("""
                INSERT INTO alunos (nome, telefone, email, data_nascimento, membro_igreja, turma_id)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (nome, telefone, email, data_nascimento, membro_igreja, turma_id))
            conn.commit()
            flash(f"Aluno '{nome}' cadastrado!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "alunos.email" in str(e):
                flash("Já existe um aluno com este e-mail!", "erro")
            else:
                flash(f"Erro de integridade ao cadastrar aluno: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao cadastrar aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("alunos"))

    # Para requisições GET ou em caso de erro no POST
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
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
        data_nascimento = request.form.get("data_nascimento", "").strip() or None
        membro_igreja = 1 if request.form.get("membro_igreja") else 0
        turma_id      = request.form.get("turma_id", type=int) or None
        if not nome:
            flash("Nome é obrigatório!", "erro")
            # Recarregar dados para o template em caso de erro
            cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
            aluno = cursor.fetchone()
            cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
            turmas_disponiveis = cursor.fetchall()
            conn.close()
            return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_disponiveis)
        try:
            cursor.execute("""
                UPDATE alunos
                SET nome=?, telefone=?, email=?, data_nascimento=?, membro_igreja=?, turma_id=?
                WHERE id=?
            """, (nome, telefone, email, data_nascimento, membro_igreja, turma_id, id))
            conn.commit()
            flash("Aluno atualizado!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "alunos.email" in str(e):
                flash("Este e-mail já está cadastrado para outro aluno!", "erro")
            else:
                flash(f"Erro de integridade ao atualizar aluno: {e}", "erro")
        except Exception as e:
            flash(f"Erro inesperado ao atualizar aluno: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("alunos"))

    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    cursor.execute("SELECT id, nome, faixa_etaria FROM turmas WHERE ativa=1 ORDER BY nome")
    turmas_disponiveis = cursor.fetchall()
    conn.close()
    if not aluno:
        flash("Aluno não encontrado!", "erro")
        return redirect(url_for("alunos"))
    return render_template("editar_aluno.html", aluno=aluno, turmas=turmas_disponiveis)


@app.route("/alunos/<int:id>/excluir", methods=["POST"])
@login_required
def excluir_aluno(id):
    conn   = conectar()
    cursor = conn.cursor()
    try:
        # Excluir matrículas do aluno primeiro
        cursor.execute("DELETE FROM presencas WHERE matricula_id IN (SELECT id FROM matriculas WHERE aluno_id=?)", (id,))
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
            SELECT a.*, t.nome AS turma_nome, t.faixa_etaria
            FROM alunos a
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE a.id = ?
        """, (id,))
        aluno_raw = cursor.fetchone()

        if aluno_raw:
            aluno = dict(aluno_raw) # Converter para dict mutável

            cursor.execute("""
                SELECT m.id AS matricula_id,
                       d.nome AS disciplina_nome,
                       d.tem_atividades,
                       d.frequencia_minima,
                       m.data_inicio, m.data_conclusao, m.status,
                       m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                       m.meditacao, m.versiculos, m.desafio_nota, m.visitante
                FROM matriculas m
                JOIN disciplinas d ON m.disciplina_id = d.id
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
                matricula_dict['historico_chamadas'] = [dict(c) for c in historico_chamadas_raw] # Converter para dict

                presencas = sum(1 for c in historico_chamadas_raw if c['presente'])
                total_aulas = len(historico_chamadas_raw)
                atividades_feitas = sum(1 for c in historico_chamadas_raw if c['fez_atividade'])

                matricula_dict['presencas'] = presencas
                matricula_dict['total_aulas'] = total_aulas
                matricula_dict['atividades_feitas'] = atividades_feitas

                frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
                matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

                # --- Cálculo de Notas e Status ---
                faixa_etaria = aluno.get('faixa_etaria', 'adultos') # Usar a faixa etaria do aluno

                nota_final_calc = None
                status_frequencia = None
                status_notas = None
                status_display = matricula_dict['status'] # Default

                if faixa_etaria in ['criancas_0_3', 'criancas_4_7', 'criancas_8_12']:
                    # Crianças: Apenas frequência
                    nota_final_calc = None # Não há nota
                    matricula_dict['media_display'] = 'N/A'
                    if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                        status_frequencia = 'aprovado'
                    else:
                        status_frequencia = 'reprovado'
                    status_display = f"Frequência: {status_frequencia.capitalize()}"

                elif faixa_etaria in ['adolescentes_13_15', 'jovens_16_17']:
                    # Adolescentes/Jovens: Soma dos componentes
                    meditacao = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
                    versiculos = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
                    desafio_nota = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
                    visitante = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
                    nota_final_calc = meditacao + versiculos + desafio_nota + visitante
                    matricula_dict['media_display'] = f"{nota_final_calc:.1f}"

                    # Atualizar nota1 com a soma para consistência
                    matricula_dict['nota1'] = nota_final_calc

                    # Lógica de status para Adolescentes/Jovens
                    if nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                        status_notas = 'aprovado'
                    elif nota_final_calc >= 5.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                        status_notas = 'recuperacao'
                    else:
                        status_notas = 'reprovado'

                    if status_notas == 'aprovado' and matricula_dict['status'] == 'cursando':
                        status_display = 'Aprovado (Provisório)'
                    elif status_notas == 'recuperacao' and matricula_dict['status'] == 'cursando':
                        status_display = 'Recuperação (Provisório)'
                    elif status_notas == 'reprovado' and matricula_dict['status'] == 'cursando':
                        status_display = 'Reprovado (Provisório)'
                    else:
                        status_display = matricula_dict['status'].capitalize()


                else: # Adultos
                    n1 = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
                    n2 = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
                    nota_final_calc = (n1 + n2) / 2 if (n1 is not None and n2 is not None) else None
                    matricula_dict['media_display'] = f"{nota_final_calc:.1f}" if nota_final_calc is not None else '—'

                    # Lógica de status para Adultos
                    if nota_final_calc is not None:
                        if nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                            status_notas = 'aprovado'
                        elif nota_final_calc >= 5.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                            status_notas = 'recuperacao'
                        else:
                            status_notas = 'reprovado'
                    else:
                        status_notas = 'cursando' # Sem notas, assume cursando

                    if status_notas == 'aprovado' and matricula_dict['status'] == 'cursando':
                        status_display = 'Aprovado (Provisório)'
                    elif status_notas == 'recuperacao' and matricula_dict['status'] == 'cursando':
                        status_display = 'Recuperação (Provisório)'
                    elif status_notas == 'reprovado' and matricula_dict['status'] == 'cursando':
                        status_display = 'Reprovado (Provisório)'
                    else:
                        status_display = matricula_dict['status'].capitalize()

                matricula_dict['nota_final_calc'] = nota_final_calc
                matricula_dict['status_frequencia'] = status_frequencia
                matricula_dict['status_notas'] = status_notas
                matricula_dict['status_display'] = status_display

                matriculas.append(matricula_dict)

    except Exception as e:
        flash(f"Erro ao carregar trilha do aluno: {e}", "erro")
        app.logger.error(f"Erro na trilha do aluno {id}: {e}", exc_info=True)
        return redirect(url_for('alunos'))
    finally:
        conn.close()

    return render_template("trilha_aluno.html", aluno=aluno, matriculas=matriculas)


# ══════════════════════════════════════
# MATRÍCULAS
# ══════════════════════════════════════
def _atualizar_status_matricula(matricula_id):
    conn = conectar()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT m.aluno_id, m.disciplina_id, m.data_inicio, m.data_conclusao,
                   m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
                   m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
                   d.tem_atividades, d.frequencia_minima,
                   t.faixa_etaria
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            JOIN alunos a ON m.aluno_id = a.id
            LEFT JOIN turmas t ON a.turma_id = t.id
            WHERE m.id = ?
        """, (matricula_id,))
        matricula_data = cursor.fetchone()

        if not matricula_data:
            return

        # Converter para dict mutável
        matricula_data = dict(matricula_data)

        # --- Cálculo de Frequência ---
        cursor.execute("""
            SELECT presente, fez_atividade
            FROM presencas
            WHERE matricula_id = ?
        """, (matricula_id,))
        historico_chamadas = cursor.fetchall()

        presencas = sum(1 for c in historico_chamadas if c['presente'])
        total_aulas = len(historico_chamadas)
        frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0

        # --- Determinar Faixa Etária ---
        faixa_etaria = matricula_data.get('faixa_etaria', 'adultos') # Default para adultos

        novo_status = "cursando" # Default
        nota_final_calculada = None

        if faixa_etaria in ['criancas_0_3', 'criancas_4_7', 'criancas_8_12']:
            # Crianças: Apenas frequência
            if frequencia_porcentagem >= matricula_data['frequencia_minima']:
                novo_status = 'aprovado'
            else:
                novo_status = 'reprovado'
            nota_final_calculada = None # Não há nota para crianças

        elif faixa_etaria in ['adolescentes_13_15', 'jovens_16_17']:
            # Adolescentes/Jovens: Soma dos componentes
            meditacao = matricula_data['meditacao'] if matricula_data['meditacao'] is not None else 0
            versiculos = matricula_data['versiculos'] if matricula_data['versiculos'] is not None else 0
            desafio_nota = matricula_data['desafio_nota'] if matricula_data['desafio_nota'] is not None else 0
            visitante = matricula_data['visitante'] if matricula_data['visitante'] is not None else 0
            nota_final_calculada = meditacao + versiculos + desafio_nota + visitante

            if nota_final_calculada >= 7.0 and frequencia_porcentagem >= matricula_data['frequencia_minima']:
                novo_status = 'aprovado'
            elif nota_final_calculada >= 5.0 and frequencia_porcentagem >= matricula_data['frequencia_minima']:
                novo_status = 'recuperacao'
            else:
                novo_status = 'reprovado'

            # Atualizar nota1 com a soma para consistência
            matricula_data['nota1'] = nota_final_calculada

        else: # Adultos
            n1 = matricula_data['nota1'] if matricula_data['nota1'] is not None else 0
            n2 = matricula_data['nota2'] if matricula_data['nota2'] is not None else 0

            if n1 is not None and n2 is not None:
                nota_final_calculada = (n1 + n2) / 2
            else:
                nota_final_calculada = None # Não é possível calcular a média se uma das notas for None

            if nota_final_calculada is not None:
                if nota_final_calculada >= 7.0 and frequencia_porcentagem >= matricula_data['frequencia_minima']:
                    novo_status = 'aprovado'
                elif nota_final_calculada >= 5.0 and frequencia_porcentagem >= matricula_data['frequencia_minima']:
                    novo_status = 'recuperacao'
                else:
                    novo_status = 'reprovado'
            else:
                novo_status = 'cursando' # Se não há notas, permanece cursando

        # Se a data de conclusão foi definida e já passou, o status final é o calculado
        if matricula_data['data_conclusao'] and date.fromisoformat(matricula_data['data_conclusao']) <= date.today():
            final_status = novo_status
        else:
            # Se ainda está cursando, o status é 'cursando' ou o status provisório
            final_status = matricula_data['status'] # Mantém o status atual do banco se não for para concluir

        # Atualizar o status e a nota1 (para Adolescentes/Jovens) no banco de dados
        cursor.execute("""
            UPDATE matriculas
            SET status = ?, nota1 = ?
            WHERE id = ?
        """, (final_status, matricula_data['nota1'], matricula_id)) # Usa matricula_data['nota1'] que pode ter sido atualizado
        conn.commit()

    except Exception as e:
        app.logger.error(f"Erro ao atualizar status da matrícula {matricula_id}: {e}", exc_info=True)
    finally:
        conn.close()


@app.route("/matriculas")
@login_required
def matriculas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.id, a.nome AS aluno_nome, d.nome AS disciplina_nome,
               t.nome AS turma_nome, t.faixa_etaria,
               m.data_inicio, m.data_conclusao, m.status,
               m.nota1, m.nota2, m.participacao, m.desafio, m.prova,
               m.meditacao, m.versiculos, m.desafio_nota, m.visitante,
               d.tem_atividades, d.frequencia_minima
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
            SELECT presente
            FROM presencas
            WHERE matricula_id = ?
        """, (matricula_dict['id'],))
        historico_chamadas = cursor.fetchall()

        presencas = sum(1 for c in historico_chamadas if c['presente'])
        total_aulas = len(historico_chamadas)

        frequencia_porcentagem = (presencas / total_aulas * 100) if total_aulas > 0 else 0
        matricula_dict['frequencia_porcentagem'] = frequencia_porcentagem

        # --- Cálculo de Notas e Status ---
        faixa_etaria = matricula_dict.get('faixa_etaria', 'adultos') # Usar a faixa etaria da turma do aluno

        nota_final_calc = None
        status_display = matricula_dict['status'] # Default

        if faixa_etaria in ['criancas_0_3', 'criancas_4_7', 'criancas_8_12']:
            matricula_dict['media_display'] = 'N/A'
            if frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                status_display = 'Aprovado (Frequência)'
            else:
                status_display = 'Reprovado (Frequência)'

        elif faixa_etaria in ['adolescentes_13_15', 'jovens_16_17']:
            meditacao = matricula_dict['meditacao'] if matricula_dict['meditacao'] is not None else 0
            versiculos = matricula_dict['versiculos'] if matricula_dict['versiculos'] is not None else 0
            desafio_nota = matricula_dict['desafio_nota'] if matricula_dict['desafio_nota'] is not None else 0
            visitante = matricula_dict['visitante'] if matricula_dict['visitante'] is not None else 0
            nota_final_calc = meditacao + versiculos + desafio_nota + visitante
            matricula_dict['media_display'] = f"{nota_final_calc:.1f}"

            if nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                status_display = 'Aprovado'
            elif nota_final_calc >= 5.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                status_display = 'Recuperação'
            else:
                status_display = 'Reprovado'

        else: # Adultos
            n1 = matricula_dict['nota1'] if matricula_dict['nota1'] is not None else 0
            n2 = matricula_dict['nota2'] if matricula_dict['nota2'] is not None else 0
            nota_final_calc = (n1 + n2) / 2 if (n1 is not None and n2 is not None) else None
            matricula_dict['media_display'] = f"{nota_final_calc:.1f}" if nota_final_calc is not None else '—'

            if nota_final_calc is not None:
                if nota_final_calc >= 7.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status_display = 'Aprovado'
                elif nota_final_calc >= 5.0 and frequencia_porcentagem >= matricula_dict['frequencia_minima']:
                    status_display = 'Recuperação'
                else:
                    status_display = 'Reprovado'
            else:
                status_display = 'Cursando' # Sem notas, assume cursando

        matricula_dict['status_display'] = status_display
        processed_matriculas.append(matricula_dict)

    conn.close()
    return render_template("matriculas.html", matriculas=processed_matriculas)


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

        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Todos os campos obrigatórios devem ser preenchidos!", "erro")
            alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
            disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                                   alunos=alunos_disponiveis,
                                   disciplinas=disciplinas_disponiveis,
                                   now=date.today())

        try:
            # Verificar se a matrícula já existe
            cursor.execute("""
                SELECT id FROM matriculas WHERE aluno_id = ? AND disciplina_id = ?
            """, (aluno_id, disciplina_id))
            if cursor.fetchone():
                flash("Este aluno já está matriculado nesta disciplina!", "erro")
                alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
                disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
                conn.close()
                return render_template("nova_matricula.html",
                                       alunos=alunos_disponiveis,
                                       disciplinas=disciplinas_disponiveis,
                                       now=date.today())

            cursor.execute("""
                INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, data_conclusao, status)
                VALUES (?, ?, ?, ?, 'cursando')
            """, (aluno_id, disciplina_id, data_inicio, data_conclusao))
            conn.commit()
            flash("Matrícula realizada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao matricular: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
    disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
    conn.close()
    return render_template("nova_matricula.html",
                           alunos=alunos_disponiveis,
                           disciplinas=disciplinas_disponiveis,
                           now=date.today())


@app.route("/matriculas/novo_aluno_disciplina", methods=["GET", "POST"])
@login_required
def novo_aluno_disciplina():
    conn = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        aluno_id = request.form.get("aluno_id", type=int)
        disciplina_id = request.form.get("disciplina_id", type=int)
        data_inicio = request.form.get("data_inicio", "").strip()
        data_conclusao = request.form.get("data_conclusao", "").strip() or None

        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Todos os campos obrigatórios devem ser preenchidos!", "erro")
            alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
            disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
            conn.close()
            return render_template("novo_aluno_disciplina.html",
                                   alunos=alunos_disponiveis,
                                   disciplinas=disciplinas_disponiveis,
                                   now=date.today())

        try:
            # Verificar se a matrícula já existe
            cursor.execute("""
                SELECT id FROM matriculas WHERE aluno_id = ? AND disciplina_id = ?
            """, (aluno_id, disciplina_id))
            if cursor.fetchone():
                flash("Este aluno já está matriculado nesta disciplina!", "erro")
                alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
                disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
                conn.close()
                return render_template("novo_aluno_disciplina.html",
                                       alunos=alunos_disponiveis,
                                       disciplinas=disciplinas_disponiveis,
                                       now=date.today())

            cursor.execute("""
                INSERT INTO matriculas (aluno_id, disciplina_id, data_inicio, data_conclusao, status)
                VALUES (?, ?, ?, ?, 'cursando')
            """, (aluno_id, disciplina_id, data_inicio, data_conclusao))
            conn.commit()
            flash("Matrícula realizada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao matricular: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    alunos_disponiveis = cursor.execute("SELECT id, nome FROM alunos ORDER BY nome").fetchall()
    disciplinas_disponiveis = cursor.execute("SELECT id, nome FROM disciplinas ORDER BY nome").fetchall()
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
            # Isso garante que não haja dados residuais de outras faixas etárias
            cursor.execute("""
                UPDATE matriculas
                SET nota1=NULL, nota2=NULL, participacao=NULL, desafio=NULL, prova=NULL,
                    meditacao=NULL, versiculos=NULL, desafio_nota=NULL, visitante=NULL
                WHERE id=?
            """, (id,))

            if faixa_etaria_matricula in ['adolescentes_13_15', 'jovens_16_17']:
                cursor.execute("""
                    UPDATE matriculas
                    SET data_conclusao=?, status=?,
                        meditacao=?, versiculos=?, desafio_nota=?, visitante=?
                    WHERE id=?
                """, (data_conclusao, status,
                      meditacao_aj, versiculos_aj, desafio_nota_aj, visitante_aj, id))
            elif faixa_etaria_matricula == 'adultos':
                cursor.execute("""
                    UPDATE matriculas
                    SET data_conclusao=?, status=?,
                        nota1=?, nota2=?, participacao=?, desafio=?, prova=?
                    WHERE id=?
                """, (data_conclusao, status,
                      nota1_adulto, nota2_adulto, participacao_adulto, desafio_adulto, prova_adulto, id))
            else: # Crianças (0-3, 4-7, 8-12)
                cursor.execute("""
                    UPDATE matriculas
                    SET data_conclusao=?, status=?
                    WHERE id=?
                """, (data_conclusao, status, id))

            conn.commit()
            _atualizar_status_matricula(id) # Recalcular status após a edição
            flash("Matrícula atualizada com sucesso!", "sucesso")
        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
        finally:
            conn.close()
        return redirect(url_for("matriculas"))

    cursor.execute("""
        SELECT m.*, a.nome AS aluno_nome, d.nome AS disciplina_nome,
               t.nome AS turma_nome, t.faixa_etaria
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
        cursor.execute("DELETE FROM presencas WHERE matricula_id=?", (id,))
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
# DISCIPLINAS
# ══════════════════════════════════════
@app.route("/disciplinas")
@login_required
def disciplinas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT d.*, COUNT(m.id) as total_matriculas
        FROM disciplinas d
        LEFT JOIN matriculas m ON m.disciplina_id = d.id
        GROUP BY d.id ORDER BY d.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("disciplinas.html", disciplinas=lista)


@app.route("/disciplinas/nova", methods=["GET", "POST"])
@login_required
def nova_disciplina():
    if request.method == "POST":
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0 # Default 75%
        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("nova_disciplina"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO disciplinas (nome, descricao, tem_atividades, frequencia_minima)
                VALUES (?, ?, ?, ?)
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
        nome            = request.form.get("nome", "").strip()
        descricao       = request.form.get("descricao", "").strip()
        tem_atividades  = 1 if request.form.get("tem_atividades") else 0
        frequencia_minima = request.form.get("frequencia_minima", type=float) or 75.0
        ativa           = 1 if request.form.get("ativa") else 0
        try:
            cursor.execute("""
                UPDATE disciplinas
                SET nome=?, descricao=?, tem_atividades=?, frequencia_minima=?, ativa=?
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


# ══════════════════════════════════════
# USUÁRIOS (ADMIN-ONLY)
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
@admin_required
def usuarios():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM usuarios ORDER BY nome")
    lista = cursor.fetchall()
    conn.close()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
@admin_required
def novo_usuario():
    if request.method == "POST":
        nome  = request.form.get("nome", "").strip()
        email = request.form.get("email", "").strip()
        senha = request.form.get("senha", "")
        perfil = request.form.get("perfil", "usuario")
        if not nome or not email or not senha:
            flash("Todos os campos são obrigatórios!", "erro")
            return render_template("novo_usuario.html")
        if len(senha) < 6:
            flash("A senha deve ter no mínimo 6 caracteres!", "erro")
            return render_template("novo_usuario.html")
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO usuarios (nome,email,senha_hash,perfil) VALUES (?,?,?,?)",
                (nome, email, generate_password_hash(senha), perfil))
            conn.commit()
            flash(f"Usuário '{nome}' criado!", "sucesso")
        except sqlite3.IntegrityError as e:
            if "usuarios.email" in str(e):
                flash("Já existe um usuário com este e-mail!", "erro")
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
@admin_required
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
@admin_required
def excluir_usuario(id):
    if current_user.id == id:
        flash("Você não pode excluir sua própria conta de administrador!", "erro")
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
# BACKUP E RESTAURAÇÃO (ADMIN-ONLY)
# ══════════════════════════════════════
@app.route("/admin/backup", methods=["GET", "POST"])
@login_required
@admin_required
def backup_restauracao():
    if request.method == "POST":
        if "backup_action" in request.form:
            # Lógica para fazer backup
            try:
                backup_filename = f"escola_backup_{date.today().isoformat()}.db"
                return send_file(DATABASE, as_attachment=True, download_name=backup_filename)
            except Exception as e:
                flash(f"Erro ao gerar backup: {e}", "erro")
                app.logger.error(f"Erro ao gerar backup: {e}", exc_info=True)
        elif "restore_file" in request.files:
            # Lógica para restaurar
            file = request.files["restore_file"]
            if file.filename == '':
                flash("Nenhum arquivo selecionado para restauração.", "erro")
            elif file and file.filename.endswith('.db'):
                try:
                    # Fazer um backup do banco de dados atual antes de sobrescrever
                    shutil.copy(DATABASE, f"{DATABASE}.pre_restore_backup")

                    # Fechar a conexão atual com o banco de dados antes de sobrescrever
                    # Isso é crucial para evitar erros de "database is locked"
                    # No Flask, as conexões são abertas e fechadas por requisição,
                    # mas é bom garantir que não haja nenhuma conexão ativa no momento da cópia.
                    # Para SQLite, o arquivo precisa estar desbloqueado.

                    file.save(DATABASE) # Sobrescreve o arquivo escola.db

                    # Após a restauração, é uma boa prática reiniciar o banco de dados
                    # para garantir que as novas tabelas/dados sejam carregados.
                    # Para o Railway, isso geralmente significa reiniciar o contêiner.
                    # Aqui, chamamos inicializar_banco para garantir que a estrutura esteja ok.
                    inicializar_banco() 

                    flash("Banco de dados restaurado com sucesso! Pode ser necessário reiniciar o servidor para que todas as mudanças sejam aplicadas.", "sucesso")
                except Exception as e:
                    # Se houver um erro na restauração, tentar restaurar o backup pré-restauração
                    if os.path.exists(f"{DATABASE}.pre_restore_backup"):
                        shutil.copy(f"{DATABASE}.pre_restore_backup", DATABASE)
                        flash(f"Erro ao restaurar banco de dados: {e}. O backup anterior foi restaurado.", "erro")
                    else:
                        flash(f"Erro ao restaurar banco de dados: {e}. Não foi possível restaurar o backup anterior.", "erro")
                    app.logger.error(f"Erro ao restaurar banco de dados: {e}", exc_info=True)
                finally:
                    # Limpar o backup pré-restauração
                    if os.path.exists(f"{DATABASE}.pre_restore_backup"):
                        os.remove(f"{DATABASE}.pre_restore_backup")
            else:
                flash("Formato de arquivo inválido. Por favor, selecione um arquivo .db", "erro")

        return redirect(url_for("backup_restauracao"))

    return render_template("backup_restauracao.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)