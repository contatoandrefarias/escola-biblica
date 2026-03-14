import os
from datetime import date
from flask import (Flask, render_template, request,
                   redirect, url_for, flash, send_file)
from flask_login import (LoginManager, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco
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

        # Verificar se o email já existe ANTES de tentar inserir
        if email: # Email pode ser opcional, mas se preenchido, deve ser único
            cursor.execute("SELECT id FROM alunos WHERE email = ?", (email,))
            if cursor.fetchone():
                flash("Este e-mail já está cadastrado para outro aluno!", "erro")
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
        except Exception as e:
            # Capturar outros erros, como 'database is locked'
            flash(f"Erro ao cadastrar aluno: {e}", "erro")
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

        # Verificar se o email já existe para OUTRO aluno
        if email:
            cursor.execute("SELECT id FROM alunos WHERE email = ? AND id != ?", (email, id))
            if cursor.fetchone():
                flash("Este e-mail já está cadastrado para outro aluno!", "erro")
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
        except Exception as e:
            flash(f"Erro ao atualizar aluno: {e}", "erro")
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
        except Exception as e:
            flash(f"Erro ao cadastrar professor: {e}", "erro")
            return redirect(url_for("novo_professor"))
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
        try:
            cursor.execute("""
                UPDATE professores
                SET nome=?,telefone=?,email=?,especialidade=?
                WHERE id=?
            """, (nome, telefone, email, especialidade, id))
            conn.commit()
            flash("Professor atualizado!", "sucesso")
            return redirect(url_for("professores"))
        except Exception as e:
            flash(f"Erro ao atualizar professor: {e}", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
            prof = cursor.fetchone()
            conn.close()
            return render_template("editar_professor.html", professor=prof)
        finally:
            conn.close()
    cursor.execute("SELECT * FROM professores WHERE id=?", (id,))
    prof = cursor.fetchone()
    conn.close()
    if not prof:
        flash("Professor nao encontrado!", "erro")
        return redirect(url_for("professores"))
    return render_template("editar_professor.html", professor=prof)


# ══════════════════════════════════════
# DISCIPLINAS
# ══════════════════════════════════════
@app.route("/disciplinas")
@login_required
def disciplinas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT d.*, p.nome as prof_nome
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
        nome      = request.form.get("nome", "").strip()
        descricao = request.form.get("descricao", "").strip()
        semanas   = request.form.get("duracao_semanas", "4")
        nota_min  = request.form.get("nota_minima", "6.0")
        freq_min  = request.form.get("frequencia_minima", "75")
        tem_ativ  = 1 if request.form.get("tem_atividades") else 0
        prof_id   = request.form.get("professor_id") or None
        if not nome:
            flash("Nome e obrigatorio!", "erro")
            cursor.execute("SELECT * FROM professores ORDER BY nome")
            profs = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=profs)
        try:
            cursor.execute("""
                INSERT INTO disciplinas
                    (nome,descricao,duracao_semanas,nota_minima,
                     frequencia_minima,tem_atividades,professor_id)
                VALUES (?,?,?,?,?,?,?)
            """, (nome, descricao, int(semanas), float(nota_min),
                  float(freq_min), tem_ativ, prof_id))
            conn.commit()
            flash(f"Disciplina '{nome}' cadastrada!", "sucesso")
            return redirect(url_for("disciplinas"))
        except Exception as e:
            flash(f"Erro ao cadastrar disciplina: {e}", "erro")
            cursor.execute("SELECT * FROM professores ORDER BY nome")
            profs = cursor.fetchall()
            conn.close()
            return render_template("nova_disciplina.html", professores=profs)
        finally:
            conn.close()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    profs = cursor.fetchall()
    conn.close()
    return render_template("nova_disciplina.html", professores=profs)


@app.route("/disciplinas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_disciplina(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        descricao = request.form.get("descricao", "").strip()
        semanas   = request.form.get("duracao_semanas", "4")
        nota_min  = request.form.get("nota_minima", "6.0")
        freq_min  = request.form.get("frequencia_minima", "75")
        tem_ativ  = 1 if request.form.get("tem_atividades") else 0
        prof_id   = request.form.get("professor_id") or None
        ativa     = 1 if request.form.get("ativa") else 0
        try:
            cursor.execute("""
                UPDATE disciplinas
                SET nome=?,descricao=?,duracao_semanas=?,nota_minima=?,
                    frequencia_minima=?,tem_atividades=?,professor_id=?,ativa=?
                WHERE id=?
            """, (nome, descricao, int(semanas), float(nota_min),
                  float(freq_min), tem_ativ, prof_id, ativa, id))
            conn.commit()
            flash("Disciplina atualizada!", "sucesso")
            return redirect(url_for("disciplinas"))
        except Exception as e:
            flash(f"Erro ao atualizar disciplina: {e}", "erro")
            # Recarregar dados para o template
            cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
            disc = cursor.fetchone()
            cursor.execute("SELECT * FROM professores ORDER BY nome")
            profs = cursor.fetchall()
            conn.close()
            return render_template("editar_disciplina.html", disciplina=disc, professores=profs)
        finally:
            conn.close()
    cursor.execute("SELECT * FROM disciplinas WHERE id=?", (id,))
    disc = cursor.fetchone()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    profs = cursor.fetchall()
    conn.close()
    if not disc:
        flash("Disciplina nao encontrada!", "erro")
        return redirect(url_for("disciplinas"))
    return render_template("editar_disciplina.html",
        disciplina=disc, professores=profs)


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
               m.data_inicio, m.data_conclusao, m.status, m.nota_final
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
        aluno_id      = request.form.get("aluno_id")
        disciplina_id = request.form.get("disciplina_id")
        data_inicio   = request.form.get("data_inicio")
        if not aluno_id or not disciplina_id or not data_inicio:
            flash("Todos os campos sao obrigatorios!", "erro")
            cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                alunos=alunos_lista, disciplinas=disciplinas_lista)
        try:
            cursor.execute("""
                INSERT INTO matriculas
                    (aluno_id,disciplina_id,data_inicio,status)
                VALUES (?,?,?,?)
            """, (aluno_id, disciplina_id, data_inicio, 'cursando'))
            conn.commit()
            flash("Matricula criada com sucesso!", "sucesso")
            return redirect(url_for("matriculas"))
        except sqlite3.IntegrityError:
            flash("Este aluno ja esta matriculado nesta disciplina!", "erro")
            cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                alunos=alunos_lista, disciplinas=disciplinas_lista)
        except Exception as e:
            flash(f"Erro ao criar matricula: {e}", "erro")
            cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
            alunos_lista = cursor.fetchall()
            cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
            disciplinas_lista = cursor.fetchall()
            conn.close()
            return render_template("nova_matricula.html",
                alunos=alunos_lista, disciplinas=disciplinas_lista)
        finally:
            conn.close()
    cursor.execute("SELECT id,nome FROM alunos ORDER BY nome")
    alunos_lista = cursor.fetchall()
    cursor.execute("SELECT id,nome FROM disciplinas WHERE ativa=1 ORDER BY nome")
    disciplinas_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html",
        alunos=alunos_lista, disciplinas=disciplinas_lista)


@app.route("/matriculas/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_matricula(id):
    conn   = conectar()
    cursor = conn.cursor()

    if request.method == "POST":
        nota1_str         = request.form.get("nota1")
        nota2_str         = request.form.get("nota2")
        nota_final_str    = request.form.get("nota_final")
        data_inicio       = request.form.get("data_inicio")
        data_conclusao    = request.form.get("data_conclusao")

        nota1 = float(nota1_str) if nota1_str else None
        nota2 = float(nota2_str) if nota2_str else None
        nota_final = float(nota_final_str) if nota_final_str else None

        # Obter dados da matrícula e disciplina para cálculo
        cursor.execute("""
            SELECT m.aluno_id, m.disciplina_id, d.nota_minima, d.frequencia_minima
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.id = ?
        """, (id,))
        mat_disc_data = cursor.fetchone()

        if not mat_disc_data:
            flash("Matrícula ou disciplina não encontrada!", "erro")
            conn.close()
            return redirect(url_for("matriculas"))

        aluno_id          = mat_disc_data['aluno_id']
        disciplina_id     = mat_disc_data['disciplina_id']
        nota_minima       = mat_disc_data['nota_minima']
        frequencia_minima = mat_disc_data['frequencia_minima']

        # 1. Calcular Nota Final (se não sobrescrita)
        if nota_final is None and nota1 is not None and nota2 is not None:
            nota_final = (nota1 + nota2) / 2

        # 2. Calcular Frequência
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
            frequencia_percentual = (presencas / total_aulas) * 100

        # 3. Determinar Status
        novo_status = 'cursando'
        if data_conclusao: # Só pode ser aprovado/reprovado se houver data de conclusão
            if nota_final is not None and frequencia_percentual >= frequencia_minima:
                if nota_final >= nota_minima:
                    novo_status = 'aprovado'
                else:
                    novo_status = 'reprovado'
            else:
                novo_status = 'reprovado' # Reprovado por nota ou frequência insuficiente

        try:
            cursor.execute("""
                UPDATE matriculas
                SET nota1=?, nota2=?, nota_final=?, status=?,
                    data_inicio=?, data_conclusao=?
                WHERE id=?
            """, (nota1, nota2, nota_final, novo_status,
                  data_inicio, data_conclusao, id))
            conn.commit()
            flash("Matrícula atualizada e status calculado!", "sucesso")
            return redirect(url_for("matriculas"))
        except Exception as e:
            flash(f"Erro ao atualizar matrícula: {e}", "erro")
            # Recarregar dados para o template
            # (Este bloco é para o caso de erro no UPDATE, então precisamos recarregar os dados)
            cursor.execute("""
                SELECT m.*, a.nome as aluno_nome, d.nome as disciplina_nome,
                       d.nota_minima, d.frequencia_minima
                FROM matriculas m
                JOIN alunos a ON m.aluno_id = a.id
                JOIN disciplinas d ON m.disciplina_id = d.id
                WHERE m.id = ?
            """, (id,))
            matricula = cursor.fetchone()

            # Recalcular frequência para exibir no formulário em caso de erro
            cursor.execute("""
                SELECT COUNT(id) as total_aulas, SUM(presente) as presencas
                FROM presencas
                WHERE matricula_id = ?
            """, (id,))
            presenca_data_erro = cursor.fetchone()
            total_aulas_erro = presenca_data_erro['total_aulas'] or 0
            presencas_erro   = presenca_data_erro['presencas'] or 0
            frequencia_percentual_erro = 0.0
            if total_aulas_erro > 0:
                frequencia_percentual_erro = (presencas_erro / total_aulas_erro) * 100

            conn.close()
            return render_template("editar_matricula.html",
                matricula=matricula,
                total_aulas=total_aulas_erro,
                presencas=presencas_erro,
                frequencia_percentual=frequencia_percentual_erro)
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

    # Calcular frequência para exibir no formulário
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
        frequencia_percentual = (presencas / total_aulas) * 100

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
    hoje = date.today().isoformat()

    disciplinas_cursando = []
    disciplinas_concluidas = []

    if current_user.is_aluno and current_user.aluno_id:
        aluno_id = current_user.aluno_id
        cursor.execute("""
            SELECT d.id, d.nome, m.status
            FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.aluno_id = ? AND d.ativa = 1
            ORDER BY d.nome
        """, (aluno_id,))
        disciplinas_aluno = cursor.fetchall()

        for disc in disciplinas_aluno:
            if disc['status'] == 'cursando':
                disciplinas_cursando.append(disc)
            else:
                disciplinas_concluidas.append(disc)
    else: # Admin ou Professor veem todas as disciplinas ativas
        cursor.execute("SELECT id, nome FROM disciplinas WHERE ativa = 1 ORDER BY nome")
        disciplinas_ativas = cursor.fetchall()
        disciplinas_cursando = disciplinas_ativas # Para admin/prof, todas ativas são "cursando" para chamada

    conn.close()
    return render_template("presenca.html",
        hoje=hoje,
        disciplinas_cursando=disciplinas_cursando,
        disciplinas_concluidas=disciplinas_concluidas)


@app.route("/presenca/chamada")
@login_required
def chamada():
    disc_id   = request.args.get("disciplina_id")
    data_aula = request.args.get("data_aula")
    if not disc_id or not data_aula:
        flash("Selecione uma disciplina e data para a chamada.", "erro")
        return redirect(url_for("presenca"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM disciplinas WHERE id=?", (disc_id,))
    disc = cursor.fetchone()
    if not disc:
        flash("Disciplina nao encontrada!", "erro")
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
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        LEFT JOIN presencas p ON p.matricula_id = m.id AND p.data_aula = ?
        WHERE m.disciplina_id = ?
        ORDER BY a.nome
    """, (data_aula, disc_id))
    alunos_lista = cursor.fetchall()
    conn.close()
    return render_template("chamada.html",
        disciplina=disc,
        alunos=alunos_lista,
        data_aula=data_aula)


@app.route("/presenca/salvar", methods=["POST"])
@login_required
def salvar_chamada():
    disc_id   = request.form.get("disciplina_id")
    data_aula = request.form.get("data_aula")
    conn   = conectar()
    cursor = conn.cursor()

    # Obter todas as matriculas para a disciplina
    cursor.execute("""
        SELECT m.id as mat_id
        FROM matriculas m
        WHERE m.disciplina_id = ?
    """, (disc_id,))
    mats = cursor.fetchall()

    try:
        for m in mats:
            mat_id   = m["mat_id"]
            presente = 1 if request.form.get(
                f"presenca_{mat_id}") else 0
            fez_ativ = 1 if request.form.get(
                f"atividade_{mat_id}") else 0

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
        except Exception:
            flash("E-mail ja cadastrado!", "erro")
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
        cursor.execute(
            "UPDATE usuarios SET senha_hash=? WHERE id=?",
            (generate_password_hash(nova_senha), current_user.id))
        conn.commit()
        conn.close()
        flash("Senha alterada com sucesso!", "sucesso")
        return redirect(url_for("index"))
    return render_template("minha_conta.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)