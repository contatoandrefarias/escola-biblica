# app.py
import os
from flask import (Flask, render_template, request,
                   redirect, url_for, flash)
from flask_login import (LoginManager, login_user,
                         logout_user, login_required,
                         current_user)
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco
from auth import Usuario, carregar_usuario, verificar_login

# ── Criar o app ────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get(
    "SECRET_KEY",
    "escola_biblica_chave_2026"
)

# ── Flask-Login ────────────────────────────────
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view             = "login"
login_manager.login_message          = "Faça login para continuar."
login_manager.login_message_category = "warning"

@login_manager.user_loader
def load_user(user_id):
    return carregar_usuario(user_id)

# ══════════════════════════════════════════════
# INICIALIZAR O BANCO AQUI — ANTES DE TUDO!
# ══════════════════════════════════════════════
inicializar_banco()


# ╔══════════════════════════════════════════════╗
# ║                   LOGIN                      ║
# ╚══════════════════════════════════════════════╝
@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        senha = request.form.get("senha", "")

        usuario = verificar_login(email, senha)
        if usuario:
            login_user(usuario, remember=True)
            flash(f"Bem-vindo(a), {usuario.nome}! 🙏", "sucesso")
            proxima = request.args.get("next")
            return redirect(proxima or url_for("index"))
        else:
            flash("E-mail ou senha incorretos!", "erro")

    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    nome = current_user.nome
    logout_user()
    flash(f"Até logo, {nome}! ✝️", "sucesso")
    return redirect(url_for("login"))


# ╔══════════════════════════════════════════════╗
# ║              PAINEL PRINCIPAL                ║
# ╚══════════════════════════════════════════════╝
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
        "SELECT COUNT(*) as t FROM disciplinas WHERE ativa=1"
    )
    total_disciplinas = cursor.fetchone()["t"]

    cursor.execute(
        "SELECT COUNT(*) as t FROM matriculas WHERE status='aprovado'"
    )
    aprovados = cursor.fetchone()["t"]

    cursor.execute(
        "SELECT COUNT(*) as t FROM matriculas WHERE status='reprovado'"
    )
    reprovados = cursor.fetchone()["t"]

    cursor.execute(
        "SELECT COUNT(*) as t FROM matriculas WHERE status='cursando'"
    )
    cursando = cursor.fetchone()["t"]

    conn.close()

    return render_template("index.html",
        total_alunos      = total_alunos,
        total_professores = total_professores,
        total_disciplinas = total_disciplinas,
        aprovados         = aprovados,
        reprovados        = reprovados,
        cursando          = cursando
    )


# ╔══════════════════════════════════════════════╗
# ║                  ALUNOS                      ║
# ╚══════════════════════════════════════════════╝
@app.route("/alunos")
@login_required
def alunos():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM alunos ORDER BY nome")
    lista  = cursor.fetchall()
    conn.close()
    return render_template("alunos.html", alunos=lista)


@app.route("/alunos/novo", methods=["GET", "POST"])
@login_required
def novo_aluno():
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        telefone  = request.form.get("telefone", "").strip()
        email     = request.form.get("email", "").strip()
        data_nasc = request.form.get("data_nascimento", "").strip()
        membro    = 1 if request.form.get("membro_igreja") else 0

        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("novo_aluno"))

        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO alunos
                (nome, telefone, email,
                 data_nascimento, membro_igreja)
            VALUES (?, ?, ?, ?, ?)
        """, (nome, telefone, email, data_nasc, membro))
        conn.commit()
        conn.close()
        flash(f"Aluno '{nome}' cadastrado!", "sucesso")
        return redirect(url_for("alunos"))

    return render_template("novo_aluno.html")


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

        cursor.execute("""
            UPDATE alunos
            SET nome=?, telefone=?, email=?, membro_igreja=?
            WHERE id=?
        """, (nome, telefone, email, membro, id))
        conn.commit()
        conn.close()
        flash("Aluno atualizado!", "sucesso")
        return redirect(url_for("alunos"))

    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    conn.close()

    if not aluno:
        flash("Aluno não encontrado!", "erro")
        return redirect(url_for("alunos"))

    return render_template("editar_aluno.html", aluno=aluno)


@app.route("/alunos/<int:id>/trilha")
@login_required
def trilha(id):
    conn   = conectar()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()

    if not aluno:
        flash("Aluno não encontrado!", "erro")
        conn.close()
        return redirect(url_for("alunos"))

    cursor.execute("""
        SELECT d.nome, d.carga_horaria,
               m.nota1, m.nota2, m.nota_final, m.status,
               COUNT(p.id)     as total_aulas,
               SUM(p.presente) as presencas
        FROM matriculas m
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN presencas p ON p.matricula_id = m.id
        WHERE m.aluno_id = ?
        GROUP BY m.id, d.nome
        ORDER BY
            CASE m.status
                WHEN 'aprovado'  THEN 1
                WHEN 'cursando'  THEN 2
                WHEN 'reprovado' THEN 3
            END, d.nome
    """, (id,))
    trilha_dados = cursor.fetchall()
    conn.close()

    aprovadas  = sum(1 for t in trilha_dados
                     if t["status"] == "aprovado")
    reprovadas = sum(1 for t in trilha_dados
                     if t["status"] == "reprovado")
    cursando   = sum(1 for t in trilha_dados
                     if t["status"] == "cursando")

    return render_template("trilha.html",
        aluno      = aluno,
        trilha     = trilha_dados,
        aprovadas  = aprovadas,
        reprovadas = reprovadas,
        cursando   = cursando
    )


# ╔══════════════════════════════════════════════╗
# ║               PROFESSORES                    ║
# ╚══════════════════════════════════════════════╝
@app.route("/professores")
@login_required
def professores():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    lista  = cursor.fetchall()
    conn.close()
    return render_template("professores.html", professores=lista)


@app.route("/professores/novo", methods=["GET", "POST"])
@login_required
def novo_professor():
    if request.method == "POST":
        nome          = request.form.get("nome", "").strip()
        telefone      = request.form.get("telefone", "").strip()
        email         = request.form.get("email", "").strip()
        especialidade = request.form.get(
            "especialidade", ""
        ).strip()

        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("novo_professor"))

        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO professores
                (nome, telefone, email, especialidade)
            VALUES (?, ?, ?, ?)
        """, (nome, telefone, email, especialidade))
        conn.commit()
        conn.close()
        flash(f"Professor '{nome}' cadastrado!", "sucesso")
        return redirect(url_for("professores"))

    return render_template("novo_professor.html")


# ╔══════════════════════════════════════════════╗
# ║               DISCIPLINAS                    ║
# ╚══════════════════════════════════════════════╝
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
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        descricao = request.form.get("descricao", "").strip()
        carga     = request.form.get("carga_horaria", "0")
        nota_min  = request.form.get("nota_minima", "6.0")
        freq_min  = request.form.get("frequencia_minima", "75")
        prof_id   = request.form.get("professor_id") or None

        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("nova_disciplina"))

        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO disciplinas
                (nome, descricao, carga_horaria,
                 nota_minima, frequencia_minima, professor_id)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nome, descricao, int(carga),
              float(nota_min), float(freq_min), prof_id))
        conn.commit()
        conn.close()
        flash(f"Disciplina '{nome}' cadastrada!", "sucesso")
        return redirect(url_for("disciplinas"))

    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    @app.route("/disciplinas/nova", methods=["GET", "POST"])
@login_required
def nova_disciplina():
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        descricao = request.form.get("descricao", "").strip()
        carga     = request.form.get("carga_horaria", "0")
        nota_min  = request.form.get("nota_minima", "6.0")
        freq_min  = request.form.get("frequencia_minima", "75")
        prof_id   = request.form.get("professor_id") or None

        if not nome:
            flash("Nome é obrigatório!", "erro")
            return redirect(url_for("nova_disciplina"))

        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO disciplinas
                (nome, descricao, carga_horaria,
                 nota_minima, frequencia_minima, professor_id)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nome, descricao, int(carga),
              float(nota_min), float(freq_min), prof_id))
        conn.commit()
        conn.close()
        flash(f"Disciplina '{nome}' cadastrada!", "sucesso")
        return redirect(url_for("disciplinas"))

    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    profs  = cursor.fetchall()
    conn.close()
    return render_template("nova_disciplina.html",
                           professores=profs)

