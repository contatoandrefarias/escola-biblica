import os
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash
from database import conectar, inicializar_banco
from auth import carregar_usuario, verificar_login

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
            return redirect(request.args.get("next") or url_for("index"))
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
    cursor.execute("SELECT COUNT(*) as t FROM disciplinas WHERE ativa=1")
    total_disciplinas = cursor.fetchone()["t"]
    cursor.execute("SELECT COUNT(*) as t FROM turmas WHERE ativa=1")
    total_turmas = cursor.fetchone()["t"]
    cursor.execute("SELECT COUNT(*) as t FROM matriculas WHERE status='aprovado'")
    aprovados = cursor.fetchone()["t"]
    cursor.execute("SELECT COUNT(*) as t FROM matriculas WHERE status='reprovado'")
    reprovados = cursor.fetchone()["t"]
    cursor.execute("SELECT COUNT(*) as t FROM matriculas WHERE status='cursando'")
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
        cursor.execute("""
            INSERT INTO turmas (nome, descricao, faixa_etaria)
            VALUES (?, ?, ?)
        """, (nome, descricao, faixa_etaria))
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
            SET nome=?, descricao=?, faixa_etaria=?, ativa=?
            WHERE id=?
        """, (nome, descricao, faixa_etaria, ativa, id))
        conn.commit()
        conn.close()
        flash("Turma atualizada!", "sucesso")
        return redirect(url_for("turmas"))
    cursor.execute("SELECT * FROM turmas WHERE id=?", (id,))
    turma = cursor.fetchone()
    cursor.execute("""
        SELECT * FROM alunos WHERE turma_id=? ORDER BY nome
    """, (id,))
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
    if request.method == "POST":
        nome      = request.form.get("nome", "").strip()
        telefone  = request.form.get("telefone", "").strip()
        email     = request.form.get("email", "").strip()
        data_nasc = request.form.get("data_nascimento", "").strip()
        membro    = 1 if request.form.get("membro_igreja") else 0
        turma_id  = request.form.get("turma_id") or None
        if not nome:
            flash("Nome e obrigatorio!", "erro")
            return redirect(url_for("novo_aluno"))
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO alunos
                (nome, telefone, email, data_nascimento,
                 membro_igreja, turma_id)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (nome, telefone, email, data_nasc, membro, turma_id))
        conn.commit()
        conn.close()
        flash(f"Aluno '{nome}' cadastrado!", "sucesso")
        return redirect(url_for("alunos"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
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
        cursor.execute("""
            UPDATE alunos
            SET nome=?, telefone=?, email=?,
                membro_igreja=?, turma_id=?
            WHERE id=?
        """, (nome, telefone, email, membro, turma_id, id))
        conn.commit()
        conn.close()
        flash("Aluno atualizado!", "sucesso")
        return redirect(url_for("alunos"))
    cursor.execute("SELECT * FROM alunos WHERE id=?", (id,))
    aluno = cursor.fetchone()
    cursor.execute("SELECT id, nome FROM turmas WHERE ativa=1 ORDER BY nome")
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
        SELECT d.nome, d.duracao_semanas,
               m.nota1, m.nota2, m.nota_final, m.status,
               COUNT(p.id) as total_aulas,
               SUM(p.presente) as presencas,
               SUM(p.fez_atividade) as atividades
        FROM matriculas m
        JOIN disciplinas d ON m.disciplina_id = d.id
        LEFT JOIN presencas p ON p.matricula_id = m.id
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
    cursando   = sum(1 for t in trilha_dados if t["status"] == "cursando")
    return render_template("trilha.html",
        aluno=aluno, trilha=trilha_dados,
        aprovadas=aprovadas, reprovadas=reprovadas,
        cursando=cursando)


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
            return redirect(url_for("nova_disciplina"))
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO disciplinas
                (nome, descricao, duracao_semanas, nota_minima,
                 frequencia_minima, tem_atividades, professor_id)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (nome, descricao, int(semanas),
              float(nota_min), float(freq_min), tem_ativ, prof_id))
        conn.commit()
        conn.close()
        flash(f"Disciplina '{nome}' cadastrada!", "sucesso")
        return redirect(url_for("disciplinas"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM professores ORDER BY nome")
    profs = cursor.fetchall()
    conn.close()
    return render_template("nova_disciplina.html", professores=profs)


# ══════════════════════════════════════
# MATRICULAS
# ══════════════════════════════════════
@app.route("/matriculas")
@login_required
def matriculas():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.id, a.nome as aluno, d.nome as disciplina,
               m.nota1, m.nota2, m.nota_final, m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        ORDER BY a.nome
    """)
    lista = cursor.fetchall()
    conn.close()
    return render_template("matriculas.html", matriculas=lista)


@app.route("/matriculas/nova", methods=["GET", "POST"])
@login_required
def nova_matricula():
    if request.method == "POST":
        aluno_id = request.form.get("aluno_id")
        disc_id  = request.form.get("disciplina_id")
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO matriculas (aluno_id, disciplina_id)
                VALUES (?, ?)
            """, (aluno_id, disc_id))
            conn.commit()
            flash("Matricula realizada!", "sucesso")
        except Exception:
            flash("Aluno ja matriculado nesta disciplina!", "erro")
        conn.close()
        return redirect(url_for("matriculas"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nome FROM alunos ORDER BY nome")
    alunos_lista = cursor.fetchall()
    cursor.execute("""
        SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome
    """)
    disc_lista = cursor.fetchall()
    conn.close()
    return render_template("nova_matricula.html",
        alunos=alunos_lista, disciplinas=disc_lista)


@app.route("/matriculas/<int:id>/notas", methods=["GET", "POST"])
@login_required
def lancar_notas(id):
    conn   = conectar()
    cursor = conn.cursor()
    if request.method == "POST":
        nota1_str = request.form.get("nota1", "").strip()
        nota2_str = request.form.get("nota2", "").strip()
        cursor.execute("""
            SELECT m.*, d.nota_minima FROM matriculas m
            JOIN disciplinas d ON m.disciplina_id = d.id
            WHERE m.id=?
        """, (id,))
        mat = cursor.fetchone()
        try:
            n1 = float(nota1_str.replace(",", ".")) if nota1_str else mat["nota1"]
            n2 = float(nota2_str.replace(",", ".")) if nota2_str else mat["nota2"]
        except ValueError:
            flash("Nota invalida!", "erro")
            conn.close()
            return redirect(url_for("lancar_notas", id=id))
        nota_final = None
        status = mat["status"]
        if n1 is not None and n2 is not None:
            nota_final = round((n1 + n2) / 2, 2)
            status = "aprovado" if nota_final >= mat["nota_minima"] else "reprovado"
        cursor.execute("""
            UPDATE matriculas
            SET nota1=?, nota2=?, nota_final=?, status=?
            WHERE id=?
        """, (n1, n2, nota_final, status, id))
        conn.commit()
        conn.close()
        flash("Notas salvas!", "sucesso")
        return redirect(url_for("matriculas"))
    cursor.execute("""
        SELECT m.*, a.nome as aluno,
               d.nome as disciplina, d.nota_minima
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        WHERE m.id=?
    """, (id,))
    mat = cursor.fetchone()
    conn.close()
    if not mat:
        flash("Matricula nao encontrada!", "erro")
        return redirect(url_for("matriculas"))
    return render_template("notas.html", matricula=mat)


# ══════════════════════════════════════
# PRESENCA
# ══════════════════════════════════════
@app.route("/presenca")
@login_required
def presenca():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, nome FROM disciplinas WHERE ativa=1 ORDER BY nome
    """)
    disc_lista = cursor.fetchall()
    conn.close()
    return render_template("presenca.html", disciplinas=disc_lista)


@app.route("/presenca/chamada")
@login_required
def chamada():
    disc_id   = request.args.get("disciplina_id")
    data_aula = request.args.get("data_aula")
    if not disc_id or not data_aula:
        flash("Selecione disciplina e data!", "erro")
        return redirect(url_for("presenca"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.id as mat_id, a.nome,
               COALESCE(p.presente, 0)      as presente,
               COALESCE(p.fez_atividade, 0) as fez_atividade
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        LEFT JOIN presencas p
            ON p.matricula_id = m.id AND p.data_aula = ?
        WHERE m.disciplina_id = ? AND m.status = 'cursando'
        ORDER BY a.nome
    """, (data_aula, disc_id))
    alunos_disc = cursor.fetchall()
    cursor.execute("""
        SELECT id, nome, tem_atividades
        FROM disciplinas WHERE id=?
    """, (disc_id,))
    disc = cursor.fetchone()
    conn.close()
    return render_template("chamada.html",
        alunos=alunos_disc,
        disciplina=disc,
        data_aula=data_aula)


@app.route("/presenca/salvar", methods=["POST"])
@login_required
def salvar_presenca():
    disc_id   = request.form.get("disciplina_id")
    data_aula = request.form.get("data_aula")
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT m.id as mat_id FROM matriculas m
        WHERE m.disciplina_id=? AND m.status='cursando'
    """, (disc_id,))
    matriculados = cursor.fetchall()
    for m in matriculados:
        presente = 1 if request.form.get(f"presenca_{m['mat_id']}") else 0
        fez_ativ = 1 if request.form.get(f"atividade_{m['mat_id']}") else 0
        cursor.execute("""
            SELECT id FROM presencas
            WHERE matricula_id=? AND data_aula=?
        """, (m["mat_id"], data_aula))
        existe = cursor.fetchone()
        if existe:
            cursor.execute("""
                UPDATE presencas
                SET presente=?, fez_atividade=? WHERE id=?
            """, (presente, fez_ativ, existe["id"]))
        else:
            cursor.execute("""
                INSERT INTO presencas
                    (matricula_id, data_aula, presente, fez_atividade)
                VALUES (?, ?, ?, ?)
            """, (m["mat_id"], data_aula, presente, fez_ativ))
    conn.commit()
    conn.close()
    flash("Chamada salva com sucesso!", "sucesso")
    return redirect(url_for("presenca"))


# ══════════════════════════════════════
# RELATORIOS
# ══════════════════════════════════════
@app.route("/relatorios")
@login_required
def relatorios():
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT a.nome as aluno, d.nome as disciplina,
               m.nota_final, m.status
        FROM matriculas m
        JOIN alunos a ON m.aluno_id = a.id
        JOIN disciplinas d ON m.disciplina_id = d.id
        ORDER BY m.status, a.nome
    """)
    todas = cursor.fetchall()
    cursor.execute("""
        SELECT a.nome,
               COUNT(m.id) as total,
               SUM(CASE WHEN m.status='aprovado' THEN 1 ELSE 0 END) as aprovadas,
               AVG(m.nota_final) as media
        FROM alunos a
        JOIN matriculas m ON m.aluno_id = a.id
        GROUP BY a.id, a.nome HAVING total > 0
        ORDER BY media DESC
    """)
    ranking = cursor.fetchall()
    conn.close()
    return render_template("relatorios.html",
        matriculas=todas, ranking=ranking)


# ══════════════════════════════════════
# USUARIOS
# ══════════════════════════════════════
@app.route("/usuarios")
@login_required
def usuarios():
    if not current_user.is_admin:
        flash("Acesso negado!", "erro")
        return redirect(url_for("index"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM usuarios ORDER BY nome")
    lista = cursor.fetchall()
    conn.close()
    return render_template("usuarios.html", usuarios=lista)


@app.route("/usuarios/novo", methods=["GET", "POST"])
@login_required
def novo_usuario():
    if not current_user.is_admin:
        flash("Acesso negado!", "erro")
        return redirect(url_for("index"))
    if request.method == "POST":
        nome   = request.form.get("nome", "").strip()
        email  = request.form.get("email", "").strip()
        senha  = request.form.get("senha", "")
        perfil = request.form.get("perfil", "usuario")
        if not nome or not email or not senha:
            flash("Todos os campos sao obrigatorios!", "erro")
            return redirect(url_for("novo_usuario"))
        if len(senha) < 6:
            flash("Senha: minimo 6 caracteres!", "erro")
            return redirect(url_for("novo_usuario"))
        conn   = conectar()
        cursor = conn.cursor()
        try:
            cursor.execute("""
                INSERT INTO usuarios (nome, email, senha_hash, perfil)
                VALUES (?, ?, ?, ?)
            """, (nome, email, generate_password_hash(senha), perfil))
            conn.commit()
            flash(f"Usuario '{nome}' criado!", "sucesso")
        except Exception:
            flash("E-mail ja cadastrado!", "erro")
        conn.close()
        return redirect(url_for("usuarios"))
    return render_template("novo_usuario.html")


@app.route("/usuarios/<int:id>/toggle")
@login_required
def toggle_usuario(id):
    if not current_user.is_admin:
        flash("Acesso negado!", "erro")
        return redirect(url_for("index"))
    if id == current_user.id:
        flash("Voce nao pode desativar sua conta!", "erro")
        return redirect(url_for("usuarios"))
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT ativo FROM usuarios WHERE id=?", (id,))
    u = cursor.fetchone()
    if u:
        novo = 0 if u["ativo"] else 1
        cursor.execute("UPDATE usuarios SET ativo=? WHERE id=?", (novo, id))
        conn.commit()
        flash(f"Usuario {'ativado' if novo else 'desativado'}!", "sucesso")
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
            flash("As senhas nao coincidem!", "erro")
            return redirect(url_for("minha_conta"))
        if len(nova_senha) < 6:
            flash("Senha: minimo 6 caracteres!", "erro")
            return redirect(url_for("minha_conta"))
        conn   = conectar()
        cursor = conn.cursor()
        cursor.execute("UPDATE usuarios SET senha_hash=? WHERE id=?",
            (generate_password_hash(nova_senha), current_user.id))
        conn.commit()
        conn.close()
        flash("Senha alterada!", "sucesso")
        return redirect(url_for("index"))
    return render_template("minha_conta.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)