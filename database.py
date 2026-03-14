import sqlite3
from werkzeug.security import generate_password_hash

DATABASE = 'escola.db'

def conectar():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def inicializar_banco():
    conn = conectar()
    cursor = conn.cursor()

    # Tabela de Usuários
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL,
            perfil TEXT NOT NULL DEFAULT 'usuario' -- 'admin', 'professor', 'aluno'
        )
    """)

    # Tabela de Turmas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS turmas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            descricao TEXT,
            faixa_etaria TEXT,
            ativa INTEGER DEFAULT 1
        )
    """)

    # --- INÍCIO DA MODIFICAÇÃO PARA ALUNOS ---
    # 1. Renomear a tabela alunos existente (se houver)
    try:
        cursor.execute("ALTER TABLE alunos RENAME TO alunos_old")
    except sqlite3.OperationalError as e:
        # Se a tabela não existir ou já tiver sido renomeada, ignora
        if "no such table" not in str(e) and "already exists" not in str(e):
            raise e # Re-lança outros erros inesperados

    # 2. Criar a nova tabela de Alunos SEM a restrição UNIQUE no email
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alunos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT,
            email TEXT, -- Removido UNIQUE
            data_nascimento TEXT,
            membro_igreja INTEGER DEFAULT 0,
            turma_id INTEGER,
            FOREIGN KEY (turma_id) REFERENCES turmas(id)
        )
    """)

    # 3. Copiar os dados da tabela antiga para a nova
    # Verifica se a tabela antiga existe e tem dados
    cursor.execute("PRAGMA table_info(alunos_old)")
    if cursor.fetchone(): # Se alunos_old existe
        cursor.execute("""
            INSERT INTO alunos (id, nome, telefone, email, data_nascimento, membro_igreja, turma_id)
            SELECT id, nome, telefone, email, data_nascimento, membro_igreja, turma_id
            FROM alunos_old
        """)
        # 4. Excluir a tabela antiga
        cursor.execute("DROP TABLE alunos_old")
    # --- FIM DA MODIFICAÇÃO PARA ALUNOS ---


    # Tabela de Professores
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS professores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT,
            email TEXT UNIQUE,
            especialidade TEXT
        )
    """)

    # Tabela de Disciplinas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS disciplinas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL UNIQUE,
            descricao TEXT,
            duracao_semanas INTEGER DEFAULT 4,
            nota_minima REAL DEFAULT 6.0,
            frequencia_minima REAL DEFAULT 75.0, -- Em porcentagem
            tem_atividades INTEGER DEFAULT 0,
            professor_id INTEGER,
            ativa INTEGER DEFAULT 1,
            FOREIGN KEY (professor_id) REFERENCES professores(id)
        )
    """)

    # Tabela de Matrículas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS matriculas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            aluno_id INTEGER NOT NULL,
            disciplina_id INTEGER NOT NULL,
            data_inicio TEXT,
            data_conclusao TEXT,
            nota1 REAL,
            nota2 REAL,
            nota_final REAL,
            status TEXT DEFAULT 'cursando', -- 'cursando', 'aprovado', 'reprovado'
            UNIQUE(aluno_id, disciplina_id),
            FOREIGN KEY (aluno_id) REFERENCES alunos(id),
            FOREIGN KEY (disciplina_id) REFERENCES disciplinas(id)
        )
    """)

    # Tabela de Presenças
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presencas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula_id INTEGER NOT NULL,
            data_aula TEXT NOT NULL,
            presente INTEGER DEFAULT 0,
            fez_atividade INTEGER DEFAULT 0,
            UNIQUE(matricula_id, data_aula),
            FOREIGN KEY (matricula_id) REFERENCES matriculas(id)
        )
    """)

    # Adicionar colunas se não existirem (para atualizações de banco)
    # Colunas para disciplinas
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN nota_minima REAL DEFAULT 6.0")
    except sqlite3.OperationalError:
        pass # Coluna já existe
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN frequencia_minima REAL DEFAULT 75.0")
    except sqlite3.OperationalError:
        pass # Coluna já existe
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN ativa INTEGER DEFAULT 1")
    except sqlite3.OperationalError:
        pass # Coluna já existe
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN tem_atividades INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass # Coluna já existe

    # Colunas para turmas
    try:
        cursor.execute("ALTER TABLE turmas ADD COLUMN ativa INTEGER DEFAULT 1")
    except sqlite3.OperationalError:
        pass # Coluna já existe

    # Colunas para matriculas
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN nota_final REAL")
    except sqlite3.OperationalError:
        pass # Coluna já existe
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN status TEXT DEFAULT 'cursando'")
    except sqlite3.OperationalError:
        pass # Coluna já existe

    # Inserir usuário admin padrão se não existir
    cursor.execute("SELECT COUNT(*) FROM usuarios WHERE email = 'admin@escola.com'")
    if cursor.fetchone()[0] == 0:
        admin_senha_hash = generate_password_hash("admin123")
        cursor.execute(
            "INSERT INTO usuarios (nome, email, senha_hash, perfil) VALUES (?, ?, ?, ?)",
            ("Administrador", "admin@escola.com", admin_senha_hash, "admin")
        )
        print("  Adm criado: admin@escola.com / admin123")

    conn.commit()
    conn.close()