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
            faixa_etaria TEXT, -- 'criancas', 'adolescentes_jovens', 'adultos'
            ativa INTEGER DEFAULT 1
        )
    """)

    # --- MODIFICAÇÃO PARA ALUNOS ---
    try:
        cursor.execute("ALTER TABLE alunos RENAME TO alunos_old")
    except sqlite3.OperationalError as e:
        # Apenas ignora se a tabela não existe. Não re-lança o erro.
        if "no such table" not in str(e):
            raise e # Re-lança outros erros inesperados
    except Exception as e:
        # Captura qualquer outro erro inesperado durante o rename
        print(f"Aviso: Erro ao tentar renomear tabela 'alunos': {e}")


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

    # Copiar os dados da tabela antiga para a nova, se a tabela antiga existir
    cursor.execute("PRAGMA table_info(alunos_old)")
    if cursor.fetchone(): # Se alunos_old existe
        cursor.execute("""
            INSERT INTO alunos (id, nome, telefone, email, data_nascimento, membro_igreja, turma_id)
            SELECT id, nome, telefone, email, data_nascimento, membro_igreja, turma_id
            FROM alunos_old
        """)
        cursor.execute("DROP TABLE alunos_old")
    # --- FIM DA MODIFICAÇÃO PARA ALUNOS ---


    # --- MODIFICAÇÃO PARA PROFESSORES ---
    try:
        cursor.execute("ALTER TABLE professores RENAME TO professores_old")
    except sqlite3.OperationalError as e:
        # Apenas ignora se a tabela não existe. Não re-lança o erro.
        if "no such table" not in str(e):
            raise e # Re-lança outros erros inesperados
    except Exception as e:
        # Captura qualquer outro erro inesperado durante o rename
        print(f"Aviso: Erro ao tentar renomear tabela 'professores': {e}")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS professores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT,
            email TEXT,
            especialidade TEXT
        )
    """)

    cursor.execute("PRAGMA table_info(professores_old)")
    if cursor.fetchone(): # Se professores_old existe
        cursor.execute("""
            INSERT INTO professores (id, nome, telefone, email, especialidade)
            SELECT id, nome, telefone, email, especialidade
            FROM professores_old
        """)
        cursor.execute("DROP TABLE professores_old")
    # --- FIM DA MODIFICAÇÃO PARA PROFESSORES ---


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
            participacao REAL,
            desafio REAL,
            prova REAL,
            status TEXT DEFAULT 'cursando',
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
        pass
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN frequencia_minima REAL DEFAULT 75.0")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN ativa INTEGER DEFAULT 1")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE disciplinas ADD COLUMN tem_atividades INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    # Colunas para turmas
    try:
        cursor.execute("ALTER TABLE turmas ADD COLUMN ativa INTEGER DEFAULT 1")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE turmas ADD COLUMN faixa_etaria TEXT DEFAULT 'adultos'")
    except sqlite3.OperationalError:
        pass

    # Colunas para matriculas
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN nota_final REAL")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN status TEXT DEFAULT 'cursando'")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN participacao REAL")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN desafio REAL")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN prova REAL")
    except sqlite3.OperationalError:
        pass

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