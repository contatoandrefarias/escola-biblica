import sqlite3
from werkzeug.security import generate_password_hash

DB_NAME = "escola_biblica.db"

def conectar():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def inicializar_banco():
    conn   = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            nome          TEXT NOT NULL,
            email         TEXT NOT NULL UNIQUE,
            senha_hash    TEXT NOT NULL,
            perfil        TEXT DEFAULT 'usuario',
            ativo         INTEGER DEFAULT 1,
            data_cadastro TEXT DEFAULT (date('now'))
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS turmas (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            nome          TEXT NOT NULL,
            descricao     TEXT,
            faixa_etaria  TEXT,
            ativa         INTEGER DEFAULT 1,
            data_cadastro TEXT DEFAULT (date('now'))
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS professores (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            nome          TEXT NOT NULL,
            telefone      TEXT,
            email         TEXT,
            especialidade TEXT,
            data_cadastro TEXT DEFAULT (date('now'))
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alunos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nome            TEXT NOT NULL,
            telefone        TEXT,
            email           TEXT,
            data_nascimento TEXT,
            membro_igreja   INTEGER DEFAULT 0,
            turma_id        INTEGER,
            data_cadastro   TEXT DEFAULT (date('now')),
            FOREIGN KEY (turma_id) REFERENCES turmas(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS disciplinas (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            nome              TEXT NOT NULL,
            descricao         TEXT,
            duracao_semanas   INTEGER DEFAULT 4,
            nota_minima       REAL DEFAULT 6.0,
            frequencia_minima REAL DEFAULT 75.0,
            tem_atividades    INTEGER DEFAULT 0,
            professor_id      INTEGER,
            ativa             INTEGER DEFAULT 1,
            FOREIGN KEY (professor_id) REFERENCES professores(id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS matriculas (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            aluno_id       INTEGER NOT NULL,
            disciplina_id  INTEGER NOT NULL,
            nota1          REAL,
            nota2          REAL,
            nota_final     REAL,
            status         TEXT DEFAULT 'cursando',
            data_inicio    TEXT DEFAULT (date('now')),
            data_conclusao TEXT,
            data_matricula TEXT DEFAULT (date('now')),
            FOREIGN KEY (aluno_id)      REFERENCES alunos(id),
            FOREIGN KEY (disciplina_id) REFERENCES disciplinas(id),
            UNIQUE(aluno_id, disciplina_id)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presencas (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula_id  INTEGER NOT NULL,
            data_aula     TEXT NOT NULL,
            presente      INTEGER DEFAULT 0,
            fez_atividade INTEGER DEFAULT 0,
            FOREIGN KEY (matricula_id) REFERENCES matriculas(id)
        )
    """)

    conn.commit()

    cursor.execute(
        "SELECT id FROM usuarios WHERE email=?",
        ("admin@escola.com",)
    )
    if not cursor.fetchone():
        cursor.execute("""
            INSERT INTO usuarios
                (nome, email, senha_hash, perfil)
            VALUES (?, ?, ?, ?)
        """, ("Administrador", "admin@escola.com",
              generate_password_hash("admin123"), "admin"))
        conn.commit()
        print("  Adm criado: admin@escola.com / admin123")

    conn.close()