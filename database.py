# ================================================
# database.py
# Banco de dados completo com tabela de usuários
# ================================================
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

    # ── Tabela de Usuários (LOGIN) ─────────────
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
    # perfil: 'admin' ou 'usuario'

    # ── Tabela de Professores ──────────────────
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

    # ── Tabela de Alunos ───────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alunos (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            nome            TEXT NOT NULL,
            telefone        TEXT,
            email           TEXT,
            data_nascimento TEXT,
            membro_igreja   INTEGER DEFAULT 0,
            data_cadastro   TEXT DEFAULT (date('now'))
        )
    """)

    # ── Tabela de Disciplinas ──────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS disciplinas (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            nome              TEXT NOT NULL,
            descricao         TEXT,
            carga_horaria     INTEGER DEFAULT 0,
            nota_minima       REAL DEFAULT 6.0,
            frequencia_minima REAL DEFAULT 75.0,
            professor_id      INTEGER,
            ativa             INTEGER DEFAULT 1,
            FOREIGN KEY (professor_id) REFERENCES professores(id)
        )
    """)

    # ── Tabela de Matrículas ───────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS matriculas (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            aluno_id       INTEGER NOT NULL,
            disciplina_id  INTEGER NOT NULL,
            nota1          REAL,
            nota2          REAL,
            nota_final     REAL,
            status         TEXT DEFAULT 'cursando',
            data_matricula TEXT DEFAULT (date('now')),
            FOREIGN KEY (aluno_id)      REFERENCES alunos(id),
            FOREIGN KEY (disciplina_id) REFERENCES disciplinas(id),
            UNIQUE(aluno_id, disciplina_id)
        )
    """)

    # ── Tabela de Presenças ────────────────────
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presencas (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula_id INTEGER NOT NULL,
            data_aula    TEXT NOT NULL,
            presente     INTEGER DEFAULT 0,
            FOREIGN KEY (matricula_id) REFERENCES matriculas(id)
        )
    """)

    conn.commit()

    # Criar usuário ADMIN padrão se não existir
    cursor.execute(
        "SELECT id FROM usuarios WHERE email = ?",
        ("admin@escola.com",)
    )
    if not cursor.fetchone():
        senha = generate_password_hash("admin123")
        cursor.execute("""
            INSERT INTO usuarios (nome, email, senha_hash, perfil)
            VALUES (?, ?, ?, ?)
        """, ("Administrador", "admin@escola.com", senha, "admin"))
        conn.commit()
        print("  ✅ Usuário admin criado!")
        print("  📧 Email: admin@escola.com")
        print("  🔑 Senha: admin123")
        print("  ⚠️  TROQUE A SENHA APÓS O PRIMEIRO LOGIN!\n")

    conn.close()