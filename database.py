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
            nome TEXT UNIQUE NOT NULL,
            descricao TEXT,
            faixa_etaria TEXT NOT NULL, -- 'criancas_0_3', 'criancas_4_7', 'criancas_8_12', 'adolescentes_13_15', 'jovens_16_17', 'adultos'
            ativa INTEGER DEFAULT 1
        )
    """)

    # Tabela de Disciplinas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS disciplinas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT UNIQUE NOT NULL,
            descricao TEXT,
            professor_id INTEGER,
            tem_atividades INTEGER DEFAULT 0, -- 0 para não, 1 para sim
            frequencia_minima REAL DEFAULT 75.0, -- % de frequência mínima
            ativa INTEGER DEFAULT 1,
            FOREIGN KEY (professor_id) REFERENCES usuarios (id)
        )
    """)

    # Tabela de Alunos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alunos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            data_nascimento TEXT,
            telefone TEXT,
            email TEXT,
            membro_igreja INTEGER DEFAULT 0, -- 0 para não, 1 para sim
            turma_id INTEGER,
            -- NOVOS CAMPOS
            nome_pai TEXT,
            nome_mae TEXT,
            endereco TEXT,
            -- FIM NOVOS CAMPOS
            FOREIGN KEY (turma_id) REFERENCES turmas (id)
        )
    """)

    # Tabela de Matrículas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS matriculas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            aluno_id INTEGER NOT NULL,
            disciplina_id INTEGER NOT NULL,
            data_inicio TEXT NOT NULL,
            data_conclusao TEXT,
            status TEXT DEFAULT 'cursando', -- 'cursando', 'aprovado', 'reprovado', 'trancado'
            nota1 REAL,
            nota2 REAL,
            participacao REAL, -- para adultos
            desafio REAL,      -- para adultos
            prova REAL,        -- para adultos
            meditacao REAL,    -- para adolescentes/jovens
            versiculos REAL,   -- para adolescentes/jovens
            desafio_nota REAL, -- para adolescentes/jovens (renomeado para evitar conflito)
            visitante REAL,    -- para adolescentes/jovens
            FOREIGN KEY (aluno_id) REFERENCES alunos (id),
            FOREIGN KEY (disciplina_id) REFERENCES disciplinas (id),
            UNIQUE (aluno_id, disciplina_id)
        )
    """)

    # Tabela de Presenças
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presencas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            matricula_id INTEGER NOT NULL,
            data_aula TEXT NOT NULL,
            presente INTEGER DEFAULT 0, -- 0 para falta, 1 para presente
            fez_atividade INTEGER DEFAULT 0, -- 0 para não, 1 para sim
            FOREIGN KEY (matricula_id) REFERENCES matriculas (id),
            UNIQUE (matricula_id, data_aula)
        )
    """)

    # --- Lógica de Migração para adicionar novas colunas ---
    # Adicionar nome_pai, nome_mae, endereco à tabela alunos se não existirem
    cursor.execute("PRAGMA table_info(alunos)")
    colunas_alunos = [col[1] for col in cursor.fetchall()]

    if 'nome_pai' not in colunas_alunos:
        cursor.execute("ALTER TABLE alunos ADD COLUMN nome_pai TEXT")
        print("Coluna 'nome_pai' adicionada à tabela 'alunos'.")
    if 'nome_mae' not in colunas_alunos:
        cursor.execute("ALTER TABLE alunos ADD COLUMN nome_mae TEXT")
        print("Coluna 'nome_mae' adicionada à tabela 'alunos'.")
    if 'endereco' not in colunas_alunos:
        cursor.execute("ALTER TABLE alunos ADD COLUMN endereco TEXT")
        print("Coluna 'endereco' adicionada à tabela 'alunos'.")

    # Adicionar desafio_nota à tabela matriculas se não existir (renomeado)
    cursor.execute("PRAGMA table_info(matriculas)")
    colunas_matriculas = [col[1] for col in cursor.fetchall()]

    if 'meditacao' not in colunas_matriculas:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN meditacao REAL")
        print("Coluna 'meditacao' adicionada à tabela 'matriculas'.")
    if 'versiculos' not in colunas_matriculas:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN versiculos REAL")
        print("Coluna 'versiculos' adicionada à tabela 'matriculas'.")
    if 'desafio_nota' not in colunas_matriculas: # Usando o nome corrigido
        cursor.execute("ALTER TABLE matriculas ADD COLUMN desafio_nota REAL")
        print("Coluna 'desafio_nota' adicionada à tabela 'matriculas'.")
    if 'visitante' not in colunas_matriculas:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN visitante REAL")
        print("Coluna 'visitante' adicionada à tabela 'matriculas'.")

    # Criar usuário admin padrão se não existir
    cursor.execute("SELECT id FROM usuarios WHERE email = 'admin@escola.com'")
    if not cursor.fetchone():
        senha_hash = generate_password_hash("admin123")
        cursor.execute(
            "INSERT INTO usuarios (nome, email, senha_hash, perfil) VALUES (?, ?, ?, ?)",
            ("Administrador", "admin@escola.com", senha_hash, "admin")
        )
        print("ADM criado: admin@escola.com / admin123")

    conn.commit()
    conn.close()
    print("Banco de dados inicializado/atualizado com sucesso!") # Adicionado para log

if __name__ == '__main__':
    inicializar_banco()