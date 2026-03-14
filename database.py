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
            faixa_etaria TEXT, -- 'criancas_0_3', 'criancas_4_7', 'criancas_8_12', 'adolescentes_13_15', 'jovens_16_17', 'adultos'
            ativa INTEGER DEFAULT 1
        )
    """)

    # --- MODIFICAÇÃO PARA ALUNOS ---
    try:
        cursor.execute("ALTER TABLE alunos RENAME TO alunos_old")
    except sqlite3.OperationalError as e:
        if "no such table" in str(e):
            pass # Ignora se a tabela não existe, pois será criada logo em seguida
        else:
            print(f"Aviso: Erro inesperado ao renomear alunos: {e}") # Imprime o erro mas não trava
    except Exception as e:
        print(f"Aviso: Erro inesperado ao renomear alunos: {e}")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS alunos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            telefone TEXT,
            email TEXT,
            data_nascimento TEXT,
            membro_igreja INTEGER DEFAULT 0,
            turma_id INTEGER,
            FOREIGN KEY (turma_id) REFERENCES turmas(id)
        )
    """)

    cursor.execute("PRAGMA table_info(alunos_old)")
    if cursor.fetchone():
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
        if "no such table" in str(e):
            pass
        else:
            print(f"Aviso: Erro inesperado ao renomear professores: {e}")
    except Exception as e:
        print(f"Aviso: Erro inesperado ao renomear professores: {e}")

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
    if cursor.fetchone():
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
            meditacao REAL,   -- Nova coluna
            versiculos REAL,  -- Nova coluna
            desafio_nota REAL, -- Nova coluna (renomeado para evitar conflito com 'desafio' antigo)
            visitante REAL,   -- Nova coluna
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

    # Lógica de migração para a coluna faixa_etaria na tabela turmas
    # Mapeia valores antigos para os novos
    try:
        cursor.execute("SELECT id, faixa_etaria FROM turmas WHERE faixa_etaria IN ('criancas', 'adolescentes_jovens', 'adolescentes_jovens_13_17', 'adolescentes_13_14', 'jovens_16_17', 'adultos', 'adultos_18_mais')")
        turmas_para_atualizar = cursor.fetchall()
        for turma_id, old_faixa_etaria in turmas_para_atualizar:
            new_faixa_etaria = old_faixa_etaria # Valor padrão se não houver mapeamento
            if old_faixa_etaria == 'criancas':
                new_faixa_etaria = 'criancas_8_12' # Define um padrão para crianças antigas
            elif old_faixa_etaria == 'adolescentes_jovens' or old_faixa_etaria == 'adolescentes_jovens_13_17' or old_faixa_etaria == 'adolescentes_13_14':
                new_faixa_etaria = 'adolescentes_13_15' # Mapeia para a nova categoria de adolescentes
            elif old_faixa_etaria == 'jovens_16_17':
                new_faixa_etaria = 'jovens_16_17' # Mantém jovens 16-17
            elif old_faixa_etaria == 'adultos' or old_faixa_etaria == 'adultos_18_mais':
                new_faixa_etaria = 'adultos' # Simplifica para 'adultos'

            if new_faixa_etaria != old_faixa_etaria:
                cursor.execute("UPDATE turmas SET faixa_etaria = ? WHERE id = ?", (new_faixa_etaria, turma_id))
        conn.commit()
    except sqlite3.OperationalError as e:
        print(f"Aviso: Erro ao migrar faixa_etaria em turmas (pode ser a primeira inicialização): {e}")
    except Exception as e:
        print(f"Erro inesperado durante a migração de faixa_etaria: {e}")


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
    # Novas colunas para Adolescentes/Jovens
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN meditacao REAL")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN versiculos REAL")
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN desafio_nota REAL") # Renomeado para evitar conflito
    except sqlite3.OperationalError:
        pass
    try:
        cursor.execute("ALTER TABLE matriculas ADD COLUMN visitante REAL")
    except sqlite3.OperationalError:
        pass


    # Inserir usuário admin padrão se não existir
    cursor.execute("SELECT COUNT(*) FROM usuarios WHERE email = 'admin@escola.com'")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO usuarios (nome,email,senha_hash,perfil)
            VALUES (?,?,?,'admin')
        """, ("Administrador", "admin@escola.com",
              generate_password_hash("admin123")))
        print("Adm criado: admin@escola.com / admin123")

    conn.commit()
    conn.close()