from werkzeug.security import check_password_hash
from flask_login import UserMixin
from database import conectar

class Usuario(UserMixin):
    def __init__(self, id, nome, email, senha_hash, perfil):
        self.id = id
        self.nome = nome
        self.email = email
        self.senha_hash = senha_hash
        self.perfil = perfil

    def get_id(self):
        return str(self.id)

    @property
    def is_admin(self):
        return self.perfil == "admin"

    @property
    def is_aluno(self):
        return self.perfil == "aluno"

    @property
    def aluno_id(self):
        """Retorna o ID do aluno se o perfil for 'aluno', caso contrário None."""
        if self.is_aluno:
            conn = conectar()
            cursor = conn.cursor()
            # Assumindo que o email do usuário é o mesmo do aluno
            cursor.execute("SELECT id FROM alunos WHERE email = ?", (self.email,))
            aluno = cursor.fetchone()
            conn.close()
            return aluno['id'] if aluno else None
        return None


def carregar_usuario(user_id):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM usuarios WHERE id = ?", (user_id,))
    user_data = cursor.fetchone()
    conn.close()
    if user_data:
        return Usuario(user_data["id"], user_data["nome"],
                       user_data["email"], user_data["senha_hash"],
                       user_data["perfil"])
    return None

def verificar_login(email, senha):
    conn = conectar()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM usuarios WHERE email = ?", (email,))
    user_data = cursor.fetchone()
    conn.close()
    if user_data and check_password_hash(user_data["senha_hash"], senha):
        return Usuario(user_data["id"], user_data["nome"],
                       user_data["email"], user_data["senha_hash"],
                       user_data["perfil"])
    return None