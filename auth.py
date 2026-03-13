# ================================================
# auth.py
# Gerenciamento de autenticação e usuários
# ================================================
from flask_login import UserMixin
from database import conectar
from werkzeug.security import check_password_hash


class Usuario(UserMixin):
    """
    Classe que representa um usuário logado.
    UserMixin fornece métodos padrão do Flask-Login:
    is_authenticated, is_active, get_id()
    """
    def __init__(self, id, nome, email, perfil, ativo):
        self.id    = id
        self.nome  = nome
        self.email = email
        self.perfil = perfil
        self.ativo  = ativo

    def get_id(self):
        return str(self.id)

    @property
    def is_admin(self):
        """Verifica se é administrador."""
        return self.perfil == "admin"

    @property
    def is_active(self):
        """Usuário ativo pode fazer login."""
        return bool(self.ativo)


def carregar_usuario(user_id):
    """
    Carrega usuário pelo ID.
    Chamado automaticamente pelo Flask-Login
    a cada requisição.
    """
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT * FROM usuarios WHERE id = ?",
        (int(user_id),)
    )
    u = cursor.fetchone()
    conn.close()

    if u:
        return Usuario(
            id     = u["id"],
            nome   = u["nome"],
            email  = u["email"],
            perfil = u["perfil"],
            ativo  = u["ativo"]
        )
    return None


def verificar_login(email, senha):
    """
    Verifica se email e senha estão corretos.
    Retorna o usuário ou None.
    """
    conn   = conectar()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT * FROM usuarios WHERE email = ? AND ativo = 1",
        (email,)
    )
    u = cursor.fetchone()
    conn.close()

    if u and check_password_hash(u["senha_hash"], senha):
        return Usuario(
            id     = u["id"],
            nome   = u["nome"],
            email  = u["email"],
            perfil = u["perfil"],
            ativo  = u["ativo"]
        )
    return None