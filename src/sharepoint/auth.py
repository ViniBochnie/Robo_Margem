import msal

def authenticate(client_id: str, authority: str, client_secret: str, username: str, password: str, scopes: list) -> str:
    """Autentica o usuário e retorna o token de acesso."""
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    result = app.acquire_token_by_username_password(
        username=username,
        password=password,
        scopes=scopes,
    )

    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Erro ao obter o token:", result.get("error"), result.get("error_description"))