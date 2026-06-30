# =============================================================================
# config.py — configuração mínima do add-in (API-only)
# -----------------------------------------------------------------------------
# Os segredos (token B3, API key FI) e o proxy vêm de VARIÁVEIS DE AMBIENTE.
# Ver apis.py / CLAUDE.md para os nomes das variáveis.
#
# ENV_PATH é apenas um fallback de desenvolvimento: caminho de um .env opcional
# na própria pasta. No banco esse arquivo NÃO existe e todos os valores vêm das
# variáveis de ambiente (apis.py trata a ausência sem erro).
# =============================================================================

import os

# Fallback de dev: por padrão procura um .env na própria pasta (não versionado).
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")

# Override de desenvolvimento LOCAL (fora do versionamento): para apontar a um
# .env FORA da pasta do repo, crie um config_local.py (gitignored) com:
#     ENV_PATH = r"C:\caminho\para\.env"
# No banco isso não existe e os segredos vêm das variáveis de ambiente.
try:
    from config_local import ENV_PATH  # noqa: F811
except Exception:
    pass
