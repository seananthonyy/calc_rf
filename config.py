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

ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
