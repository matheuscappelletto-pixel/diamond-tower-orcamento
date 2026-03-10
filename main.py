"""
DIAMOND TOWER — Automação Orçamentária
Roda via GitHub Actions:
1. Acessa Guarida e baixa extrato do mês anterior
2. Classifica lançamentos rigorosamente com Claude AI
3. Lança na planilha Google Sheets criando COMENTÁRIOS com o histórico
4. Envia e-mail com resumo
"""

import os
import re
import json
import time
import smtplib
import calendar
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ── Dependências ──────────────────────────────────────────────
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import anthropic
import gspread
from google.oauth2.service_account import Credentials
import unicodedata

# ── Configuração via variáveis de ambiente (GitHub Secrets) ───
GUARIDA_URL      = "https://agenciavirtual3.guarida.com.br/financeiro/extrato-condominio"
GUARIDA_USER     = os.environ["GUARIDA_USER"]
GUARIDA_PASS     = os.environ["GUARIDA_PASS"]
ANTHROPIC_KEY    = os.environ["ANTHROPIC_KEY"]
SPREADSHEET_ID   = os.environ["SPREADSHEET_ID"]   # ID do Google Sheets
GOOGLE_CREDS_JSON= os.environ["GOOGLE_CREDS_JSON"] # JSON da service account
NOTIFY_EMAIL     = os.environ.get("NOTIFY_EMAIL", "matheuscappelletto@gmail.com")
GMAIL_USER       = os.environ.get("GMAIL_USER", "")
GMAIL_PASS       = os.environ.get("GMAIL_PASS", "")

# ── Mapeamento mês → coluna na planilha ───────────────────────
MONTH_TO_COL = {
    (3,  2025): "F",
    (4,  2025): "G",
    (5,  2025): "H",
    (6,  2025): "I",
    (7,  2025): "J",
    (8,  2025): "K",
    (9,  2025): "L",
    (10, 2025): "M",
    (11, 2025): "N",
    (12, 2025): "O",
    (1,  2026): "P",
    (2,  2026): "Q",
    (3,  2026): "R",
}

# ── Mapeamento linhas da planilha ─────────────────────────────
ROW_MAP = {
    4:  {"empresa": "ALTO PADRAO",  "desc": "Alto Padrão — Mão de Obra (Monitoramento/Ronda/Porteiro/Limpeza/Serviços Gerais)",
         "keywords": ["alto padrao", "monitoramento", "ronda", "porteiro", "expedicao", "recepcao", "limpeza 03", "servicos gerais"]},
    11: {"empresa": "CAPPELLETTO",  "desc": "Cappelletto — Gestores presenciais",
         "keywords": ["cappelletto", "gestor"]},
    12: {"empresa": "GUARIDA",      "desc": "Guarida — Auxiliar Administração",
         "keywords": ["auxiliar administracao", "guarida administracao", "taxa de administracao"]},
    13: {"empresa": "ORGANICO",     "desc": "Orgânico — Oficial Manutenção 44h",
         "keywords": ["oficial manutencao", "marco aurelio", "marcos aurelio", "organico"]},
    14: {"empresa": "ORGANICO",     "desc": "Orgânico — Remuneração Subsíndico",
         "keywords": ["subsidico", "pro-labore subsidico", "remuneracao subsidico"]},
    17: {"empresa": "FG PISCINAS",  "desc": "FG Piscinas — Limpeza/Química",
         "keywords": ["fg piscinas", "piscina", "quimica"]},
    18: {"empresa": "ELITE",        "desc": "Elite — Elevadores Atlas",
         "keywords": ["elite", "elevador", "atlas"]},
    19: {"empresa": "STEMAC",       "desc": "Stemac — Gerador",
         "keywords": ["stemac", "gerador", "motor gerador"]},
    20: {"empresa": "BELLINI",      "desc": "Bellini — Consultoria Jurídica",
         "keywords": ["bellini", "juridic", "advocac"]},
    21: {"empresa": "AUDI PASTAS",  "desc": "Audi Pastas — Auditoria Externa",
         "keywords": ["audi pastas", "auditoria"]},
    24: {"empresa": "MULTIPLAN",    "desc": "Multiplan — Lixo/Entulho",
         "keywords": ["multiplan", "lixo", "entulho"]},
    25: {"empresa": "ALPINISMO",    "desc": "Alpinismo — Fachada",
         "keywords": ["alpinismo", "fachada"]},
    26: {"empresa": "DEMANDA",      "desc": "Manutenção Bombas Recalque",
         "keywords
