"""
DIAMOND TOWER — Automação Orçamentária
Roda via GitHub Actions:
1. Acessa Guarida (Login 2 Etapas / Next.js) e baixa extrato
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
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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
         "keywords": ["bomba", "recalque"]},
    27: {"empresa": "",             "desc": "Manutenção Interfones",
         "keywords": ["interfone"]},
    28: {"empresa": "",             "desc": "Manutenção Predial",
         "keywords": ["manutencao predial", "predial"]},
    29: {"empresa": "",             "desc": "Incêndio",
         "keywords": ["incendio"]},
    30: {"empresa": "",             "desc": "Obrigações Legais",
         "keywords": ["obrigacao legal", "certificado", "laudo", "spda", "extintores recarga"]},
    31: {"empresa": "",             "desc": "Controle Pragas / Caixa D'água",
         "keywords": ["praga", "dedetiz", "caixa d agua"]},
    32: {"empresa": "",             "desc": "Paisagismo",
         "keywords": ["paisagismo"]},
    35: {"empresa": "",             "desc": "Chaveiro",
         "keywords": ["chaveiro"]},
    36: {"empresa": "",             "desc": "Lâmpadas",
         "keywords": ["lampada"]},
    37: {"empresa": "",             "desc": "Obras Diversas",
         "keywords": ["alvenaria", "serralheria", "obra"]},
    38: {"empresa": "",             "desc": "Material Ferragem/Elétrica/Hidráulica",
         "keywords": ["ferragem", "eletrica", "hidraulica", "ferrament"]},
    39: {"empresa": "",             "desc": "Material Expediente",
         "keywords": ["expediente", "escritorio"]},
    40: {"empresa": "",             "desc": "Material Limpeza",
         "keywords": ["material limpeza", "consumivel"]},
    41: {"empresa": "ELITE",        "desc": "Elite — Peças Elevador",
         "keywords": ["peca elevador", "pecas elevador"]},
    42: {"empresa": "DEMANDA",      "desc": "Móveis e Utensílios",
         "keywords": ["movel", "utensilio"]},
    43: {"empresa": "",             "desc": "Peças Ar Condicionado",
         "keywords": ["ar condicionado", "exaustor"]},
    44: {"empresa": "",             "desc": "Peças Extintores/Hidrantes",
         "keywords": ["extintor", "hidrante"]},
    45: {"empresa": "",             "desc": "Peças Gerador / Óleo Diesel",
         "keywords": ["grupo gerador", "oleo diesel", "diesel"]},
    46: {"empresa": "",             "desc": "Peças Catracas",
         "keywords": ["catraca"]},
    49: {"empresa": "",             "desc": "Telefonia/Internet",
         "keywords": ["telefon", "internet", "claro", "vivo"]},
    50: {"empresa": "",             "desc": "Correio/Motoboy",
         "keywords": ["correio", "motoboy", "sedex"]},
    51: {"empresa": "",             "desc": "Impostos DIRF/DARF/ISS/PIS/COFINS",
         "keywords": ["dirf", "darf", "issqn", "pis", "cofins", "imposto", "inss"]},
    52: {"empresa": "",             "desc": "Assembleia",
         "keywords": ["assembleia"]},
    53: {"empresa": "",             "desc": "Custas Judiciais",
         "keywords": ["custas", "judicial", "processo"]},
    54: {"empresa": "GOOGLE",       "desc": "Google — Hospedagem/Domínio",
         "keywords": ["google", "hospedagem", "dominio"]},
    59: {"empresa": "",             "desc": "Água e Esgoto",
         "keywords": ["agua", "esgoto", "corsan", "dmae"]},
    60: {"empresa": "",             "desc": "Energia Elétrica",
         "keywords": ["energia", "eletrica", "ceee"]},
    61: {"empresa": "",             "desc": "Seguro",
         "keywords": ["seguro"]},
    65: {"empresa": "",             "desc": "Despesas Reembolsáveis",
         "keywords": ["reembolso", "reembolsavel"]},
    66: {"empresa": "",             "desc": "Honorários Advocatícios",
         "keywords": ["honorario", "advocaticio"]},
    67: {"empresa": "",             "desc": "Ampliação CFTV",
         "keywords": ["cftv", "camera"]},
    68: {"empresa": "",             "desc": "Transferências Entre Contas",
         "keywords": ["transferencia entre contas", "aplicacao"]}
}

# ── Helpers ───────────────────────────────────────────────────

def norm(texto):
    """Remove acentos e coloca em minúsculas para comparação."""
    t = unicodedata.normalize("NFKD", str(texto).lower())
    return "".join(c for c in t if not unicodedata.combining(c))

def mes_anterior():
    hoje = date.today()
    primeiro_do_mes = hoje.replace(day=1)
    ultimo_mes = primeiro_do_mes - relativedelta(months=1)
    return ultimo_mes.month, ultimo_mes.year

def ultimo_dia(mes, ano):
    return calendar.monthrange(ano, mes)[1]

def col_para_indice(col_letra):
    """F→5, G→6, etc. (1-indexed para gspread)"""
    return ord(col_letra.upper()) - ord('A') + 1

# ── 1. COLETAR EXTRATO DA GUARIDA (REACT SAFE) ────────────────

def coletar_extrato(mes, ano):
    print(f"\n[1/4] Acessando Guarida — extrato {mes:02d}/{ano}")

    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(options=opts)
    wait = WebDriverWait(driver, 30) # Maior tolerância
    lancamentos = []

    try:
        print("   Acessando página de login...")
        driver.get("https://agenciavirtual3.guarida.com.br/login")
        time.sleep(6) # Deixa a página dar 'hydrate'

        print("   Fazendo login (Etapas React.js Human-Like)...")
        try:
            # 1. Email + TAB + ENTER
            campo_email = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[type='email'], input[name='email']")
            ))
            campo_email.click()
            time.sleep(0.5)
            campo_email.send_keys(GUARIDA_USER)
            time.sleep(1)
            campo_email.send_keys(Keys.TAB)
            time.sleep(0.5)
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            
            time.sleep(5) # Delay para a tela de senha aparecer sem mudar a página

            # 2. Senha + TAB + ENTER
            campo_senha = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "input[type='password']")
            ))
            campo_senha.click()
            time.sleep(0.5)
            campo_senha.send_keys(GUARIDA_PASS)
            time.sleep(1)
            campo_senha.send_keys(Keys.TAB)
            time.sleep(0.5)
            ActionChains(driver).send_keys(Keys.ENTER).perform()

            time.sleep(10) # Tempo ocioso para logar e abrir o Dashboard (a imagem com os cartões)
        except Exception as e:
            print(f"   Aviso no Login Híbrido: {str(e)[:100]}. Testando clique Javascript direto...")
            try:
                # Fallback brutal se o "Abo" (TAB) não tiver ido
                driver.execute_script("document.querySelector('button').click()")
                time.sleep(6)
            except:
                pass

        print("   Navegando do Dashboard para o Extrato...")
        try:
            # Na sua tela a URL correta depois do login é essa:
            driver.get("https://agenciavirtual3.guarida.com.br/financeiro/extrato-condominio")
            time.sleep(8)
        except Exception as e:
            print("   Erro na navegação do Extrato. Acesso falhou.")

        data_inicio = f"01/{mes:02d}/{ano}"
        data_fim    = f"{ultimo_dia(mes, ano):02d}/{mes:02d}/{ano}"
        print(f"   Período: {data_inicio} a {data_fim}")

        try:
            # Seleção mecânica do período do Extrato
            inputs_data = driver.find_elements(By.CSS_SELECTOR, "input[type='date'], input[id*='data'], input[id*='date']")
            if len(inputs_data) >= 2:
                # Limpa e digita usando injeção Javascript (100% à prova de falhas de foco HTML)
                driver.execute_script(f"arguments[0].value = '{ano}-{mes:02d}-01'; arguments[0].dispatchEvent(new Event('change'));", inputs_data[0])
                driver.execute_script(f"arguments[0].value = '{ano}-{mes:02d}-{ultimo_dia(mes,ano):02d}'; arguments[0].dispatchEvent(new Event('change'));", inputs_data[1])
                time.sleep(2)

            btn_filtro = driver.find_element(By.XPATH, "//*[contains(text(),'Pesquisar') or contains(text(),'Filtrar') or contains(text(),'Buscar')]")
            driver.execute_script("arguments[0].click();", btn_filtro)
            time.sleep(6)
        except Exception as e:
            print(f"   Aviso: Seleção de data falhou, mas procuraremos capturar dados mesmo assim.")

        print("   Coletando dados da tabela de despesas...")
        lancamentos = _scrape_lancamentos(driver, mes, ano, ignorar_mes_ano=False)
        
        if not lancamentos:
            print("   Tentando capturar elementos ignorando validação do mês...")
            lancamentos = _scrape_lancamentos(driver, mes, ano, ignorar_mes_ano=True)

        print(f"   {len(lancamentos)} débitos coletados.")

    finally:
        driver.quit()

    return lancamentos


def _scrape_lancamentos(driver, mes, ano, ignorar_mes_ano=False):
    lancamentos = []
    pagina_texto = driver.page_source

    re_data_extenso = r'(\d{2}/\d{2}/\d{4})'
    re_valor_negativo = r'-\s*([\d.]+,\d{2})'

    linhas = driver.find_elements(By.CSS_SELECTOR, "table tr, .extrato-linha, .item-extrato, [class*='extrato']")
    for linha in linhas:
        texto = linha.text.strip()
        if not texto or len(texto) < 10: continue

        m_data = re.search(re_data_extenso, texto)
        if not m_data: continue

        data_str = m_data.group(1)
        partes = data_str.split('/')
        if len(partes) != 3: continue
        d, m, a = int(partes[0]), int(partes[1]), int(partes[2])
        
        if not ignorar_mes_ano:
            if m != mes or a != ano: continue

        m_valor = re.search(re_valor_negativo, texto)
        if not m_valor: continue

        val_str = m_valor.group(1).replace('.', '').replace(',', '.')
        try:
            valor = float(val_str)
        except:
            continue

        desc = texto[m_data.end():texto.rfind(m_valor.group(0))].strip()
        desc = re.sub(r'\s+', ' ', desc).strip()
        if not desc: desc = texto[:80]

        if valor > 0 and desc:
            lancamentos.append({
                "data":      data_str,
                "descricao": desc,
                "valor":     valor,
            })

    # Fallback caso a tabela HTML não abra corretamente
    if not lancamentos:
        for m in re.finditer(r'(\d{2}/\d{2}/\d{4})\s+([\w\s\d./:,()#\-]+?)\s+-\s*([\d.,]+)', pagina_texto):
            data_str, desc, val_str = m.group(1), m.group(2).strip(), m.group(3)
            partes = data_str.split('/')
            d, mes_val, ano_val = int(partes[0]), int(partes[1]), int(partes[2])
            
            if not ignorar_mes_ano:
                if mes_val != mes or ano_val != ano: continue
                
            val_str = val_str.replace('.', '').replace(',', '.')
            try:
                valor = float(val_str)
                if valor > 0 and len(desc) > 5:
                    lancamentos.append({"data": data_str, "descricao": desc, "valor": valor})
            except:
                pass

    vistos = set()
    unicos = []
    for l in lancamentos:
        chave = (l["data"], l["descricao"][:30], round(l["valor"], 2))
        if chave not in vistos:
            vistos.add(chave)
            unicos.append(l)

    return unicos

# ── 2. CLASSIFICAR COM CLAUDE AI ──────────────────────────────

def classificar_com_claude(lancamentos):
    print(f"\n[2/4] Classificando {len(lancamentos)} lançamentos com Claude AI...")
    client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)

    categorias = "\n".join([
        f"Linha {num}: {info['empresa'] + ' — ' if info['empresa'] else ''}{info['desc']} (Keywords: {', '.join(info['keywords'])})"
        for num, info in ROW_MAP.items()
    ])

    lista_formatada = "\n".join([
        f"{i+1}. Data: {l['data']} | Valor: R$ {l['valor']:.2f} | Descrição: '{l['descricao']}'"
        for i, l in enumerate(lancamentos)
    ])

    prompt = f"""Você é o analista financeiro principal do 'Diamond Tower'. Sua tarefa é mapear os débitos extraídos do extrato bancário (da imobiliária Guarida) para o plano de contas em Excel do condomínio.

CATEGORIAS (PLANO DE CONTAS) E SUAS LINHAS:
{categorias}

Linha 999: NÃO CLASSIFICADA / DÚVIDA

REGRA DE PRECEDÊNCIA (LEIA COM MUITA ATENÇÃO):
1. SE a descrição contiver "MARCO AURELIO" ou "OFICIAL DE MANUTENCAO", é a LINHA 13 (independentemente de dizer 'PG SALARIO').
2. SE a descrição contiver "PRO-LABORE SUBISIDICO" ou "SUBSINDICO", é a LINHA 14.
3. SE a descrição for sobre "ALTO PADRAO" (monitoramento, ronda, vigilante, porteiro, expedição, limpeza), é a LINHA 4.
4. Taxa de Administração da GUARIDA vai para a LINHA 12.
5. Pagamento à CAPPELLETTO Gestores vai para LINHA 11.
6. Tarifas bancárias de transferência, pagamento ou não reconhecidas vão para Linha 999 (para revisar).
7. Impostos federais (DARF, ISS, PIS, COFINS, INSS) vão para a Linha 51.

LANÇAMENTOS DO MÊS:
{lista_formatada}

Você DEVE responder ESTRITAMENTE num formato JSON VÁLIDO contendo APENAS a chave "classificacoes".
EXEMPLO ESTRITO:
{{
  "classificacoes": [
    {{"n": 1, "linha": 13, "motivo": "Menciona Marco Aurelio, que é o Oficial de Manutenção."}},
    {{"n": 2, "linha": 999, "motivo": "Taxa DOC não possui categoria clara."}}
  ]
}}
NÃO ADICIONE TEXTO FORA DO JSON."""

    try:
        resp = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        texto = resp.content[0].text.strip()
        texto = re.sub(r'```json\n?|\n?```', '', texto).strip()
        resultado = json.loads(texto)

        classif = {}
        for c in resultado.get("classificacoes", []):
            idx = c["n"] - 1 # n-1 para ser index 0 da nossa lista
            if 0 <= idx < len(lancamentos):
                classif[idx] = {"linha": c["linha"], "motivo": c.get("motivo", "")}

        for i in range(len(lancamentos)):
            if i not in classif:
                linha, _ = _classificar_keywords(lancamentos[i]["descricao"])
                classif[i] = {"linha": linha, "motivo": "Fallback Keywords"}

        return classif

    except Exception as e:
        print(f"   Erro Claude API (Retorno Inválido): {e}")
        return {i: {"linha": _classificar_keywords(l["descricao"])[0], "motivo": "Fallback Keywords geral"}
                for i, l in enumerate(lancamentos)}

def _classificar_keywords(descricao):
    desc = norm(descricao)
    melhor, score = 999, 0
    for num, info in ROW_MAP.items():
        for kw in info["keywords"]:
            if norm(kw) in desc and len(kw) > score:
                score = len(kw)
                melhor = num
    return melhor, score

# ── 3. LANÇAR NO GOOGLE SHEETS COM COMENTÁRIOS E HISTÓRICO ────

def lancar_no_sheets(lancamentos, classificacoes, mes, ano):
    if not lancamentos: return [], []
    print(f"\n[3/4] Lançando valores e inserindo comentários na planilha Google Sheets...")

    creds_dict = json.loads(GOOGLE_CREDS_JSON)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    aba_nome = f"PF {ano if mes >= 3 else ano - 1}"
    try:
        ws = sh.worksheet(aba_nome)
    except:
        abas = [w.title for w in sh.worksheets()]
        pf_abas = [a for a in abas if a.startswith("PF")]
        if not pf_abas: raise Exception("Nenhuma aba PF encontrada!")
        aba_nome = sorted(pf_abas)[-1]
        ws = sh.worksheet(aba_nome)

    col_letra = MONTH_TO_COL.get((mes, ano))
    if not col_letra: raise Exception(f"Mês {mes}/{ano} não mapeado para coluna!")
    
    col_idx = col_para_indice(col_letra)
    sheet_id = ws.id 

    por_linha = {}
    nao_classificados = []

    for i, lanc in enumerate(lancamentos):
        linha = classificacoes.get(i, {}).get("linha", 999)
        motivo = classificacoes.get(i, {}).get("motivo", "")

        if 5 <= linha <= 10: linha = 4

        if linha == 999 or linha not in ROW_MAP:
            nao_classificados.append(lanc)
            continue

        if linha not in por_linha:
            por_linha[linha] = {"total": 0, "itens": []}
            
        por_linha[linha]["total"] += lanc["valor"]
        
        texto_historico = f"{lanc['data']} - {lanc['descricao']} | R$ {lanc['valor']:.2f}"
        por_linha[linha]["itens"].append(texto_historico)

    resumo = []
    atualizacoes_valores = []
    atualizacoes_notas = [] 

    for linha, dados in por_linha.items():
        cel_ref = f"{col_letra}{linha}"

        try:
            val_atual_str = ws.cell(linha, col_idx).value or "0"
            val_atual_str = str(val_atual_str).replace("R$", "").replace(".", "").replace(",", ".").strip()
            val_atual = float(val_atual_str) if val_atual_str and val_atual_str != "-" else 0.0
        except:
            val_atual = 0.0

        novo_val = round(val_atual + dados["total"], 2)
        novo_val_fmt = f"{novo_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        atualizacoes_valores.append({
            "range": cel_ref,
            "values": [[novo_val_fmt]]
        })

        nota_texto = "LANÇAMENTOS DA GUARIDA:\n" + "\n".join(dados["itens"])
        
        atualizacoes_notas.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": linha - 1,
                    "endRowIndex": linha,
                    "startColumnIndex": col_idx - 1,
                    "endColumnIndex": col_idx
                },
                "rows": [
                    {
                        "values": [
                            {
                                "note": nota_texto 
                            }
                        ]
                    }
                ],
                "fields": "note" 
            }
        })

        resumo.append({
            "celula":  cel_ref,
            "desc":    ROW_MAP[linha]["desc"][:50],
            "valor":   novo_val,
            "itens":   dados["itens"],
        })

    if atualizacoes_valores:
        ws.batch_update(atualizacoes_valores)
        sh.batch_update({"requests": atualizacoes_notas})
        print(f"   {len(atualizacoes_valores)} células (e seus comentários) atualizadas na Planilha Principal!")
    else:
        print("   Nenhum valor para atualizar.")

    return resumo, nao_classificados

# ── 4. ENVIAR E-MAIL DE RESUMO ────────────────────────────────

def enviar_email(mes, ano, resumo, nao_classificados, lancamentos_total):
    if not GMAIL_USER or not GMAIL_PASS:
        print("\n[4/4] E-mail não configurado — pulando.")
        return

    print(f"\n[4/4] Enviando e-mail para {NOTIFY_EMAIL}...")

    nomes_meses = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                   "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

    total_lancado = sum(r["valor"] for r in resumo)
    total_fmt = f"R$ {total_lancado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    linhas_html = ""
    for r in resumo:
        val_fmt = f"R$ {r['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        detalhes = "<br>".join([f"&nbsp;&nbsp;• {it}" for it in r["itens"]])
        linhas_html += f"""
        <tr>
          <td style='padding:6px 10px;color:#888;font-family:monospace'>{r['celula']}</td>
          <td style='padding:6px 10px'>{r['desc']}</td>
          <td style='padding:6px 10px;color:#4caf7d;font-family:monospace;text-align:right'>{val_fmt}</td>
        </tr>
        <tr><td colspan='3' style='padding:0 10px 8px;font-size:12px;color:#666'>{detalhes}</td></tr>
        """

    nao_class_html = ""
    if nao_classificados:
        itens = "".join([f"<li>{l['data']} | {l['descricao']} | R$ {l['valor']:.2f}</li>" for l in nao_classificados])
        nao_class_html = f"""
        <div style='margin-top:20px;padding:12px;background:#2a1a1a;border-left:3px solid #c94c4c;border-radius:4px'>
          <strong style='color:#c94c4c'>⚠ {len(nao_classificados)} lançamento(s) não classificado(s) — verificar manualmente:</strong>
          <ul style='margin-top:8px;color:#ccc'>{itens}</ul>
        </div>"""

    html = f"""
    <div style='background:#0f0f0f;color:#e8e0d0;font-family:sans-serif;padding:30px;max-width:700px;margin:0 auto'>
      <div style='border-bottom:1px solid #333;padding-bottom:15px;margin-bottom:20px'>
        <h2 style='color:#c9a84c;margin:0;font-size:22px'>◈ Diamond Tower</h2>
        <p style='color:#666;margin:4px 0 0'>Lançamento automático — {nomes_meses[mes]} {ano}</p>
      </div>
      <div style='background:#181818;border:1px solid #2a2a2a;border-radius:6px;padding:15px;margin-bottom:20px'>
        <div style='display:flex;justify-content:space-between'>
          <span style='color:#888'>Lançamentos processados</span>
          <strong>{lancamentos_total}</strong>
        </div>
        <div style='display:flex;justify-content:space-between;margin-top:8px'>
          <span style='color:#888'>Total lançado</span>
          <strong style='color:#c9a84c;font-size:18px'>{total_fmt}</strong>
        </div>
      </div>
      <table style='width:100%;border-collapse:collapse;background:#181818;border:1px solid #2a2a2a;border-radius:6px'>
        <thead>
          <tr style='border-bottom:1px solid #2a2a2a'>
            <th style='padding:8px 10px;text-align:left;color:#666;font-size:12px'>CÉLULA</th>
            <th style='padding:8px 10px;text-align:left;color:#666;font-size:12px'>DESCRIÇÃO</th>
            <th style='padding:8px 10px;text-align:right;color:#666;font-size:12px'>VALOR</th>
          </tr>
        </thead>
        <tbody>{linhas_html}</tbody>
      </table>
      {nao_class_html}
      <p style='color:#444;font-size:11px;margin-top:25px'>
        Lançado automaticamente em {datetime.now().strftime('%d/%m/%Y às %H:%M')} UTC<br>
        <i>As notas e comentários foram salvos com sucesso na planilha.</i>
      </p>
    </div>
    """

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"✓ Diamond Tower — Extrato {nomes_meses[mes]}/{ano} lançado"
    msg["From"]    = GMAIL_USER
    msg["To"]      = NOTIFY_EMAIL
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_PASS)
            smtp.send_message(msg)
        print("   E-mail enviado!")
    except Exception as e:
        print(f"   Erro ao enviar e-mail: {e}")

# ── MAIN ──────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  DIAMOND TOWER — Automação Orçamentária")
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} UTC")
    print("=" * 55)

    mes, ano = mes_anterior()
    nomes = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
             "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    print(f"\nProcessando: {nomes[mes]} {ano}")

    lancamentos = coletar_extrato(mes, ano)
    if not lancamentos:
        print("\n⚠ Nenhum lançamento encontrado. Verifique o acesso ao site.")
        return

    classificacoes = classificar_com_claude(lancamentos)

    resumo, nao_classificados = lancar_no_sheets(lancamentos, classificacoes, mes, ano)

    enviar_email(mes, ano, resumo, nao_classificados, len(lancamentos))

    print("\n✓ Concluído!")
    for r in resumo:
        val_fmt = f"R$ {r['valor']:,.2f}".replace(",","X").replace(".",",").replace("X",".")
        print(f"  {r['celula']:6s} {r['desc'][:45]:45s} {val_fmt}")

    if nao_classificados:
        print(f"\n⚠ {len(nao_classificados)} não classificados:")
        for l in nao_classificados:
            print(f"  {l['data']} | {l['descricao'][:50]} | R$ {l['valor']:.2f}")

if __name__ == "__main__":
    main()
