import pandas as pd
import streamlit as st
import json
import base64
from datetime import datetime, date
from zoneinfo import ZoneInfo
import calendar
from playwright.sync_api import sync_playwright
import traceback
import os
import sys
import subprocess
from pathlib import Path
import re

MOSTRAR_SEPARACAO_ANALITICO = False  # caso queira a separação VARA CÍVEL e JUIZADO, troque pra True

# =========================
# CONFIGURAÇÃO DA PÁGINA
# =========================
st.set_page_config(page_title="Métricas Bradesco", layout="wide")

# =========================
# LINKS
# =========================
REL_ENCERRADO = "https://reports.elaw.com.br/Relatorio/Filtros?idRelatorio=45243"
REL_PENDENTE = "https://reports.elaw.com.br/Relatorio/Filtros?idRelatorio=45245"

LOGIN_URL = "https://ribeiroandrade.elawio.com.br/login"
LOGIN_EMAIL = "Mateusbradesco"
LOGIN_SENHA = "prazo"

METAS_FILE = "metas.json"
SITE_METRICAS_FILE = "metricas_site.json"

# =========================
# ESTILO
# =========================
st.markdown("""
<style>
    .cards-row {
        display: flex;
        justify-content: center;
        align-items: flex-start;
        gap: 22px;
        flex-wrap: wrap;
        margin: 0 auto 10px;
        max-width: 1400px;
    }

    .stApp {
        background-color: #d6d6d6;
    }

    .main-header {
        background-color: #ffcccc;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        text-align: center;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }

    .header-title {
        font-size: 75px;
        font-weight: bold;
        color: #990000;
        text-align: left;
        margin: 0;
    }

    .stColumn {
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
    }

    .card-completo {
        display: flex;
        flex-direction: column;
        align-items: center;
        width: 270px;
        margin: 0 auto;
    }

    .card-titulo {
        background-color: #ffcccc;
        color: #990000;
        font-size: 1.8rem;
        font-weight: bold;
        text-align: center;
        padding: 12px 20px;
        border-radius: 10px 10px 0 0;
        width: 100%;
        border: 2px solid #ffcccc;
        border-bottom: none;
        margin: 0;
    }

    .card-conteudo {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 0 0 15px 15px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        text-align: center;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        border: 2px solid #ffcccc;
        border-top: none;
        width: 100%;
        height: 160px;
    }

    .meta-externa {
        background-color: #ffcccc;
        padding: 8px 12px;
        border-radius: 8px;
        margin-top: 10px;
        font-size: 1.5rem;
        font-weight: bold;
        color: #990000;
        text-align: center;
        border: 1px solid #cc0000;
        width: 100%;
    }

    .metric-value {
        font-size: 2rem;
        font-weight: bold;
        color: #2c3e50;
        margin-bottom: 5px;
        line-height: 1.2;
    }

    .metric-label {
        font-size: 0.9rem;
        color: #7f8c8d;
        margin-bottom: 5px;
        line-height: 1.3;
    }

    .percentage {
        font-size: 1.5rem;
        font-weight: bold;
        color: #27ae60;
        margin: 5px 0;
    }

    .stButton > button {
        background-color: #ff6666 !important;
        color: white !important;
        border: 2px solid #cc0000 !important;
        border-radius: 10px !important;
        padding: 15px 30px !important;
        font-size: 1.2rem !important;
        font-weight: bold !important;
        transition: all 0.3s ease !important;
    }

    .stButton > button:hover {
        background-color: #ff3333 !important;
        transform: scale(1.05) !important;
    }

    .section-title {
        color: #990000;
        font-size: 1.8rem;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 20px;
        text-align: center;
    }

    .ticket-title {
        color: #990000;
        font-size: 2.5rem;
        font-weight: bold;
        margin-top: 20px;
        margin-bottom: 20px;
        text-align: center;
    }

    .analitico-title {
        color: #cc0000;
        font-size: 2rem;
        font-weight: bold;
        text-align: center;
        margin: 20px 0;
        padding: 10px;
        background-color: #ffcccc;
        border-radius: 10px;
    }

    .logo-img {
        max-height: 100px;
        object-fit: contain;
    }

    .date-label {
        font-size: 1.5rem;
        color: #990000;
        font-weight: bold;
        margin-bottom: 5px;
    }

    .table-title {
        color: #990000;
        font-size: 2.5rem;
        font-weight: bold;
        margin: 0;
    }

    .tooltip {
      position: relative;
      display: inline-block;
      cursor: help;
    }

    .tooltip .tooltiptext {
      visibility: hidden;
      width: 500px;
      background-color: #555;
      color: #fff;
      text-align: left;
      padding: 6px;
      border-radius: 6px;
      position: absolute;
      z-index: 1;
      bottom: 125%;
      left: 50%;
      margin-left: -110px;
      opacity: 0;
      transition: opacity 0.2s;
      font-size: 1.5rem;
      line-height: 1.3;
    }

    .tooltip:hover .tooltiptext {
      visibility: visible;
      opacity: 1;
    }
</style>
""", unsafe_allow_html=True)

# =========================
# REGRAS DO EXCEL)
# =========================

METRICAS_XLSX = "metricas_bradesco.xlsx"
PREVISAO_XLSX = "previsao_bradesco.xlsx"

CATEGORIAS = {
    "ACORDO": ["ACORDO"],
    "SEM_ONUS": [
        "DESISTENCIA DA ACAO",
        "EXTINTO SEM JULGAMENTO DO MERITO",
        "EXTINTO COM JULGAMENTO DO MERITO",
        "EXTINCAO SEM MERITO",
        "IMPROCEDENTE"
    ],
    "LIQUIDACAO": [
        "PROCEDENTE",
        "PROCEDENTE PARCIAL",
        "PROCEDENTE EM PARTE"
    ]
}

EXCLUIR_IDX26 = [
    "ACAO CAUTELAR",
    "RECLAMACAO PRE-PROCESSUAL",
    "CADASTRO IRREGULAR PELA DEPENDENCIA",
    "ACAO SUPERENDIVIDAMENTO",
    "ACAO DE PLANOS ECONOMICOS",
    "PROC. ADMINISTRATIVO",
    "ACAO BANCO REU",
]

# =========================
# METAS
# =========================
def carregar_metas():
    padrao = {
        "meta_encerramento": 664,
        "meta_sem_onus_perc": 55,
        "meta_acordo": 147,
        "meta_ticket_geral": 7500,
        "meta_ticket_acordo": 2500,
        "meta_ticket_liquidacao": 10842
    }

    try:
        with open(METAS_FILE, "r", encoding="utf-8") as f:
            metas = json.load(f)

        for k, v in padrao.items():
            metas.setdefault(k, v)

        return metas
    except:
        return padrao

# =========================
# JSON DO SITE
# =========================
def salvar_metricas_site_json(metricas: dict):
    with open(SITE_METRICAS_FILE, "w", encoding="utf-8") as f:
        json.dump(metricas, f, indent=4, ensure_ascii=False)

def carregar_metricas_site_json():
    if not os.path.exists(SITE_METRICAS_FILE):
        return None

    try:
        with open(SITE_METRICAS_FILE, "r", encoding="utf-8") as f:
            conteudo = f.read().strip()

        if not conteudo:
            return None

        return json.loads(conteudo)

    except (json.JSONDecodeError, OSError):
        return None
    
# =========================
# UTILITÁRIOS
# =========================
def ler_arquivo_binario(caminho):
    if not os.path.exists(caminho):
        return None
    with open(caminho, "rb") as f:
        return f.read()

def render_download_excels_usados():
    st.markdown("### 📥 Arquivos usados na atualização")

    col1, col2 = st.columns(2)

    with col1:
        dados_metricas = ler_arquivo_binario(METRICAS_XLSX)
        if dados_metricas is not None:
            st.download_button(
                label="Baixar metricas_bradesco.xlsx",
                data=dados_metricas,
                file_name="metricas_bradesco.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_metricas_excel_usado"
            )
        else:
            st.info("metricas_bradesco.xlsx ainda não foi gerado.")

    with col2:
        dados_previsao = ler_arquivo_binario(PREVISAO_XLSX)
        if dados_previsao is not None:
            st.download_button(
                label="Baixar previsao_bradesco.xlsx",
                data=dados_previsao,
                file_name="previsao_bradesco.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_previsao_excel_usado"
            )
        else:
            st.info("previsao_bradesco.xlsx ainda não foi gerado.")    

def normalizar_texto_serie(serie):
    return (
        serie.astype(str)
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .str.upper()
    )

def parse_br(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    x = str(x).strip()
    x = x.replace(".", "").replace(",", ".")
    try:
        return float(x)
    except:
        return None

def carregar_df_metricas_excel():
    df = pd.read_excel(METRICAS_XLSX)

    if df.shape[1] > 50:
        nome_col_data = df.columns[50]
        serie_data = pd.to_datetime(
            df[nome_col_data].astype("string"),
            errors="coerce",
            dayfirst=True
        ).dt.normalize()
        df[nome_col_data] = serie_data

    df = df.drop_duplicates(subset=df.columns[19], keep='first')

    return df

def carregar_previsao_excel():
    df_prev = pd.read_excel(PREVISAO_XLSX)

    df_prev = df_prev[df_prev.iloc[:, 7].notna()]
    df_prev = df_prev.drop_duplicates(subset=df_prev.columns[7], keep="first")

    col_b = (
        df_prev.iloc[:, 1]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    previsao_sem_onus = (
        col_b == "SOLICITAR ENCERRAMENTO ADMINISTRATIVO"
    ).sum()

    previsao_liquidacao = (
        col_b == "SOLICITAR CUMPRIMENTO DE OP AO CLIENTE"
    ).sum()

    return previsao_sem_onus, previsao_liquidacao

def filtrar_df_por_data(df, start_date=None, end_date=None):
    filtered_df = df.copy()

    if filtered_df.shape[1] > 50:
        if start_date is None or end_date is None:
            start_date, end_date = periodo_mes_corrente_dates()

        start_dt = pd.to_datetime(start_date, dayfirst=True).normalize()
        end_dt = pd.to_datetime(end_date, dayfirst=True).normalize()

        filtered_df = filtered_df[
            (filtered_df.iloc[:, 50] >= start_dt) &
            (filtered_df.iloc[:, 50] <= end_dt)
        ]

    return filtered_df

def calcular_metricas_dos_excels(start_date=None, end_date=None):
    df = carregar_df_metricas_excel()

    previsao_sem_onus, previsao_liquidacao = carregar_previsao_excel()

    filtered_df = filtrar_df_por_data(df, start_date, end_date)

    texto_idx26 = normalizar_texto_serie(filtered_df.iloc[:, 26]).fillna("")
    excluir_idx26_norm = normalizar_texto_serie(pd.Series(EXCLUIR_IDX26)).tolist()
    regex_excluir_idx26 = "|".join(re.escape(x) for x in excluir_idx26_norm)

    filtered_df = filtered_df[
        ~texto_idx26.str.contains(regex_excluir_idx26, na=False, regex=True)
    ].copy()

    tipo_encerramento = normalizar_texto_serie(filtered_df.iloc[:, 49])

    categorias_validas = (
        CATEGORIAS["ACORDO"] +
        CATEGORIAS["SEM_ONUS"] +
        CATEGORIAS["LIQUIDACAO"]
    )

    filtered_df = filtered_df[tipo_encerramento.isin(categorias_validas)].copy()
    tipo_encerramento = normalizar_texto_serie(filtered_df.iloc[:, 49])

    acordo_mask = tipo_encerramento.isin(CATEGORIAS["ACORDO"])
    sem_onus_mask = tipo_encerramento.isin(CATEGORIAS["SEM_ONUS"])
    liquidacao_mask = tipo_encerramento.isin(CATEGORIAS["LIQUIDACAO"])

    acordo = int(acordo_mask.sum())
    sem_onus = int(sem_onus_mask.sum())
    liquidacao = int(liquidacao_mask.sum())

    total_encerrados = acordo + sem_onus + liquidacao

    valores_todos = filtered_df.iloc[:, 52].apply(parse_br).dropna()
    ticket_medio_geral = float(valores_todos.mean()) if len(valores_todos) else 0

    valores_acordo = filtered_df.loc[acordo_mask, filtered_df.columns[52]].apply(parse_br).dropna()
    ticket_medio_acordo = float(valores_acordo.mean()) if len(valores_acordo) else 0

    valores_liquidacao = filtered_df.loc[
        liquidacao_mask, filtered_df.columns[52]
    ].apply(parse_br)

    valores_liquidacao = valores_liquidacao[
        (valores_liquidacao >= 200) & (valores_liquidacao <= 55000)
    ].dropna()

    ticket_medio_liquidacao = float(valores_liquidacao.mean()) if len(valores_liquidacao) else 0

    return {
        "acordo": acordo,
        "sem_onus": sem_onus,
        "liquidacao": liquidacao,
        "total_encerrados": total_encerrados,
        "previsao_sem_onus": int(previsao_sem_onus),
        "previsao_liquidacao": int(previsao_liquidacao),
        "ticket_medio_geral": ticket_medio_geral,
        "ticket_medio_acordo": ticket_medio_acordo,
        "ticket_medio_liquidacao": ticket_medio_liquidacao,
    }

def _click_gerar_excel_e_salvar(scope, page, destino_arquivo):
    candidatos = [
        "#excel",
        "xpath=//button[@id='excel']",
        "xpath=//button[contains(normalize-space(.), 'Gerar Excel')]",
        "text=Gerar Excel",
    ]

    ultimo_erro = None

    for sel in candidatos:
        try:
            botao = scope.locator(sel).first
            botao.wait_for(state="visible", timeout=15000)
            botao.scroll_into_view_if_needed()

            with page.expect_download(timeout=120000) as download_info:
                botao.click(force=True)

            download = download_info.value
            download.save_as(destino_arquivo)
            page.wait_for_timeout(1500)
            return destino_arquivo

        except Exception as e:
            ultimo_erro = e

    raise Exception(f"Não foi possível baixar o Excel. Último erro: {ultimo_erro}")

def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

def periodo_mes_corrente_str():
    hoje = date.today()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia = hoje.replace(day=calendar.monthrange(hoje.year, hoje.month)[1])
    return primeiro_dia.strftime("%d/%m/%Y"), ultimo_dia.strftime("%d/%m/%Y")

def periodo_mes_corrente_dates():
    hoje = date.today()
    primeiro_dia = hoje.replace(day=1)
    ultimo_dia = hoje.replace(day=calendar.monthrange(hoje.year, hoje.month)[1])
    return primeiro_dia, ultimo_dia

# =========================
# PLAYWRIGHT
# =========================
def ensure_playwright_chromium():
    cache_dir = Path.home() / ".cache" / "ms-playwright"

    chromium_ok = any(cache_dir.glob("chromium-*/chrome-linux/chrome"))
    shell_ok = any(cache_dir.glob("chromium_headless_shell-*/chrome-headless-shell-linux64/chrome-headless-shell"))

    if chromium_ok or shell_ok:
        return

    subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

def _fw():
    ensure_playwright_chromium()
    return sync_playwright().start()

def _browser_context(pw):
    browser = pw.chromium.launch(headless=True)
    context = browser.new_context(
        viewport={"width": 1920, "height": 1080},
        accept_downloads=True
    )
    context.set_default_timeout(60000)
    context.set_default_navigation_timeout(90000)

    page = context.new_page()
    page.set_default_timeout(60000)
    page.set_default_navigation_timeout(90000)

    return browser, context, page

# =========================
# LOGIN / NAVEGAÇÃO
# =========================
def _login(page):
    page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=90000)

    page.locator("input[type='text'], input[placeholder*='Email'], input[name='Email']").first.fill(LOGIN_EMAIL)
    page.locator("input[type='password'], input[placeholder*='Senha'], input[name='Senha']").first.fill(LOGIN_SENHA)

    entrar = page.locator("button:has-text('Entrar'), a:has-text('Entrar'), input[value='Entrar']").first
    entrar.click()

    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)

    url_atual = page.url.lower()
    if "/home/index" not in url_atual:
        raise Exception(f"Login não concluiu corretamente. URL após login: {page.url}")

def _entrar_modulo_relatorios(page):
    context = page.context
    paginas_antes = list(context.pages)

    candidatos = [
        "xpath=/html/body/div[7]/div[1]/div/ul/li[5]/a",
        "xpath=//a[.//span[contains(normalize-space(.), 'Módulo de Relatórios')]]",
        "xpath=//li[.//span[contains(normalize-space(.), 'Módulo de Relatórios')]]/a",
    ]

    ultimo_erro = None
    link = None

    for sel in candidatos:
        try:
            loc = page.locator(sel).first
            loc.wait_for(state="visible", timeout=20000)
            loc.scroll_into_view_if_needed()
            link = loc
            break
        except Exception as e:
            ultimo_erro = e

    if link is None:
        raise Exception(f"Não encontrei o link 'Módulo de Relatórios'. Último erro: {ultimo_erro}")

    reports_page = None

    try:
        with page.expect_popup(timeout=20000) as popup_info:
            link.click(force=True)
        reports_page = popup_info.value
    except Exception:
        link.click(force=True)
        page.wait_for_timeout(5000)

        novas_paginas = [p for p in context.pages if p not in paginas_antes]
        if novas_paginas:
            reports_page = novas_paginas[-1]
        else:
            reports_page = page

    reports_page.wait_for_load_state("domcontentloaded")
    reports_page.wait_for_load_state("networkidle")
    reports_page.wait_for_timeout(3000)

    url_final = reports_page.url.lower()

    if "reports.elaw.com.br/home/index" not in url_final and "reports.elaw.com.br" not in url_final:
        raise Exception(
            "O clique em 'Módulo de Relatórios' não concluiu a abertura do Reports. "
            f"URL final: {reports_page.url}"
        )

    return reports_page

def _get_scope_relatorio(page, timeout_ms=45000):
    seletores_prontos = [
        "#form0",
        "xpath=//form[@id='form0']",
        "#excel",
        "xpath=//button[@id='excel']",
        "xpath=//button[contains(normalize-space(.), 'Gerar Excel')]",
    ]

    tentativas = max(1, timeout_ms // 1000)

    for _ in range(tentativas):
        scopes = [page] + list(page.frames)

        for scope in scopes:
            for sel in seletores_prontos:
                try:
                    loc = scope.locator(sel)
                    if loc.count() > 0:
                        loc.first.wait_for(state="attached", timeout=1000)
                        return scope
                except Exception:
                    pass

        page.wait_for_timeout(1000)

    detalhes = [f"URL atual: {page.url}"]

    try:
        detalhes.append(f"Título: {page.title()}")
    except Exception:
        pass

    for i, fr in enumerate(page.frames):
        try:
            detalhes.append(f"Frame {i}: {fr.url}")
        except Exception:
            pass

    raise Exception(
        "Não encontrei a área do relatório. "
        "Procurei por #form0, #excel e botão 'Gerar Excel'. "
        + " | ".join(detalhes)
    )

def _abrir_relatorio(page, url):
    page.goto(url, wait_until="domcontentloaded", timeout=90000)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(3000)

    # Se caiu de novo na tela de login, autentica e reabre o relatório
    try:
        if page.locator("input[type='password'], input[placeholder*='Senha'], input[name='Senha']").count() > 0:
            _login(page)
            page.goto(url, wait_until="domcontentloaded", timeout=90000)
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(3000)
    except Exception:
        pass

# =========================
# AÇÕES DE FORMULÁRIO
# =========================

def _abrir_dropdown(scope, button_xpath):
    botao = scope.locator(f"xpath={button_xpath}").first
    botao.wait_for(state="visible", timeout=15000)
    botao.scroll_into_view_if_needed()
    botao.click(force=True)
    scope.wait_for_timeout(800)

    # sobe para o container do multiselect que ficou aberto
    grupo = botao.locator("xpath=ancestor::div[contains(@class,'btn-group')][1]")
    grupo.wait_for(state="attached", timeout=10000)

    classe = (grupo.get_attribute("class") or "").lower()
    if "open" not in classe:
        # tenta mais uma vez
        botao.click(force=True)
        scope.wait_for_timeout(800)
        classe = (grupo.get_attribute("class") or "").lower()

    if "open" not in classe:
        raise Exception("O dropdown não abriu. O btn-group não recebeu a classe 'open'.")

    dropdown = grupo.locator("xpath=.//ul[contains(@class,'multiselect-container')]").first
    dropdown.wait_for(state="attached", timeout=10000)
    return dropdown

def _fechar_dropdown(scope):
    scope.locator("body").click(position={"x": 5, "y": 5})
    scope.wait_for_timeout(500)

def _marcar_checkbox_multiselect_por_xpath(scope, button_xpath, checkbox_xpaths):
    dropdown = _abrir_dropdown(scope, button_xpath)

    for cb_xpath in checkbox_xpaths:
        item = scope.locator(f"xpath={cb_xpath}").first
        item.wait_for(state="attached", timeout=10000)
        item.scroll_into_view_if_needed()

        li = item.locator("xpath=ancestor::li[1]")
        classe = (li.get_attribute("class") or "").lower()

        if "active" not in classe:
            item.click(force=True)
            scope.wait_for_timeout(300)

    _fechar_dropdown(scope)

def _set_input_value_by_xpath(scope, xpath, value_text):
    inp = scope.locator(f"xpath={xpath}").first
    inp.wait_for(state="visible", timeout=15000)
    inp.scroll_into_view_if_needed()
    inp.click(force=True)
    inp.fill("")

    el = inp.element_handle()
    scope.evaluate(
        """
        ([element, value]) => {
            element.value = '';
            element.dispatchEvent(new Event('input', { bubbles: true }));
            element.dispatchEvent(new Event('change', { bubbles: true }));

            element.value = value;
            element.dispatchEvent(new Event('input', { bubbles: true }));
            element.dispatchEvent(new Event('change', { bubbles: true }));
            element.blur();
        }
        """,
        [el, value_text]
    )

    scope.wait_for_timeout(500)

# =========================
# BAIXAR EXCEL ENCERRADOS
# =========================
def _baixar_relatorio_encerrado_excel(page, destino_arquivo=METRICAS_XLSX):
    data_inicial, data_final = periodo_mes_corrente_str()
    periodo = f"{data_inicial} - {data_final}"

    _abrir_relatorio(page, REL_ENCERRADO)
    scope = _get_scope_relatorio(page)

    X_STATUS_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[1]/td[2]/div/button'
    X_ESCRITORIO_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[4]/td[2]/div/button'

    X_STATUS_ENCERRADO = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[1]/td[2]/div/ul/li[3]/a/label'
    X_ESCRITORIO_BRADESCO = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[4]/td[2]/div/ul/li[3]/a/label'
    X_ESCRITORIO_BRADESCO_ATIVOS = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[4]/td[2]/div/ul/li[4]/a/label'

    X_TIPO_PROCESSO_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[5]/td[2]/div/button'
    X_TIPO_PROCESSO_PRINCIPAL = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[5]/td[2]/div/ul/li[3]/a/label'

    X_DATA_BAIXA = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[8]/td[2]/input'

    _marcar_checkbox_multiselect_por_xpath(
        scope,
        X_STATUS_BTN,
        [X_STATUS_ENCERRADO]
    )

    _marcar_checkbox_multiselect_por_xpath(
    scope,
    X_TIPO_PROCESSO_BTN,
    [X_TIPO_PROCESSO_PRINCIPAL]
    )

    _marcar_checkbox_multiselect_por_xpath(
        scope,
        X_ESCRITORIO_BTN,
        [X_ESCRITORIO_BRADESCO, X_ESCRITORIO_BRADESCO_ATIVOS]
    )

    _set_input_value_by_xpath(scope, X_DATA_BAIXA, periodo)

    _click_gerar_excel_e_salvar(scope, page, destino_arquivo)

# =========================
# BAIXAR EXCEL PENDENTES
# =========================

def _baixar_relatorio_pendente_excel(page, destino_arquivo=PREVISAO_XLSX):
    data_inicial, data_final = periodo_mes_corrente_str()
    periodo = f"{data_inicial} - {data_final}"

    _abrir_relatorio(page, REL_PENDENTE)
    scope = _get_scope_relatorio(page)

    X_STATUS_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[1]/td[2]/div/button'
    X_TIPO_COMPROMISSO_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[2]/td[2]/div/button'
    X_SUBTIPO_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[3]/td[2]/div/button'
    X_CELULA_BTN = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[6]/td[2]/div/button'

    X_PERIODO_PRAZO = '//*[@id="form0"]/div/div[2]/table/tbody/tr[4]/td[2]/input'

    X_STATUS_PENDENTE = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[1]/td[2]/div/ul/li[3]/a/label'

    X_TIPO_COMPROMISSO = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[2]/td[2]/div/ul/li[4]/a/label'
    X_TIPO_PRAZO = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[2]/td[2]/div/ul/li[6]/a/label'

    X_SUBTIPO_OP_CLIENTE = '//input[@type="checkbox" and @value="58|Solicitar Cumprimento de OP ao Cliente"]/parent::label'
    X_SUBTIPO_ENC_ADMIN = '//input[@type="checkbox" and @value="114|Solicitar encerramento administrativo"]/parent::label'

    X_CELULA_BRADESCO = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[6]/td[2]/div/ul/li[3]/a/label'
    X_CELULA_BRADESCO_ATIVOS = '/html/body/div[2]/div/div/form/div/div[2]/table/tbody/tr[6]/td[2]/div/ul/li[4]/a/label'

    _marcar_checkbox_multiselect_por_xpath(scope, X_STATUS_BTN, [X_STATUS_PENDENTE])

    _marcar_checkbox_multiselect_por_xpath(
        scope,
        X_TIPO_COMPROMISSO_BTN,
        [X_TIPO_COMPROMISSO, X_TIPO_PRAZO]
    )

    # aqui agora marca os dois subtipos de uma vez
    _marcar_checkbox_multiselect_por_xpath(
        scope,
        X_SUBTIPO_BTN,
        [X_SUBTIPO_OP_CLIENTE, X_SUBTIPO_ENC_ADMIN]
    )

    _set_input_value_by_xpath(scope, X_PERIODO_PRAZO, periodo)

    _marcar_checkbox_multiselect_por_xpath(
        scope,
        X_CELULA_BTN,
        [X_CELULA_BRADESCO, X_CELULA_BRADESCO_ATIVOS]
    )

    _click_gerar_excel_e_salvar(scope, page, destino_arquivo)

# =========================
# COLETA DOS ENCERRADOS
# =========================
def coletar_metricas_encerrados_site():
    pw = _fw()
    browser = context = page = None

    try:
        browser, context, page = _browser_context(pw)
        _login(page)

        reports_page = _entrar_modulo_relatorios(page)
        _baixar_relatorio_encerrado_excel(reports_page, METRICAS_XLSX)

        return {"arquivo_encerrados_ok": True}

    finally:
        if context:
            context.close()
        if browser:
            browser.close()
        pw.stop()

# =========================
# COLETA DOS PENDENTES
# =========================
def coletar_metricas_pendentes_site():
    pw = _fw()
    browser = context = page = None

    try:
        browser, context, page = _browser_context(pw)
        _login(page)

        reports_page = _entrar_modulo_relatorios(page)
        _baixar_relatorio_pendente_excel(reports_page, PREVISAO_XLSX)

        return {"arquivo_pendentes_ok": True}

    finally:
        if context:
            context.close()
        if browser:
            browser.close()
        pw.stop()

# =========================
# ATUALIZAR METRICAS e RESTAURAR BACKUP (pra não reiniciar vazio)
# =========================
def atualizar_metricas_site_json():
    coletar_metricas_encerrados_site()
    coletar_metricas_pendentes_site()

    inicio, fim = periodo_mes_corrente_dates()

    metricas = calcular_metricas_dos_excels(
        start_date=inicio,
        end_date=fim
    )

    metricas["atualizado_em"] = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M:%S")
    metricas["start_date"] = inicio.strftime("%Y-%m-%d")
    metricas["end_date"] = fim.strftime("%Y-%m-%d")

    salvar_metricas_site_json(metricas)
    return metricas


# =========================
# CÁLCULO FINAL DOS CARDS
# =========================
def calculate_metrics_from_json(base):
    metas = st.session_state.metas

    acordo = base.get("acordo", 0)
    sem_onus = base.get("sem_onus", 0)
    liquidacao = base.get("liquidacao", 0)

    previsao_sem_onus = base.get("previsao_sem_onus", 0)
    previsao_liquidacao = base.get("previsao_liquidacao", 0)

    total = acordo + sem_onus + liquidacao
    total_proj = total + previsao_sem_onus + previsao_liquidacao

    metrics = {
        "meta_encerramento": metas["meta_encerramento"],
        "meta_acordo": metas["meta_acordo"],
        "meta_sem_onus_perc": metas["meta_sem_onus_perc"],
        "meta_ticket_geral": metas["meta_ticket_geral"],
        "meta_ticket_acordo": metas["meta_ticket_acordo"],
        "meta_ticket_liquidacao": metas["meta_ticket_liquidacao"],

        "acordo": acordo,
        "sem_onus": sem_onus,
        "liquidacao": liquidacao,
        "total_encerrados": total,

        "previsao_sem_onus": previsao_sem_onus,
        "previsao_liquidacao": previsao_liquidacao,

        "ticket_medio_geral": base.get("ticket_medio_geral", 0),
        "ticket_medio_acordo": base.get("ticket_medio_acordo", 0),
        "ticket_medio_liquidacao": base.get("ticket_medio_liquidacao", 0),
    }

    metrics["perc_liquidacao"] = (liquidacao / total * 100) if total else 0
    metrics["perc_sem_onus"] = (sem_onus / total * 100) if total else 0
    metrics["perc_acordo"] = (acordo / total * 100) if total else 0

    metrics["performance_encerramento"] = (total / metas["meta_encerramento"] * 100) if metas["meta_encerramento"] else 0
    metrics["performance_encerramento_proj"] = (total_proj / metas["meta_encerramento"] * 100) if metas["meta_encerramento"] else 0

    metrics["performance_acordo"] = (acordo / metas["meta_acordo"] * 100) if metas["meta_acordo"] else 0

    meta_sem_onus_real = int(round(metas["meta_encerramento"] * metas["meta_sem_onus_perc"] / 100))
    metrics["performance_sem_onus"] = (sem_onus / meta_sem_onus_real * 100) if meta_sem_onus_real else 0
    metrics["performance_sem_onus_proj"] = ((sem_onus + previsao_sem_onus) / meta_sem_onus_real * 100) if meta_sem_onus_real else 0

    metrics["perc_liquidacao_proj"] = ((liquidacao + previsao_liquidacao) / total_proj * 100) if total_proj else 0

    metrics["performance_ticket_geral"] = (
        metrics["ticket_medio_geral"] / metas["meta_ticket_geral"] * 100
    ) if metas["meta_ticket_geral"] else 0

    metrics["performance_ticket_acordo"] = (
        metrics["ticket_medio_acordo"] / metas["meta_ticket_acordo"] * 100
    ) if metas["meta_ticket_acordo"] else 0

    metrics["performance_ticket_liquidacao"] = (
        metrics["ticket_medio_liquidacao"] / metas["meta_ticket_liquidacao"] * 100
    ) if metas["meta_ticket_liquidacao"] else 0

    return metrics

# =========================
# HEADER
# =========================
def render_header():
    try:
        img_base64 = get_base64_image("logobradesco.png")
        img_html = f'<img src="data:image/png;base64,{img_base64}" class="logo-img">'
    except:
        img_html = '''
        <div style="
            height: 100px;
            background-color: #ffcccc;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 10px;
        ">
            <span style="color: #990000;">LOGO BRADESCO</span>
        </div>
        '''

    inicio, fim = periodo_mes_corrente_dates()

    col1, col2, col3, col4 = st.columns([1.0, 2.7, 2.5, 1.8], vertical_alignment="center")

    with col1:
        st.markdown(
            f'''
            <div style="
                height: 100px;
                display: flex;
                align-items: center;
                justify-content: flex-start;
            ">
                {img_html}
            </div>
            ''',
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            '''
            <div style="
                height: 100px;
                display: flex;
                align-items: center;
                justify-content: flex-start;
            ">
                <div style="
                    font-size: 3.2rem;
                    font-weight: 800;
                    color: #990000;
                    line-height: 1;
                    white-space: nowrap;
                ">
                    MÉTRICAS BRADESCO
                </div>
            </div>
            ''',
            unsafe_allow_html=True
        )

    with col3:
        st.markdown(
            f'''
            <div style="
                height: 100px;
                display: flex;
                align-items: center;
                justify-content: center;
            ">
                <div style="
                    font-size: 1.5rem;
                    color: #990000;
                    font-weight: bold;
                    background-color: #ffcccc;
                    padding: 12px 22px;
                    border-radius: 10px;
                    white-space: nowrap;
                ">
                    📅 Período atual: {inicio.strftime("%d/%m/%Y")} - {fim.strftime("%d/%m/%Y")}
                </div>
            </div>
            ''',
            unsafe_allow_html=True
        )

    with col4:
        st.write("")
        atualizar = st.button("ATUALIZAR DADOS", use_container_width=True)

    return atualizar

# =========================
# RENDER METRICS
# =========================
def render_metrics(metrics):
    pendentes_total = (
        metrics["previsao_sem_onus"] +
        metrics["previsao_liquidacao"]
    )

    st.markdown(
        f'''
        <div class="ticket-title">
            🎯 ENCERRAMENTOS 🎯
            <span class="tooltip"> &#x1F6C8;
                <span class="tooltiptext">
                    Em () o total com os pendentes:
                    <br>• Sem Ônus: {metrics["previsao_sem_onus"]}
                    <br>• Liquidação: {metrics["previsao_liquidacao"]}
                    <br>
                    Total pendentes: {pendentes_total}
                </span>
            </span>
        </div>
        ''',
        unsafe_allow_html=True
    )

    st.markdown("")
    _, col1, col2, col3, col4, _ = st.columns([1, 1, 1, 1, 1, 1])

    with col1:
        icm_encerramento_color = "#27ae60" if metrics["performance_encerramento"] >= 100 else "#fdc303"
        st.markdown(f'''
            <div class="card-completo" style="margin-left: auto; margin-right: 10px;">
                <div class="card-titulo">🗃️ TOTAL</div>
                <div class="card-conteudo">
                    <div class="metric-value">{metrics["total_encerrados"]} ({metrics["total_encerrados"] + metrics["previsao_sem_onus"] + metrics["previsao_liquidacao"]})</div>
                    <div class="percentage" style="color: {icm_encerramento_color}; font-weight: bold;">
                        {metrics["performance_encerramento"]:.1f}% ({metrics["performance_encerramento_proj"]:.1f}%)
                    </div>
                </div>
                <div class="meta-externa">Meta: {metrics["meta_encerramento"]} casos</div>
            </div>
        ''', unsafe_allow_html=True)

    with col2:
        st.markdown(f'''
        <div class="card-completo">
            <div class="card-titulo">👨‍⚖️ LIQUIDAÇÃO</div>
            <div class="card-conteudo">
                <div class="metric-value">{metrics["liquidacao"]} ({metrics["liquidacao"] + metrics["previsao_liquidacao"]})</div>
                <div class="percentage" style="color: #7f8c8d; font-weight: bold;">{metrics["perc_liquidacao"]:.1f}%</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

    with col3:
        icm_sem_onus_color = "#27ae60" if metrics["performance_sem_onus"] >= 100 else "#fdc303"
        st.markdown(f'''
        <div class="card-completo">
            <div class="card-titulo">🍀 SEM ÔNUS</div>
            <div class="card-conteudo">
                <div class="metric-value">{metrics["sem_onus"]} ({metrics["sem_onus"] + metrics["previsao_sem_onus"]})</div>
                <div class="percentage" style="color: {icm_sem_onus_color}; font-weight: bold;">
                    {metrics["performance_sem_onus"]:.1f}% ({metrics["performance_sem_onus_proj"]:.1f}%)
                </div>
            </div>
            <div class="meta-externa">Meta: {metrics["meta_sem_onus_perc"]}% ({int(round(metrics["meta_encerramento"] * metrics["meta_sem_onus_perc"]/100))})</div>
        </div>
        ''', unsafe_allow_html=True)

    with col4:
        icm_acordo_color = "#27ae60" if metrics["performance_acordo"] >= 100 else "#fdc303"
        st.markdown(f'''
        <div class="card-completo">
            <div class="card-titulo">🤝🏻 ACORDO</div>
            <div class="card-conteudo">
                <div class="metric-value">{metrics["acordo"]}</div>
                <div class="percentage" style="color: {icm_acordo_color}; font-weight: bold;">{metrics["performance_acordo"]:.1f}%</div>
            </div>
            <div class="meta-externa">Meta: {metrics["meta_acordo"]} casos</div>
        </div>
        ''', unsafe_allow_html=True)

    st.markdown("___")
    st.markdown('<div class="ticket-title">💸 TICKET MÉDIO 💸</div>', unsafe_allow_html=True)
    st.markdown("")

    _, col_ticket1, col_ticket2, col_ticket3, _ = st.columns([2, 2, 2, 2, 2])

    with col_ticket1:
        icm_ticket_geral_color = "#cc0000" if metrics["performance_ticket_geral"] > 100 else "#27ae60"
        st.markdown(f'''
        <div class="card-completo">
            <div class="card-titulo">GERAL</div>
            <div class="card-conteudo">
                <div class="metric-value">R$ {metrics["ticket_medio_geral"]:,.2f}</div>
                <div class="percentage" style="color: {icm_ticket_geral_color};">{metrics["performance_ticket_geral"]:.1f}%</div>
            </div>
            <div class="meta-externa">Meta: R$ {metrics["meta_ticket_geral"]:,.2f}</div>
        </div>
        ''', unsafe_allow_html=True)

    with col_ticket2:
        icm_ticket_acordo_color = "#cc0000" if metrics["performance_ticket_acordo"] > 100 else "#27ae60"
        st.markdown(f'''
        <div class="card-completo">
            <div class="card-titulo">ACORDO</div>
            <div class="card-conteudo">
                <div class="metric-value">R$ {metrics["ticket_medio_acordo"]:,.2f}</div>
                <div class="percentage" style="color: {icm_ticket_acordo_color};">{metrics["performance_ticket_acordo"]:.1f}%</div>
            </div>
            <div class="meta-externa">Meta: R$ {metrics["meta_ticket_acordo"]:,.2f}</div>
        </div>
        ''', unsafe_allow_html=True)

    with col_ticket3:
        icm_ticket_liq_color = "#cc0000" if metrics["performance_ticket_liquidacao"] > 100 else "#27ae60"
        st.markdown(f'''
        <div class="card-completo">
            <div class="card-titulo">LIQUIDAÇÃO</div>
            <div class="card-conteudo">
                <div class="metric-value">R$ {metrics["ticket_medio_liquidacao"]:,.2f}</div>
                <div class="percentage" style="color: {icm_ticket_liq_color};">{metrics["performance_ticket_liquidacao"]:.1f}%</div>
            </div>
            <div class="meta-externa">Meta: R$ {metrics["meta_ticket_liquidacao"]:,.2f}</div>
        </div>
        ''', unsafe_allow_html=True)

    st.markdown("___")

@st.dialog("Instruções para arquivos fonte")
def dialog_instrucoes_excel():
    st.markdown("""
    Agora a fonte principal é o site, mas o cálculo volta a usar os Excel exportados.

    Ao clicar em **ATUALIZAR DADOS**, o robô:
    - faz login
    - entra no **Módulo de Relatórios**
    - baixa o relatório de encerrados como `metricas_bradesco.xlsx`
    - baixa o relatório de pendentes como `previsao_bradesco.xlsx`
    - aplica as regras já usadas no modo manual
    - salva os agregados em `metricas_site.json`
    """)

# =========================
# MAIN
# =========================
def main():
    if "metas" not in st.session_state:
        st.session_state.metas = carregar_metas()

    if "metricas_site" not in st.session_state:
        st.session_state.metricas_site = carregar_metricas_site_json()

    clicou_atualizar = render_header()

    if clicou_atualizar:
        with st.spinner("Atualizando dados direto do relatório..."):
            try:
                metricas_site = atualizar_metricas_site_json()
                st.session_state.metricas_site = metricas_site
                st.success("Dados atualizados com sucesso.")
                st.rerun()

            except Exception as e:
                st.error(f"Erro ao atualizar os dados: {type(e).__name__}: {repr(e)}")
                st.code(traceback.format_exc())

    st.markdown("___")

    try:
        if not st.session_state.metricas_site:
            st.info("Clique em ATUALIZAR DADOS para carregar os relatórios.")
            return

        metrics = calculate_metrics_from_json(st.session_state.metricas_site)
        render_metrics(metrics)

        st.markdown(
            f"""
            <div style="background-color:#f8f9fa; padding:15px; border-radius:10px; border:1px solid #ddd;">
                <b>Última atualização:</b> {st.session_state.metricas_site.get("atualizado_em", "-")}
            </div>
            """,
            unsafe_allow_html=True
        )

    except Exception as e:
        st.error(f"Erro ao carregar os dados: {str(e)}")

    st.markdown("___")

    col_info, _ = st.columns([1, 5])

    with col_info:
        if st.button("💭Instruções arquivos fonte"):
            dialog_instrucoes_excel()


    st.markdown("___")
    render_download_excels_usados()

if __name__ == "__main__":
    main()