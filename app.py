# app_cnpj.py
import io
import re
import time
import requests
import pandas as pd
import streamlit as st

# ==============================
# VERSÃO
# ==============================
VERSAO = "V1.0"

# ==============================
# TEMA TR (espelho do eventos)
# ==============================
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif;
            color: #444444;
        }
        h1, h2, h3 {
            color: #FF8000;
            font-weight: 700;
        }
        section[data-testid="stSidebar"] {
            background-color: #444444;
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        .stButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        .stDownloadButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stDownloadButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 16px 20px;
            margin: 12px 0;
            color: #444444;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        .instrucoes-box h4 {
            color: #FF8000;
            margin-top: 14px;
            margin-bottom: 6px;
        }
        .instrucoes-box h4:first-child { margin-top: 0; }
        .card-empresa {
            background: #FFFFFF;
            border: 1px solid #E0E0E0;
            border-left: 5px solid #FF8000;
            border-radius: 6px;
            padding: 14px 18px;
            margin-bottom: 10px;
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 14px;
            color: #333;
        }
        .card-empresa .razao {
            font-size: 16px;
            font-weight: 700;
            color: #FF8000;
            margin-bottom: 6px;
        }
        .card-empresa .tag-tipo {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: bold;
            margin-bottom: 8px;
        }
        .tag-cliente {
            background: #E3F2FD;
            color: #1565C0;
        }
        .tag-fornecedor {
            background: #FFF3E0;
            color: #E65100;
        }
        .tag-erro {
            background: #FFEBEE;
            color: #B71C1C;
        }
        .status-ok  { color: #2E7D32; font-weight: bold; }
        .status-err { color: #B71C1C; font-weight: bold; }
        </style>
    """, unsafe_allow_html=True)


# ==============================
# UTILITÁRIOS
# ==============================
def so_numeros(v: str) -> str:
    return re.sub(r"\D", "", str(v or ""))


def formatar_cnpj(cnpj: str) -> str:
    n = so_numeros(cnpj).zfill(14)
    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:14]}"


def formatar_cep(cep: str) -> str:
    n = so_numeros(cep).zfill(8)
    return f"{n[:5]}-{n[5:]}"


def formatar_telefone(ddd: str, numero: str) -> str:
    ddd = so_numeros(ddd)
    numero = so_numeros(numero)
    if not ddd and not numero:
        return ""
    return f"({ddd}) {numero}" if ddd else numero


def validar_cnpj(cnpj: str) -> bool:
    """Valida dígitos verificadores do CNPJ."""
    n = so_numeros(cnpj)
    if len(n) != 14 or len(set(n)) == 1:
        return False
    def calc(digits, pesos):
        s = sum(int(d) * p for d, p in zip(digits, pesos))
        r = s % 11
        return 0 if r < 2 else 11 - r
    p1 = [5,4,3,2,9,8,7,6,5,4,3,2]
    p2 = [6,5,4,3,2,9,8,7,6,5,4,3,2]
    return (
        calc(n[:12], p1) == int(n[12]) and
        calc(n[:13], p2) == int(n[13])
    )


# ==============================
# CONSULTA BRASILAPI
# ==============================
BRASILAPI_URL = "https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
RECEITAWS_URL = "https://receitaws.com.br/v1/cnpj/{cnpj}"

def consultar_cnpj(cnpj_raw: str) -> dict:
    """
    Consulta o CNPJ na BrasilAPI (fallback: ReceitaWS).
    Retorna dict com campos normalizados ou chave 'erro'.
    """
    cnpj = so_numeros(cnpj_raw)
    if len(cnpj) != 14:
        return {"erro": "CNPJ deve ter 14 dígitos."}
    if not validar_cnpj(cnpj):
        return {"erro": "CNPJ inválido (dígitos verificadores incorretos)."}

    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}

    # --- tentativa 1: BrasilAPI ---
    try:
        resp = requests.get(
            BRASILAPI_URL.format(cnpj=cnpj),
            headers=headers,
            timeout=15,
        )
        if resp.status_code == 200:
            return _normalizar_brasilapi(resp.json())
        elif resp.status_code == 429:
            time.sleep(2)
            resp = requests.get(
                BRASILAPI_URL.format(cnpj=cnpj),
                headers=headers,
                timeout=15,
            )
            if resp.status_code == 200:
                return _normalizar_brasilapi(resp.json())
    except Exception:
        pass

    # --- fallback: ReceitaWS ---
    try:
        resp = requests.get(
            RECEITAWS_URL.format(cnpj=cnpj),
            headers=headers,
            timeout=15,
        )
        if resp.status_code == 200:
            return _normalizar_receitaws(resp.json())
    except Exception:
        pass

    return {"erro": "Não foi possível consultar o CNPJ. Tente novamente mais tarde."}


def _normalizar_brasilapi(data: dict) -> dict:
    """Mapeia resposta BrasilAPI para campos internos."""
    if data.get("type") == "service_error" or "message" in data:
        return {"erro": data.get("message", "CNPJ não encontrado.")}

    municipio = data.get("municipio", "") or data.get("municipio_nome", "")
    uf        = data.get("uf", "")
    bairro    = data.get("bairro", "")
    logradouro= data.get("logradouro", "")
    numero    = data.get("numero", "")
    complemento = data.get("complemento", "")
    cep       = so_numeros(str(data.get("cep", "")))
    ddd_tel   = str(data.get("ddd_telefone_1", ""))
    # BrasilAPI retorna ddd_telefone_1 como "11 99999999" ou "(11) 9999-9999"
    tel_fmt   = _extrair_telefone(ddd_tel)
    ie_lista  = data.get("qsa", [])  # não tem IE diretamente

    return {
        "cnpj"           : so_numeros(str(data.get("cnpj", ""))),
        "razao_social"   : (data.get("razao_social") or "").strip(),
        "nome_fantasia"  : (data.get("nome_fantasia") or "").strip(),
        "situacao"       : (data.get("descricao_situacao_cadastral") or "").strip(),
        "uf"             : uf,
        "municipio"      : municipio,
        "bairro"         : bairro,
        "logradouro"     : logradouro,
        "numero"         : numero,
        "complemento"    : complemento,
        "cep"            : cep,
        "telefone"       : tel_fmt,
        "inscricao_estadual" : "",   # BrasilAPI não retorna IE
        "inscricao_municipal": "",
        "fonte"          : "BrasilAPI",
    }


def _normalizar_receitaws(data: dict) -> dict:
    """Mapeia resposta ReceitaWS para campos internos."""
    if data.get("status") == "ERROR":
        return {"erro": data.get("message", "CNPJ não encontrado.")}

    return {
        "cnpj"           : so_numeros(str(data.get("cnpj", ""))),
        "razao_social"   : (data.get("nome") or "").strip(),
        "nome_fantasia"  : (data.get("fantasia") or "").strip(),
        "situacao"       : (data.get("situacao") or "").strip(),
        "uf"             : (data.get("uf") or "").strip(),
        "municipio"      : (data.get("municipio") or "").strip(),
        "bairro"         : (data.get("bairro") or "").strip(),
        "logradouro"     : (data.get("logradouro") or "").strip(),
        "numero"         : (data.get("numero") or "").strip(),
        "complemento"    : (data.get("complemento") or "").strip(),
        "cep"            : so_numeros(str(data.get("cep", ""))),
        "telefone"       : _extrair_telefone(data.get("telefone", "")),
        "inscricao_estadual" : "",
        "inscricao_municipal": "",
        "fonte"          : "ReceitaWS",
    }


def _extrair_telefone(raw: str) -> str:
    """Normaliza string de telefone para formato (DD) NNNNN-NNNN."""
    if not raw:
        return ""
    nums = so_numeros(str(raw))
    if len(nums) >= 10:
        ddd    = nums[:2]
        numero = nums[2:]
        if len(numero) == 9:
            return f"({ddd}) {numero[:5]}-{numero[5:]}"
        elif len(numero) == 8:
            return f"({ddd}) {numero[:4]}-{numero[4:]}"
        return f"({ddd}) {numero}"
    return raw.strip()


# ==============================
# GERAÇÃO DO EXCEL (modelo planilhamodelodominio)
# ==============================
COLUNAS_EXCEL = [
    "C: cliente / F: Fornecedor",
    "Razão Social",
    "CPF/CNPJ",
    "Inscrição Estadual",
    "Inscrição Municipal",
    "UF",
    "Município",
    "Bairro",
    "Endereço",
    "Número Endereço",
    "Complemento Endereço",
    "CEP",
    "Telefone",
    "Conta débito",
    "Conta Patrimonial",
]


def gerar_excel(registros: list[dict]) -> bytes:
    """
    Gera Excel no mesmo formato do planilhamodelodominio.xls.
    registros: lista de dicts com campos do cadastro + tipo (C/F).
    """
    rows = []
    for r in registros:
        cnpj_num = so_numeros(r.get("cnpj", ""))
        rows.append({
            "C: cliente / F: Fornecedor": r.get("tipo", "C"),
            "Razão Social"              : r.get("razao_social", ""),
            "CPF/CNPJ"                  : float(cnpj_num) if cnpj_num else "",
            "Inscrição Estadual"        : r.get("inscricao_estadual", ""),
            "Inscrição Municipal"       : r.get("inscricao_municipal", ""),
            "UF"                        : r.get("uf", ""),
            "Município"                 : r.get("municipio", ""),
            "Bairro"                    : r.get("bairro", ""),
            "Endereço"                  : r.get("logradouro", ""),
            "Número Endereço"           : r.get("numero", ""),
            "Complemento Endereço"      : r.get("complemento", ""),
            "CEP"                       : float(so_numeros(r.get("cep", ""))) if so_numeros(r.get("cep", "")) else "",
            "Telefone"                  : r.get("telefone", ""),
            "Conta débito"              : r.get("conta_debito", ""),
            "Conta Patrimonial"         : r.get("conta_patrimonial", ""),
        })

    df = pd.DataFrame(rows, columns=COLUNAS_EXCEL)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Plan1")
        ws = writer.sheets["Plan1"]

        # --- formatação básica de largura ---
        col_widths = {
            "A": 28, "B": 40, "C": 20, "D": 18, "E": 18,
            "F": 6,  "G": 25, "H": 25, "I": 35, "J": 14,
            "K": 22, "L": 12, "M": 20, "N": 14, "O": 16,
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # --- cabeçalho laranja ---
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        header_fill  = PatternFill("solid", fgColor="FF8000")
        header_font  = Font(bold=True, color="FFFFFF", name="Segoe UI", size=10)
        thin_side    = Side(style="thin", color="CCCCCC")
        thin_border  = Border(left=thin_side, right=thin_side,
                              top=thin_side, bottom=thin_side)
        center_align = Alignment(horizontal="center", vertical="center")
        left_align   = Alignment(horizontal="left",   vertical="center")

        for cell in ws[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = center_align
            cell.border    = thin_border

        # --- linhas de dados ---
        row_fills = [
            PatternFill("solid", fgColor="FFFFFF"),
            PatternFill("solid", fgColor="F5F5F5"),
        ]
        data_font = Font(name="Segoe UI", size=10)
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=0):
            fill = row_fills[row_idx % 2]
            for cell in row:
                cell.fill      = fill
                cell.font      = data_font
                cell.alignment = left_align
                cell.border    = thin_border

        ws.freeze_panes = "A2"

    buf.seek(0)
    return buf.read()


# ==============================
# GERAÇÃO DO TXT (formato cli_for.txt)
# ==============================
def gerar_txt(registros: list[dict]) -> bytes:
    """
    Gera arquivo TXT no mesmo padrão do cli_for.txt
    (campos separados por ponto-e-vírgula, uma linha por registro).
    """
    linhas = []
    for r in registros:
        cnpj_fmt = formatar_cnpj(r.get("cnpj", ""))
        cep_fmt  = formatar_cep(r.get("cep", ""))
        campos = [
            r.get("tipo", "C"),
            r.get("razao_social", ""),
            cnpj_fmt,
            r.get("inscricao_estadual", ""),
            r.get("inscricao_municipal", ""),
            r.get("uf", ""),
            r.get("municipio", ""),
            r.get("bairro", ""),
            r.get("logradouro", ""),
            r.get("numero", ""),
            r.get("complemento", ""),
            cep_fmt,
            r.get("telefone", ""),
            r.get("conta_debito", ""),
            r.get("conta_patrimonial", ""),
        ]
        linhas.append(";".join(str(c) for c in campos))
    return ("\n".join(linhas) + "\n").encode("utf-8", errors="replace")


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="Domínio Sistemas | Cadastro de Clientes/Fornecedores",
        page_icon="🟠",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_tr_theme()

    # ---------- cabeçalho ----------
    st.markdown(
        f"""
        <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px;
                    border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
                🏢 Cadastro de Clientes / Fornecedores &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                Consulte CNPJs diretamente da Receita Federal e exporte no formato
                <strong>Domínio Sistemas</strong>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---------- sidebar ----------
    with st.sidebar:
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters**")
        st.markdown("**Domínio Sistemas**")
        st.markdown("---")
        st.markdown("### 📡 Fonte de dados")
        st.markdown(
            "Os dados são consultados em tempo real via "
            "**BrasilAPI** (espelho público da Receita Federal), "
            "com fallback automático para **ReceitaWS**."
        )
        st.markdown("---")
        st.markdown("### ⚙ Campos opcionais")
        st.markdown(
            "Inscrição Estadual, Inscrição Municipal, "
            "Conta Débito e Conta Patrimonial **não são retornados** "
            "pela API pública — preencha manualmente na tabela "
            "antes de exportar."
        )

    # ---------- instruções ----------
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">
            <h4>🔹 Passo 1 — Adicionar CNPJs</h4>
            <p>Digite um CNPJ por linha na caixa de texto e clique em
            <b>▶ Consultar CNPJs</b>. Você pode colar vários de uma vez.</p>

            <h4>🔹 Passo 2 — Definir tipo e contas</h4>
            <p>Após a consulta, selecione se cada empresa é
            <b>Cliente (C)</b> ou <b>Fornecedor (F)</b> e preencha
            os campos <b>Conta Débito</b> e <b>Conta Patrimonial</b>
            diretamente na tabela editável.</p>

            <h4>🔹 Passo 3 — Exportar</h4>
            <ul>
                <li>Clique em <b>⬇ Baixar Excel (.xlsx)</b> para obter a
                planilha no formato <code>planilhamodelodominio</code>.</li>
                <li>Clique em <b>⬇ Baixar TXT</b> para obter o arquivo
                <code>cli_for.txt</code> pronto para importação.</li>
            </ul>

            <h4>🔹 Passo 4 — Importar no Domínio</h4>
            <p>Contabilidade → <b>Utilitários → Importação →
            Clientes/Fornecedores</b>.</p>

            <hr>
            <h4>⚠ Observações</h4>
            <ul>
                <li>A API pública não retorna Inscrição Estadual/Municipal —
                    preencha manualmente.</li>
                <li>Limite de consultas: ~5 por minuto (BrasilAPI).
                    CNPJs são consultados com intervalo automático.</li>
                <li>CNPJs inválidos são sinalizados e não são exportados.</li>
            </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ---------- estado da sessão ----------
    if "registros"   not in st.session_state:
        st.session_state.registros   = []   # lista de dicts
    if "log_cnpj"    not in st.session_state:
        st.session_state.log_cnpj    = ["Aplicação pronta."]
    if "xlsx_bytes"  not in st.session_state:
        st.session_state.xlsx_bytes  = None
    if "txt_bytes"   not in st.session_state:
        st.session_state.txt_bytes   = None

    # ---------- entrada de CNPJs ----------
    st.markdown("#### 📋 Informe os CNPJs (um por linha)")
    cnpjs_input = st.text_area(
        label="CNPJs",
        placeholder="11.222.333/0001-81\n44555666000177\n99.999.999/0001-91",
        height=140,
        label_visibility="collapsed",
    )

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        consultar = st.button(
            "▶ Consultar CNPJs",
            disabled=not cnpjs_input.strip(),
            use_container_width=True,
            type="primary",
        )
    with col_btn2:
        limpar = st.button("🗑 Limpar tudo", use_container_width=True)

    if limpar:
        st.session_state.registros  = []
        st.session_state.log_cnpj   = ["Campos limpos."]
        st.session_state.xlsx_bytes = None
        st.session_state.txt_bytes  = None
        st.rerun()

    # ---------- processamento ----------
    if consultar and cnpjs_input.strip():
        linhas_cnpj = [l.strip() for l in cnpjs_input.splitlines() if l.strip()]
        cnpjs_unicos = list(dict.fromkeys(so_numeros(c) for c in linhas_cnpj if so_numeros(c)))

        # evita re-consultar CNPJs já presentes
        cnpjs_ja_consultados = {so_numeros(r["cnpj"]) for r in st.session_state.registros}
        cnpjs_novos = [c for c in cnpjs_unicos if c not in cnpjs_ja_consultados]

        st.session_state.log_cnpj = [f"Iniciando consulta de {len(cnpjs_novos)} CNPJ(s)..."]

        if cnpjs_novos:
            progress_bar = st.progress(0, text="Consultando...")
            total = len(cnpjs_novos)

            for idx, cnpj in enumerate(cnpjs_novos):
                progress_bar.progress(
                    (idx + 1) / total,
                    text=f"Consultando {formatar_cnpj(cnpj)} ({idx+1}/{total})..."
                )
                resultado = consultar_cnpj(cnpj)

                if "erro" in resultado:
                    st.session_state.log_cnpj.append(
                        f"✗ {formatar_cnpj(cnpj)} → {resultado['erro']}"
                    )
                    # adiciona mesmo com erro para mostrar na tabela
                    st.session_state.registros.append({
                        "cnpj"               : cnpj,
                        "tipo"               : "C",
                        "razao_social"       : f"[ERRO] {resultado['erro']}",
                        "nome_fantasia"      : "",
                        "situacao"           : "ERRO",
                        "uf"                 : "",
                        "municipio"          : "",
                        "bairro"             : "",
                        "logradouro"         : "",
                        "numero"             : "",
                        "complemento"        : "",
                        "cep"                : "",
                        "telefone"           : "",
                        "inscricao_estadual" : "",
                        "inscricao_municipal": "",
                        "conta_debito"       : "",
                        "conta_patrimonial"  : "",
                        "fonte"              : "—",
                        "_erro"              : True,
                    })
                else:
                    resultado["tipo"]              = "C"
                    resultado["inscricao_estadual"]= ""
                    resultado["inscricao_municipal"]= ""
                    resultado["conta_debito"]      = ""
                    resultado["conta_patrimonial"] = ""
                    resultado["_erro"]             = False
                    st.session_state.registros.append(resultado)
                    st.session_state.log_cnpj.append(
                        f"✓ {formatar_cnpj(cnpj)} → {resultado['razao_social']} "
                        f"[{resultado['fonte']}]"
                    )

                # delay para respeitar rate-limit da API pública
                if idx < total - 1:
                    time.sleep(0.8)

            progress_bar.empty()

        else:
            st.session_state.log_cnpj.append("Nenhum CNPJ novo para consultar.")

        st.session_state.xlsx_bytes = None
        st.session_state.txt_bytes  = None
        st.rerun()

    # ---------- tabela de resultados ----------
    if st.session_state.registros:
        st.markdown("---")
        st.markdown("#### 📊 Resultados — edite tipo e contas antes de exportar")

        # monta DataFrame editável
        df_edit = pd.DataFrame([
            {
                "CNPJ"                : formatar_cnpj(r["cnpj"]),
                "Tipo (C/F)"          : r.get("tipo", "C"),
                "Razão Social"        : r.get("razao_social", ""),
                "Situação"            : r.get("situacao", ""),
                "UF"                  : r.get("uf", ""),
                "Município"           : r.get("municipio", ""),
                "Telefone"            : r.get("telefone", ""),
                "Insc. Estadual"      : r.get("inscricao_estadual", ""),
                "Insc. Municipal"     : r.get("inscricao_municipal", ""),
                "Conta Débito"        : r.get("conta_debito", ""),
                "Conta Patrimonial"   : r.get("conta_patrimonial", ""),
            }
            for r in st.session_state.registros
        ])

        edited_df = st.data_editor(
            df_edit,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "CNPJ": st.column_config.TextColumn(
                    "CNPJ", disabled=True, width="medium"
                ),
                "Tipo (C/F)": st.column_config.SelectboxColumn(
                    "Tipo", options=["C", "F"], required=True, width="small"
                ),
                "Razão Social": st.column_config.TextColumn(
                    "Razão Social", width="large"
                ),
                "Situação": st.column_config.TextColumn(
                    "Situação", disabled=True, width="small"
                ),
                "UF": st.column_config.TextColumn(
                    "UF", disabled=True, width="small"
                ),
                "Município": st.column_config.TextColumn(
                    "Município", disabled=True, width="medium"
                ),
                "Telefone": st.column_config.TextColumn(
                    "Telefone", width="medium"
                ),
                "Insc. Estadual": st.column_config.TextColumn(
                    "Insc. Estadual", width="medium"
                ),
                "Insc. Municipal": st.column_config.TextColumn(
                    "Insc. Municipal", width="medium"
                ),
                "Conta Débito": st.column_config.TextColumn(
                    "Conta Débito", width="medium"
                ),
                "Conta Patrimonial": st.column_config.TextColumn(
                    "Conta Patrimonial", width="medium"
                ),
            },
            key="tabela_registros",
        )

        # propaga edições de volta para session_state
        for i, row in edited_df.iterrows():
            if i < len(st.session_state.registros):
                st.session_state.registros[i]["tipo"]               = row["Tipo (C/F)"]
                st.session_state.registros[i]["razao_social"]        = row["Razão Social"]
                st.session_state.registros[i]["telefone"]            = row["Telefone"]
                st.session_state.registros[i]["inscricao_estadual"]  = row["Insc. Estadual"]
                st.session_state.registros[i]["inscricao_municipal"] = row["Insc. Municipal"]
                st.session_state.registros[i]["conta_debito"]        = row["Conta Débito"]
                st.session_state.registros[i]["conta_patrimonial"]   = row["Conta Patrimonial"]

        # ---------- métricas ----------
        total_reg  = len(st.session_state.registros)
        total_ok   = sum(1 for r in st.session_state.registros if not r.get("_erro"))
        total_err  = total_reg - total_ok
        total_cli  = sum(1 for r in st.session_state.registros if r.get("tipo") == "C" and not r.get("_erro"))
        total_forn = sum(1 for r in st.session_state.registros if r.get("tipo") == "F" and not r.get("_erro"))

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total consultados", total_reg)
        c2.metric("✅ Com sucesso",     total_ok)
        c3.metric("👤 Clientes",        total_cli)
        c4.metric("🏭 Fornecedores",    total_forn)
        if total_err:
            st.warning(f"⚠ {total_err} CNPJ(s) com erro — não serão incluídos na exportação.")

        # ---------- botões de exportação ----------
        st.markdown("---")
        st.markdown("#### ⬇ Exportar")

        registros_validos = [r for r in st.session_state.registros if not r.get("_erro")]

        col_x1, col_x2, col_x3 = st.columns([1, 1, 1])

        with col_x1:
            if st.button("📦 Preparar exportação", use_container_width=True, type="primary"):
                if registros_validos:
                    st.session_state.xlsx_bytes = gerar_excel(registros_validos)
                    st.session_state.txt_bytes  = gerar_txt(registros_validos)
                    st.session_state.log_cnpj.append(
                        f"Arquivos gerados: {len(registros_validos)} registro(s)."
                    )
                    st.rerun()
                else:
                    st.warning("Nenhum registro válido para exportar.")

        with col_x2:
            if st.session_state.xlsx_bytes is not None:
                st.download_button(
                    label="⬇ Baixar Excel (.xlsx)",
                    data=st.session_state.xlsx_bytes,
                    file_name="clientes_fornecedores.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )
            else:
                st.button("⬇ Baixar Excel (.xlsx)", disabled=True, use_container_width=True)

        with col_x3:
            if st.session_state.txt_bytes is not None:
                st.download_button(
                    label="⬇ Baixar TXT (cli_for.txt)",
                    data=st.session_state.txt_bytes,
                    file_name="cli_for.txt",
                    mime="text/plain",
                    use_container_width=True,
                    type="primary",
                )
            else:
                st.button("⬇ Baixar TXT", disabled=True, use_container_width=True)

    # ---------- log ----------
    st.markdown("---")
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log_cnpj)
    tem_erro  = any("✗" in str(l) or "ERRO" in str(l) for l in st.session_state.log_cnpj)
    cor_borda = "#D32F2F" if tem_erro else "#388E3C"

    st.markdown(
        f"""
        <div style="background:#FCFCFC; border:1px solid {cor_borda};
                    border-radius:6px; padding:14px;
                    font-family:Consolas,monospace; font-size:13px;
                    white-space:pre-wrap; max-height:280px;
                    overflow-y:auto; color:#1F1F1F;">
{log_texto}
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
