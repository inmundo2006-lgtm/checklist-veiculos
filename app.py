import streamlit as st
import streamlit.components.v1 as components
import requests, json, time, base64
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from itens_checklist import ITENS_CHECKLIST, ITENS_VAN

# ─────────────────────────────────────────────
st.set_page_config(page_title="Checklist Veículos", page_icon="🚗",
                   layout="wide", initial_sidebar_state="collapsed")

TENANT_ID     = st.secrets["TENANT_ID"]
CLIENT_ID     = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
SITE_HOST     = st.secrets.get("SITE_HOST", "metalcana.sharepoint.com")
SITE_NAME     = st.secrets.get("SITE_NAME", "AppKanbanFrotas")
LISTA_FROTAS      = st.secrets.get("LISTA_FROTAS",      "KanbanFrotas")
LISTA_CHECKLIST   = st.secrets.get("LISTA_CHECKLIST",   "ChecklistVeiculos")
LISTA_FOTOS       = st.secrets.get("LISTA_FOTOS",       "ChecklistFotos")
LISTA_PECAS       = st.secrets.get("LISTA_PECAS",       "SolicitacaoPecas")

TIPOS_VEICULO = ["Carro", "Caminhonete", "Van / Sprinter"]
AVALIACOES    = ["✅ Bom", "⚠️ Regular", "❌ Ruim"]
AVAL_COR      = {"✅ Bom": "#d1fae5", "⚠️ Regular": "#fef3c7", "❌ Ruim": "#fee2e2"}
AVAL_TX       = {"✅ Bom": "#065f46", "⚠️ Regular": "#92400e", "❌ Ruim": "#991b1b"}

# ─────────────────────────────────────────────
# GRAPH API
# ─────────────────────────────────────────────
@st.cache_data(ttl=3500)
def get_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type":"client_credentials","client_id":CLIENT_ID,
              "client_secret":CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"})
    r.raise_for_status()
    return r.json()["access_token"]

@st.cache_data(ttl=3500)
def get_site_id():
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/sites/{SITE_HOST}:/sites/{SITE_NAME}",
        headers={"Authorization": f"Bearer {get_token()}"})
    r.raise_for_status()
    return r.json()["id"]

def hdrs():
    return {"Authorization": f"Bearer {get_token()}", "Content-Type": "application/json"}

def lista_items(lista, filtro=""):
    site_id = get_site_id()
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista}/items?expand=fields&$top=500"
    if filtro:
        url += f"&$filter={filtro}"
    r = requests.get(url, headers=hdrs())
    r.raise_for_status()
    return r.json().get("value", [])

def criar_item(lista, fields):
    site_id = get_site_id()
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista}/items"
    r = requests.post(url, headers=hdrs(), json={"fields": fields})
    r.raise_for_status()
    return r.json()

def patch_item(lista, item_id, fields):
    site_id = get_site_id()
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lista}/items/{item_id}/fields"
    r = requests.patch(url, headers=hdrs(), json=fields)
    r.raise_for_status()
    return r.json()

# ─────────────────────────────────────────────
# CARREGAR FROTAS DO KANBAN
# ─────────────────────────────────────────────
@st.cache_data(ttl=60)
def carregar_frotas():
    items = lista_items(LISTA_FROTAS)
    return [{
        "id":          i["id"],
        "nome":        i["fields"].get("Title",""),
        "tipo":        i["fields"].get("Tipo",""),
        "chassi":      i["fields"].get("Chassi",""),
        "ano":         i["fields"].get("Ano",""),
        "cc_nome":     i["fields"].get("CCNome",""),
        "frente_nome": i["fields"].get("FrenteNome",""),
        "status":      i["fields"].get("Status","Ativo"),
    } for i in items
      if i["fields"].get("Tipo","") in ("Veículo","Caminhão","Apoio")
      and i["fields"].get("Status","Ativo") == "Ativo"]

@st.cache_data(ttl=30)
def carregar_checklists(frota_id=""):
    items = lista_items(LISTA_CHECKLIST)
    cl = [{
        "id":            i["id"],
        "frota_nome":    i["fields"].get("FrotaNome",""),
        "frota_id":      i["fields"].get("FrotaId",""),
        "cc_nome":       i["fields"].get("CCNome",""),
        "frente_nome":   i["fields"].get("FrenteNome",""),
        "tipo_veiculo":  i["fields"].get("TipoVeiculo",""),
        "chassi":        i["fields"].get("Chassi",""),
        "ano":           i["fields"].get("Ano",""),
        "data":          i["fields"].get("DataChecklist",""),
        "operador":      i["fields"].get("NomeOperador",""),
        "inspetor":      i["fields"].get("NomeInspetor",""),
        "km":            i["fields"].get("KmAtual",""),
        "motivo":        i["fields"].get("MotivoEntrada",""),
        "resultado":     i["fields"].get("Resultado",""),
        "itens":         i["fields"].get("Itens","{}"),
        "obs":           i["fields"].get("Observacoes",""),
        "status":        i["fields"].get("Status","Rascunho"),
        "created":       i["fields"].get("Created",""),
    } for i in items]
    if frota_id:
        cl = [c for c in cl if c["frota_id"] == frota_id]
    return sorted(cl, key=lambda x: x.get("created",""), reverse=True)

def invalidar():
    carregar_frotas.clear()
    carregar_checklists.clear()
    carregar_pecas.clear()
@st.cache_data(ttl=30)
def carregar_pecas(checklist_id=""):
    items = lista_items(LISTA_PECAS)
    pecas = [{
        "id":               i["id"],
        "checklist_id":     i["fields"].get("ChecklistId",""),
        "checklist_titulo": i["fields"].get("ChecklistTitulo",""),
        "frota_nome":       i["fields"].get("FrotaNome",""),
        "frota_id":         i["fields"].get("FrotaId",""),
        "cc_nome":          i["fields"].get("CCNome",""),
        "inspetor":         i["fields"].get("NomeInspetor",""),
        "nome_peca":        i["fields"].get("NomePeca",""),
        "quantidade":       i["fields"].get("Quantidade",1),
        "urgencia":         i["fields"].get("Urgencia","Normal"),
        "observacao":       i["fields"].get("Observacao",""),
        "status":           i["fields"].get("Status","Solicitado"),
        "comprador":        i["fields"].get("NomeComprador",""),
        "data_compra":      i["fields"].get("DataCompra",""),
        "data_recebimento": i["fields"].get("DataRecebimento",""),
        "data_instalacao":  i["fields"].get("DataInstalacao",""),
        "valor_unitario":   i["fields"].get("ValorUnitario",0),
        "obs_comprador":    i["fields"].get("ObsComprador",""),
        "created":          i["fields"].get("Created",""),
    } for i in items]
    if checklist_id:
        pecas = [p for p in pecas if p["checklist_id"] == checklist_id]
    return sorted(pecas, key=lambda x: x["created"], reverse=True)

STATUS_PECAS  = ["Solicitado","Comprado","Recebido","Instalado"]
URGENCIA_OPTS = ["Normal","Urgente","Crítico"]
URGENCIA_COR  = {"Normal":"#dbeafe","Urgente":"#fef3c7","Crítico":"#fee2e2"}
URGENCIA_TX   = {"Normal":"#1e40af","Urgente":"#92400e","Crítico":"#991b1b"}
STATUS_COR    = {
    "Solicitado": "#f3f4f6", "Comprado":  "#dbeafe",
    "Recebido":   "#d1fae5", "Instalado": "#d1fae5",
}
STATUS_TX     = {
    "Solicitado": "#374151", "Comprado":  "#1e40af",
    "Recebido":   "#065f46", "Instalado": "#065f46",
}
STATUS_EMOJI  = {
    "Solicitado":"🟡","Comprado":"🔵","Recebido":"🟢","Instalado":"✅"
}



# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer, header {visibility:hidden;}
.block-container {padding:1rem 1.5rem !important;}
.item-row {
    display:flex; align-items:center; justify-content:space-between;
    padding:8px 12px; border-radius:8px; margin-bottom:6px;
    background:#fff; border:1px solid #e5e7eb;
}
.item-nome {font-size:13px; color:#111; font-weight:500;}
.cat-header {
    font-size:13px; font-weight:700; padding:8px 12px;
    border-radius:8px; margin:12px 0 6px; color:#fff;
}
.resultado-box {
    border-radius:10px; padding:16px 20px; text-align:center;
    font-size:20px; font-weight:700; margin:8px 0;
}
.hist-card {
    background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px;
    padding:12px 16px; margin-bottom:10px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ESTADO DA SESSÃO
# ─────────────────────────────────────────────
if "checklist_itens" not in st.session_state:
    st.session_state.checklist_itens = {}
if "fotos" not in st.session_state:
    st.session_state.fotos = {}
if "etapa" not in st.session_state:
    st.session_state.etapa = "identificacao"  # identificacao | checklist | revisao | concluido
if "frota_selecionada" not in st.session_state:
    st.session_state.frota_selecionada = None
if "tipo_veiculo" not in st.session_state:
    st.session_state.tipo_veiculo = "Carro"

# ─────────────────────────────────────────────
# TÍTULO E ABAS
# ─────────────────────────────────────────────
st.markdown("### 🚗 Checklist de Veículos — Teston / Metalcana")
aba_novo, aba_hist, aba_pecas, aba_export = st.tabs([
    "📋 Novo Checklist",
    "🕐 Histórico",
    "🔩 Peças",
    "📤 Exportar",
])

# ══════════════════════════════════════════════
# ABA 1 — NOVO CHECKLIST
# ══════════════════════════════════════════════
with aba_novo:

    # ── PROGRESSO ────────────────────────────
    etapas = {"identificacao": 1, "checklist": 2, "revisao": 3, "concluido": 4}
    pct = {1: 25, 2: 60, 3: 90, 4: 100}
    etapa_atual = st.session_state.etapa
    prog = pct.get(etapas.get(etapa_atual, 1), 25)
    labels = ["1. Identificação", "2. Checklist", "3. Revisão", "4. Concluído"]
    col_prog = st.columns(4)
    for i, (col, lbl) in enumerate(zip(col_prog, labels)):
        ativo = etapas.get(etapa_atual, 1) >= i+1
        col.markdown(
            f'<div style="text-align:center;font-size:12px;font-weight:{"700" if ativo else "400"};'
            f'color:{"#1D9E75" if ativo else "#9ca3af"}">{lbl}</div>',
            unsafe_allow_html=True)
    st.progress(prog / 100)
    st.markdown("<br>", unsafe_allow_html=True)

    # ════════════════════════════════════════
    # ETAPA 1 — IDENTIFICAÇÃO
    # ════════════════════════════════════════
    if etapa_atual == "identificacao":
        st.subheader("1️⃣ Identificação do veículo e responsáveis")

        frotas = carregar_frotas()
        nomes_frotas = sorted([f["nome"] for f in frotas])

        col1, col2 = st.columns([2, 1])
        with col1:
            busca_frota = st.text_input("🔍 Buscar frota pelo código ou nome",
                                         placeholder="Ex: 1185 ou TOYOTA")
        with col2:
            tipo_v = st.selectbox("Tipo de veículo", TIPOS_VEICULO)

        frotas_filtradas = [f for f in frotas
                            if not busca_frota or busca_frota.lower() in f["nome"].lower()]
        nomes_filtrados = ["— Selecione —"] + [f["nome"] for f in frotas_filtradas]

        frota_nome = st.selectbox("Frota", nomes_filtrados)
        frota_obj  = next((f for f in frotas if f["nome"] == frota_nome), None)

        if frota_obj:
            st.markdown("**📍 Dados do Kanban:**")
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Centro de Custo", frota_obj["cc_nome"] or "—")
            c2.metric("Frente",          frota_obj["frente_nome"] or "—")
            c3.metric("Chassi",          frota_obj["chassi"] or "—")
            c4.metric("Ano",             frota_obj["ano"] or "—")

        st.divider()
        c1,c2,c3 = st.columns(3)
        with c1: operador = st.text_input("👤 Nome do operador / motorista")
        with c2: inspetor  = st.text_input("🔍 Nome do inspetor")
        with c3: km_atual  = st.number_input("🛣️ KM atual", min_value=0, step=1)

        motivo = st.selectbox("📝 Motivo da entrada na oficina", [
            "Manutenção preventiva", "Manutenção corretiva",
            "Revisão periódica", "Troca de óleo", "Checklist de rotina",
            "Após acidente / sinistro", "Outro",
        ])
        if motivo == "Outro":
            motivo = st.text_input("Descreva o motivo")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("▶️ Iniciar Checklist", type="primary", use_container_width=True):
            if frota_nome == "— Selecione —":
                st.error("Selecione uma frota.")
            elif not operador.strip():
                st.error("Informe o nome do operador.")
            elif not inspetor.strip():
                st.error("Informe o nome do inspetor.")
            else:
                st.session_state.frota_selecionada = frota_obj
                st.session_state.tipo_veiculo = tipo_v
                st.session_state.dados_ident = {
                    "operador": operador.strip(),
                    "inspetor": inspetor.strip(),
                    "km": km_atual,
                    "motivo": motivo,
                }
                st.session_state.checklist_itens = {}
                st.session_state.fotos = {}
                st.session_state.etapa = "checklist"
                st.rerun()

    # ════════════════════════════════════════
    # ETAPA 2 — CHECKLIST
    # ════════════════════════════════════════
    elif etapa_atual == "checklist":
        frota = st.session_state.frota_selecionada
        ident = st.session_state.dados_ident
        tipo_v = st.session_state.tipo_veiculo

        st.markdown(
            f'<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;'
            f'padding:10px 16px;margin-bottom:16px">'
            f'<b>🚗 {frota["nome"]}</b> &nbsp;|&nbsp; {frota["cc_nome"] or "Sem CC"} '
            f'&nbsp;|&nbsp; {tipo_v} &nbsp;|&nbsp; Operador: {ident["operador"]} '
            f'&nbsp;|&nbsp; Inspetor: {ident["inspetor"]}</div>',
            unsafe_allow_html=True)

        # Monta categorias
        cats = dict(ITENS_CHECKLIST)
        if tipo_v == "Van / Sprinter":
            cats.update(ITENS_VAN)

        CORES_CAT = ["#1D9E75","#378ADD","#BA7517","#D85A30","#7F77DD",
                     "#D4537E","#639922","#0F6E56","#185FA5"]

        itens_state = st.session_state.checklist_itens
        fotos_state = st.session_state.fotos

        for ci, (cat, itens) in enumerate(cats.items()):
            cor = CORES_CAT[ci % len(CORES_CAT)]
            st.markdown(
                f'<div class="cat-header" style="background:{cor}">{cat}</div>',
                unsafe_allow_html=True)

            for item in itens:
                key_aval = f"aval_{cat}_{item}"
                key_obs  = f"obs_{cat}_{item}"
                key_foto = f"foto_{cat}_{item}"

                c1, c2 = st.columns([3, 2])
                with c1:
                    st.markdown(f'<div style="font-size:13px;padding:8px 0">{item}</div>',
                                unsafe_allow_html=True)
                with c2:
                    aval = st.radio(
                        label=item, options=AVALIACOES,
                        key=key_aval, horizontal=True,
                        label_visibility="collapsed",
                        index=0,
                    )

                itens_state[f"{cat}||{item}"] = aval

                if aval == "❌ Ruim":
                    obs_col, foto_col = st.columns([2,1])
                    with obs_col:
                        obs = st.text_input(f"Observação — {item[:30]}...",
                                            key=key_obs, placeholder="Descreva o problema")
                        itens_state[f"{cat}||{item}__obs"] = obs
                    with foto_col:
                        foto = st.camera_input(f"📷 Foto obrigatória", key=key_foto)
                        if foto:
                            fotos_state[f"{cat}||{item}"] = base64.b64encode(foto.read()).decode()
                            st.success("Foto capturada ✅")
                        elif f"{cat}||{item}" not in fotos_state:
                            st.warning("⚠️ Foto obrigatória")

                st.markdown('<hr style="margin:4px 0;border-color:#f3f4f6">', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        obs_geral = st.text_area("📝 Observações gerais (opcional)",
                                  placeholder="Anotações adicionais sobre o veículo...")
        st.session_state.obs_geral = obs_geral

        # Verifica fotos obrigatórias
        itens_ruins = [k for k,v in itens_state.items()
                       if v == "❌ Ruim" and not k.endswith("__obs")]
        fotos_faltando = [k for k in itens_ruins if k not in fotos_state]

        col_back, col_next = st.columns(2)
        with col_back:
            if st.button("◀️ Voltar", use_container_width=True):
                st.session_state.etapa = "identificacao"; st.rerun()
        with col_next:
            if st.button("▶️ Revisar e Finalizar", type="primary", use_container_width=True):
                if fotos_faltando:
                    st.error(f"⚠️ {len(fotos_faltando)} item(ns) marcado(s) como Ruim precisam de foto obrigatória.")
                    for f in fotos_faltando[:3]:
                        cat, item = f.split("||")
                        st.warning(f"📷 Foto faltando: **{cat}** → {item}")
                else:
                    st.session_state.etapa = "revisao"; st.rerun()

    # ════════════════════════════════════════
    # ETAPA 3 — REVISÃO
    # ════════════════════════════════════════
    elif etapa_atual == "revisao":
        frota  = st.session_state.frota_selecionada
        ident  = st.session_state.dados_ident
        itens  = st.session_state.checklist_itens
        fotos  = st.session_state.fotos
        obs_g  = st.session_state.get("obs_geral","")

        st.subheader("3️⃣ Revisão do checklist")

        # Calcula resultado
        avals = {k:v for k,v in itens.items() if not k.endswith("__obs")}
        n_bom  = sum(1 for v in avals.values() if v == "✅ Bom")
        n_reg  = sum(1 for v in avals.values() if v == "⚠️ Regular")
        n_ruim = sum(1 for v in avals.values() if v == "❌ Ruim")
        total  = len(avals)

        if n_ruim == 0 and n_reg == 0:
            resultado = "✅ Aprovado"
            res_cor   = "#d1fae5"; res_tx = "#065f46"
        elif n_ruim == 0:
            resultado = "⚠️ Aprovado com pendências"
            res_cor   = "#fef3c7"; res_tx = "#92400e"
        else:
            resultado = "❌ Reprovado — requer atenção"
            res_cor   = "#fee2e2"; res_tx = "#991b1b"

        # Resumo
        st.markdown(
            f'<div class="resultado-box" style="background:{res_cor};color:{res_tx}">'
            f'{resultado}</div>', unsafe_allow_html=True)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("Total de itens", total)
        c2.metric("✅ Bom", n_bom)
        c3.metric("⚠️ Regular", n_reg)
        c4.metric("❌ Ruim", n_ruim)

        st.divider()

        # Dados identificação
        st.markdown("**📋 Identificação**")
        ic1,ic2,ic3,ic4 = st.columns(4)
        ic1.markdown(f"**Frota:** {frota['nome']}")
        ic2.markdown(f"**CC:** {frota['cc_nome'] or '—'}")
        ic3.markdown(f"**Operador:** {ident['operador']}")
        ic4.markdown(f"**Inspetor:** {ident['inspetor']}")

        st.divider()

        # Itens com problema
        itens_problema = {k:v for k,v in avals.items() if v != "✅ Bom"}
        if itens_problema:
            st.markdown("**⚠️ Itens com problema:**")
            for chave, aval in itens_problema.items():
                cat, item = chave.split("||")
                obs = itens.get(f"{chave}__obs","")
                tem_foto = chave in fotos
                bg = AVAL_COR[aval]; tx = AVAL_TX[aval]
                st.markdown(
                    f'<div style="background:{bg};border-radius:8px;padding:8px 12px;'
                    f'margin-bottom:6px;color:{tx}">'
                    f'<b>{cat}</b> → {item} &nbsp; <b>{aval}</b>'
                    f'{(" &nbsp;|&nbsp; " + obs) if obs else ""}'
                    f'{" &nbsp; 📷" if tem_foto else ""}'
                    f'</div>', unsafe_allow_html=True)
        else:
            st.success("Todos os itens estão em bom estado!")

        if obs_g:
            st.divider()
            st.markdown(f"**Observações gerais:** {obs_g}")

        # ── SOLICITAÇÃO DE PEÇAS ─────────────
        st.divider()
        st.markdown("**🔩 Solicitar peças para este checklist**")
        st.caption("Adicione as peças necessárias antes de finalizar.")

        if "pecas_solicitadas" not in st.session_state:
            st.session_state.pecas_solicitadas = []

        with st.form("form_add_peca", clear_on_submit=True):
            pc1, pc2, pc3, pc4 = st.columns([3,1,1,1])
            with pc1: nome_peca  = st.text_input("Nome da peça", placeholder="Ex: Pastilha de freio dianteira")
            with pc2: qtd_peca   = st.number_input("Qtd", min_value=1, value=1, step=1)
            with pc3: urg_peca   = st.selectbox("Urgência", URGENCIA_OPTS)
            with pc4:
                st.markdown("<br>", unsafe_allow_html=True)
                add_peca = st.form_submit_button("➕ Adicionar", use_container_width=True)
            obs_peca = st.text_input("Observação da peça (opcional)", placeholder="Referência, modelo específico...")
            if add_peca and nome_peca.strip():
                st.session_state.pecas_solicitadas.append({
                    "nome": nome_peca.strip(),
                    "quantidade": qtd_peca,
                    "urgencia": urg_peca,
                    "obs": obs_peca.strip(),
                })
                st.rerun()

        if st.session_state.pecas_solicitadas:
            for pi, peca in enumerate(st.session_state.pecas_solicitadas):
                urg_bg = URGENCIA_COR.get(peca["urgencia"],"#f3f4f6")
                urg_tx = URGENCIA_TX.get(peca["urgencia"],"#374151")
                pc_col, rm_col = st.columns([5,1])
                with pc_col:
                    st.markdown(
                        f'<div style="background:{urg_bg};border-radius:8px;padding:7px 12px;'
                        f'margin-bottom:4px;color:{urg_tx};font-size:13px">'
                        f'<b>{peca["nome"]}</b> &nbsp;|&nbsp; Qtd: {peca["quantidade"]} '
                        f'&nbsp;|&nbsp; {peca["urgencia"]}'
                        f'{(" &nbsp;|&nbsp; " + peca["obs"]) if peca["obs"] else ""}'
                        f'</div>', unsafe_allow_html=True)
                with rm_col:
                    if st.button("🗑️", key=f"rm_peca_{pi}"):
                        st.session_state.pecas_solicitadas.pop(pi)
                        st.rerun()
        else:
            st.info("Nenhuma peça adicionada ainda.")

        st.divider()
        cb, cc = st.columns(2)
        with cb:
            if st.button("◀️ Voltar ao checklist", use_container_width=True):
                st.session_state.etapa = "checklist"; st.rerun()
        with cc:
            if st.button("✅ Confirmar e Salvar", type="primary", use_container_width=True):
                with st.spinner("Salvando no SharePoint..."):
                    try:
                        # Salva checklist principal
                        item_json = json.dumps(
                            {k:v for k,v in itens.items()}, ensure_ascii=False)
                        novo = criar_item(LISTA_CHECKLIST, {
                            "Title":         f"{frota['nome']} — {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                            "FrotaNome":     frota["nome"],
                            "FrotaId":       frota["id"],
                            "CCNome":        frota["cc_nome"],
                            "FrenteNome":    frota["frente_nome"],
                            "TipoVeiculo":   st.session_state.tipo_veiculo,
                            "Chassi":        frota["chassi"],
                            "Ano":           frota["ano"],
                            "DataChecklist": datetime.now().isoformat(),
                            "NomeOperador":  ident["operador"],
                            "NomeInspetor":  ident["inspetor"],
                            "KmAtual":       ident["km"],
                            "MotivoEntrada": ident["motivo"],
                            "Resultado":     resultado,
                            "Itens":         item_json[:3900],  # limite SP
                            "Observacoes":   obs_g,
                            "Status":        "Concluído",
                        })
                        cl_id = novo["id"]

                        # Salva fotos
                        for chave, foto_b64 in fotos.items():
                            cat, item = chave.split("||")
                            aval = itens.get(chave,"")
                            obs_item = itens.get(f"{chave}__obs","")
                            criar_item(LISTA_FOTOS, {
                                "Title":       f"{frota['nome']} — {item[:50]}",
                                "ChecklistId": cl_id,
                                "FrotaNome":   frota["nome"],
                                "Item":        item,
                                "Categoria":   cat,
                                "Avaliacao":   aval,
                                "FotoBase64":  foto_b64[:3900],
                                "Observacao":  obs_item,
                            })
                            time.sleep(0.1)

                        # Salva peças solicitadas
                        pecas_para_salvar = st.session_state.get("pecas_solicitadas",[])
                        for peca in pecas_para_salvar:
                            criar_item(LISTA_PECAS, {
                                "Title":            f"{frota['nome']} — {peca['nome'][:50]}",
                                "ChecklistId":      cl_id,
                                "ChecklistTitulo":  f"{frota['nome']} — {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                                "FrotaNome":        frota["nome"],
                                "FrotaId":          frota["id"],
                                "CCNome":           frota["cc_nome"],
                                "NomeInspetor":     ident["inspetor"],
                                "NomePeca":         peca["nome"],
                                "Quantidade":       peca["quantidade"],
                                "Urgencia":         peca["urgencia"],
                                "Observacao":       peca["obs"],
                                "Status":           "Solicitado",
                            })
                            time.sleep(0.1)
                        st.session_state.pecas_solicitadas = []

                        invalidar()
                        st.session_state.ultimo_checklist_id = cl_id
                        st.session_state.ultimo_resultado    = resultado
                        st.session_state.etapa = "concluido"
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar: {e}")

    # ════════════════════════════════════════
    # ETAPA 4 — CONCLUÍDO
    # ════════════════════════════════════════
    elif etapa_atual == "concluido":
        frota   = st.session_state.frota_selecionada
        result  = st.session_state.get("ultimo_resultado","")
        res_cor = "#d1fae5" if "Aprovado" in result and "pendências" not in result else \
                  "#fef3c7" if "pendências" in result else "#fee2e2"
        res_tx  = "#065f46" if "Aprovado" in result and "pendências" not in result else \
                  "#92400e" if "pendências" in result else "#991b1b"

        st.markdown(
            f'<div class="resultado-box" style="background:{res_cor};color:{res_tx};font-size:24px">'
            f'🎉 Checklist salvo com sucesso!<br>'
            f'<span style="font-size:16px">{result}</span></div>',
            unsafe_allow_html=True)
        st.markdown(f"**Veículo:** {frota['nome']}  |  "
                    f"**Data:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        st.divider()
        if st.button("➕ Novo Checklist", type="primary", use_container_width=True):
            st.session_state.etapa = "identificacao"
            st.session_state.checklist_itens = {}
            st.session_state.fotos = {}
            st.session_state.frota_selecionada = None
            st.rerun()

# ══════════════════════════════════════════════
# ABA 2 — HISTÓRICO
# ══════════════════════════════════════════════
with aba_hist:
    st.subheader("🕐 Histórico de Checklists")

    hf1, hf2, hf3 = st.columns([2,2,1])
    with hf1:
        busca_h = st.text_input("🔍 Buscar frota", placeholder="Nome ou código...", key="bh")
    with hf2:
        filtro_res = st.selectbox("Resultado", ["Todos","✅ Aprovado","⚠️ Pendências","❌ Reprovado"], key="fr")
    with hf3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Atualizar", use_container_width=True):
            invalidar(); st.rerun()

    try:
        todos_cl = carregar_checklists()
    except Exception as e:
        st.error(f"Erro ao carregar histórico: {e}")
        todos_cl = []

    # Filtra
    if busca_h:
        todos_cl = [c for c in todos_cl if busca_h.lower() in c["frota_nome"].lower()]
    if filtro_res != "Todos":
        todos_cl = [c for c in todos_cl if filtro_res.split()[1] in c["resultado"]]

    st.caption(f"{len(todos_cl)} checklist(s) encontrado(s)")
    st.divider()

    for cl in todos_cl[:50]:
        res = cl["resultado"]
        res_cor = "#d1fae5" if "Aprovado" in res and "pendências" not in res else \
                  "#fef3c7" if "pendências" in res else "#fee2e2"
        res_tx  = "#065f46" if "Aprovado" in res and "pendências" not in res else \
                  "#92400e" if "pendências" in res else "#991b1b"

        with st.expander(
            f"🚗 {cl['frota_nome']} — {cl['data'][:10] if cl['data'] else cl['created'][:10]} "
            f"| {res}"
        ):
            cc1,cc2,cc3,cc4 = st.columns(4)
            cc1.markdown(f"**CC:** {cl['cc_nome'] or '—'}")
            cc2.markdown(f"**Frente:** {cl['frente_nome'] or '—'}")
            cc3.markdown(f"**Operador:** {cl['operador']}")
            cc4.markdown(f"**Inspetor:** {cl['inspetor']}")

            st.markdown(f"**KM:** {cl['km']}  |  **Motivo:** {cl['motivo']}")

            # Mostra itens com problema
            try:
                itens_dict = json.loads(cl["itens"]) if cl["itens"] else {}
                problemas  = {k:v for k,v in itens_dict.items()
                              if v in ("⚠️ Regular","❌ Ruim") and not k.endswith("__obs")}
                if problemas:
                    st.markdown("**Itens com problema:**")
                    for chave, aval in problemas.items():
                        parts = chave.split("||")
                        item_nome = parts[1] if len(parts)>1 else chave
                        obs_item  = itens_dict.get(f"{chave}__obs","")
                        bg = AVAL_COR.get(aval,"#f3f4f6")
                        tx = AVAL_TX.get(aval,"#374151")
                        st.markdown(
                            f'<div style="background:{bg};border-radius:6px;padding:5px 10px;'
                            f'margin-bottom:4px;font-size:12px;color:{tx}">'
                            f'<b>{aval}</b> — {item_nome}'
                            f'{(" &nbsp;|&nbsp; " + obs_item) if obs_item else ""}</div>',
                            unsafe_allow_html=True)
                else:
                    st.success("Todos os itens aprovados")
            except:
                pass

            if cl["obs"]:
                st.markdown(f"**Obs gerais:** {cl['obs']}")

# ══════════════════════════════════════════════
# ABA 3 — PEÇAS
# ══════════════════════════════════════════════
with aba_pecas:
    st.subheader("🔩 Solicitação de Peças")

    # ── Filtros ──────────────────────────────
    pp1, pp2, pp3 = st.columns([2,2,1])
    with pp1:
        busca_p = st.text_input("🔍 Buscar frota ou peça", key="bp")
    with pp2:
        filtro_status_p = st.selectbox("Status", ["Todos"] + STATUS_PECAS, key="fsp")
    with pp3:
        filtro_urg_p = st.selectbox("Urgência", ["Todas"] + URGENCIA_OPTS, key="fup")

    pp4, pp5 = st.columns([2,1])
    with pp4:
        filtro_cc_p = st.selectbox("Centro de Custo", ["Todos os CCs"] + 
                                    list(dict.fromkeys([p["cc_nome"] for p in carregar_pecas() if p["cc_nome"]])), 
                                    key="fcp")
    with pp5:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 Atualizar", key="ref_p", use_container_width=True):
            invalidar(); st.rerun()

    try:
        todas_pecas = carregar_pecas()
    except Exception as e:
        st.error(f"Erro: {e}"); todas_pecas = []

    if busca_p:
        todas_pecas = [p for p in todas_pecas 
                       if busca_p.lower() in p["frota_nome"].lower() 
                       or busca_p.lower() in p["nome_peca"].lower()]
    if filtro_status_p != "Todos":
        todas_pecas = [p for p in todas_pecas if p["status"] == filtro_status_p]
    if filtro_urg_p != "Todas":
        todas_pecas = [p for p in todas_pecas if p["urgencia"] == filtro_urg_p]
    if filtro_cc_p != "Todos os CCs":
        todas_pecas = [p for p in todas_pecas if p["cc_nome"] == filtro_cc_p]

    # Métricas rápidas
    m1,m2,m3,m4 = st.columns(4)
    for col, status, emoji in [
        (m1,"Solicitado","🟡"), (m2,"Comprado","🔵"),
        (m3,"Recebido","🟢"),   (m4,"Instalado","✅")
    ]:
        n = sum(1 for p in todas_pecas if p["status"] == status)
        col.markdown(
            f'<div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:8px;'
            f'padding:8px 14px;text-align:center">'
            f'<div style="font-size:22px;font-weight:700">{n}</div>'
            f'<div style="font-size:11px;color:#6b7280">{emoji} {status}</div></div>',
            unsafe_allow_html=True)

    st.markdown(f"<br>*{len(todas_pecas)} peça(s) encontrada(s)*", unsafe_allow_html=True)
    st.divider()

    # ── Lista de peças ────────────────────────
    for peca in todas_pecas:
        urg_bg = URGENCIA_COR.get(peca["urgencia"],"#f3f4f6")
        urg_tx = URGENCIA_TX.get(peca["urgencia"],"#374151")
        st_bg  = STATUS_COR.get(peca["status"],"#f3f4f6")
        st_tx  = STATUS_TX.get(peca["status"],"#374151")
        emoji  = STATUS_EMOJI.get(peca["status"],"🟡")

        with st.expander(
            f'{emoji} **{peca["nome_peca"]}** — {peca["frota_nome"]} '
            f'| Qtd: {peca["quantidade"]} | {peca["urgencia"]} | {peca["status"]}'
        ):
            ec1,ec2,ec3 = st.columns(3)
            ec1.markdown(f"**Frota:** {peca['frota_nome']}")
            ec1.markdown(f"**CC:** {peca['cc_nome'] or '—'}")
            ec2.markdown(f"**Inspetor:** {peca['inspetor']}")
            ec2.markdown(f"**Checklist:** {peca['checklist_titulo']}")
            ec3.markdown(f"**Solicitado em:** {peca['created'][:10]}")
            if peca["observacao"]:
                ec3.markdown(f"**Obs:** {peca['observacao']}")

            st.markdown("**Atualizar status:**")
            with st.form(f"form_status_{peca['id']}"):
                sc1,sc2,sc3,sc4,sc5 = st.columns([2,1,1,2,2])
                with sc1:
                    novo_status = st.selectbox("Status", STATUS_PECAS,
                                               index=STATUS_PECAS.index(peca["status"]) 
                                               if peca["status"] in STATUS_PECAS else 0,
                                               key=f"ns_{peca['id']}")
                with sc2:
                    val_unit = st.number_input("Valor unit. R$", min_value=0.0,
                                               value=float(peca["valor_unitario"] or 0),
                                               step=0.01, key=f"vu_{peca['id']}")
                with sc3:
                    nome_comp = st.text_input("Comprador", value=peca["comprador"],
                                              key=f"nc_{peca['id']}")
                with sc4:
                    obs_comp = st.text_input("Obs do comprador", value=peca["obs_comprador"],
                                             key=f"oc_{peca['id']}")
                with sc5:
                    st.markdown("<br>", unsafe_allow_html=True)
                    salvar_st = st.form_submit_button("💾 Salvar", use_container_width=True)

                if salvar_st:
                    fields_upd = {
                        "Status":          novo_status,
                        "NomeComprador":   nome_comp,
                        "ValorUnitario":   val_unit,
                        "ObsComprador":    obs_comp,
                    }
                    # Preenche datas automáticas por status
                    hoje_iso = datetime.now().strftime("%Y-%m-%d")
                    if novo_status == "Comprado"  and not peca["data_compra"]:
                        fields_upd["DataCompra"]      = hoje_iso
                    if novo_status == "Recebido"  and not peca["data_recebimento"]:
                        fields_upd["DataRecebimento"] = hoje_iso
                    if novo_status == "Instalado" and not peca["data_instalacao"]:
                        fields_upd["DataInstalacao"]  = hoje_iso
                    try:
                        patch_item(LISTA_PECAS, peca["id"], fields_upd)
                        invalidar()
                        st.success("Status atualizado!"); st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")

# ══════════════════════════════════════════════
# ABA 4 — EXPORTAR
# ══════════════════════════════════════════════
with aba_export:
    st.subheader("📤 Exportar relatório de checklists")

    ex1, ex2 = st.columns([2,2])
    with ex1:
        busca_ex = st.text_input("Filtrar por frota", placeholder="Nome ou código...")
    with ex2:
        periodo  = st.selectbox("Período", ["Todos","Hoje","Últimos 7 dias","Últimos 30 dias"])

    try:
        todos_ex = carregar_checklists()
    except:
        todos_ex = []

    if busca_ex:
        todos_ex = [c for c in todos_ex if busca_ex.lower() in c["frota_nome"].lower()]
    if periodo == "Hoje":
        hoje = datetime.now().strftime("%Y-%m-%d")
        todos_ex = [c for c in todos_ex if c["created"][:10] == hoje]
    elif periodo == "Últimos 7 dias":
        from datetime import timedelta
        limite = (datetime.now() - timedelta(days=7)).isoformat()
        todos_ex = [c for c in todos_ex if c["created"] >= limite]
    elif periodo == "Últimos 30 dias":
        from datetime import timedelta
        limite = (datetime.now() - timedelta(days=30)).isoformat()
        todos_ex = [c for c in todos_ex if c["created"] >= limite]

    st.caption(f"{len(todos_ex)} checklist(s) para exportar")

    if st.button("⬇️ Gerar Excel", type="primary", use_container_width=True) and todos_ex:
        wb = Workbook(); wb.remove(wb.active)

        # Aba resumo
        ws = wb.create_sheet("Resumo")
        ws.merge_cells("A1:J1")
        ws["A1"] = f"RELATÓRIO DE CHECKLISTS — {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        ws["A1"].font = Font(bold=True,size=12,color="FFFFFF")
        ws["A1"].fill = PatternFill("solid",fgColor="1D9E75")
        ws["A1"].alignment = Alignment(horizontal="center",vertical="center")
        ws.row_dimensions[1].height = 26

        hdrs_r = ["Data","Frota","Tipo","CC","Frente","Operador","Inspetor","KM","Motivo","Resultado"]
        for ci,h in enumerate(hdrs_r,1): ws.cell(row=2,column=ci,value=h)
        fill_h = PatternFill("solid",fgColor="0F6E56")
        for ci in range(1,11):
            ws.cell(row=2,column=ci).fill  = fill_h
            ws.cell(row=2,column=ci).font  = Font(bold=True,color="FFFFFF",size=10)
            ws.cell(row=2,column=ci).alignment = Alignment(horizontal="center")

        for ri, cl in enumerate(todos_ex, 3):
            data_str = cl["data"][:10] if cl["data"] else cl["created"][:10]
            linha = [data_str, cl["frota_nome"], cl["tipo_veiculo"],
                     cl["cc_nome"], cl["frente_nome"], cl["operador"],
                     cl["inspetor"], cl["km"], cl["motivo"], cl["resultado"]]
            for ci,v in enumerate(linha,1): ws.cell(row=ri,column=ci,value=v)
            # Cor pelo resultado
            res_bg = "d1fae5" if "Aprovado" in cl["resultado"] and "pendências" not in cl["resultado"] else \
                     "fef3c7" if "pendências" in cl["resultado"] else "fee2e2"
            for ci in range(1,11):
                ws.cell(row=ri,column=ci).fill = PatternFill("solid",fgColor=res_bg)
                ws.cell(row=ri,column=ci).font = Font(size=10)
                ws.cell(row=ri,column=ci).border = Border(
                    left=Side("thin","E5E7EB"),right=Side("thin","E5E7EB"),
                    top=Side("thin","E5E7EB"),bottom=Side("thin","E5E7EB"))

        for ci,w in zip(range(1,11),[12,30,14,26,20,18,18,10,22,28]):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A3"

        buf = io.BytesIO(); wb.save(buf); buf.seek(0)
        nome = f"checklists_{datetime.now().strftime('%Y%m%d')}.xlsx"
        st.download_button("📥 Baixar Excel", buf, nome,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True, type="primary")
