"""
integracao_os.py — Ponte Checklist de Veículos → Assistência Técnica Teston

Este módulo vive no repositório do APP DE CHECKLIST. Ele grava direto na
lista AT_Teston_OS (a mesma que o app da oficina usa), sem passar pela tela
da oficina.

Regras de negócio implementadas aqui:
  · OS vinda do checklist é SEMPRE interna (tipo_os = "interna")
  · origem = "checklist" — é o que autoriza a abertura de OS interna sem
    ser supervisor/admin
  · tecnico_designado = o inspetor que fez o checklist (é o mecânico da oficina)
  · avaliacao = montada automaticamente a partir dos itens com problema
  · idempotente: um checklist gera no máximo uma OS (coluna OSNumero)

Secrets necessários no .streamlit/secrets.toml do app de checklist
(os quatro primeiros já existem; os dois últimos são novos):

    TENANT_ID     = "..."
    CLIENT_ID     = "..."
    CLIENT_SECRET = "..."
    OS_SITE_ID    = "..."   # SITE_ID do site onde está a lista AT_Teston_OS
    OS_LIST_ID    = "..."   # LIST_ID da lista AT_Teston_OS

Se a lista AT_Teston_OS estiver no MESMO site do checklist (AppKanbanFrotas),
OS_SITE_ID é o id desse site — o mesmo que get_site_id() devolve.
"""

import json
import unicodedata
from datetime import datetime

import requests
import streamlit as st


# ─────────────────────────────────────────────
#  CENTROS DE CUSTO — espelho do dados.py da oficina
#  Mantenha os dois em sincronia (ver nota no fim do arquivo).
# ─────────────────────────────────────────────

CENTROS_CUSTO = {
    1:  "SERVIÇOS PARTICULARES",
    3:  "COLHEITA COOPERVAL",
    4:  "VALE DO IVAI (RENUKA)",
    5:  "NOVA PRODUTIVA COLHEITA",
    7:  "TRANSPORTE",
    11: "TESTON",
    14: "AGRO CIANORTE",
    16: "ESTOQUE",
    21: "ASSISTÊNCIA ELÉTRICA",
    22: "ASSISTÊNCIA TÉCNICA TESTON",
    23: "PLANTIO MECÂNICO",
    28: "AGRO VALE DO IVAI",
    29: "RIO AMAMBAI COLHEITA",
    37: "AGRO NAVIRAÍ",
    38: "AGRO ASTORGA PLANTIO/CUL",
    41: "RAIZEN",
    42: "CONCESSIONÁRIA SERTÃOZINHO",
    44: "SOL NASCENTE",
    46: "SERVIÇOS DE PREPARO DE SOLO",
    49: "AGRO SANTA CÂNDIDA",
    50: "LOBO GUARÁ",
    51: "COLHEITA COGO",
    54: "AGRO RORAIMA",
    100: "METALCANA",
}

# Técnicos da oficina — precisa bater com o TECNICOS do dados.py.
# É esta lista que alimenta o selectbox de inspetor no checklist, para que
# o nome digitado vire um cod_tecnico exato em vez de texto livre.
TECNICOS_OFICINA = {
    1:  "Zaqueu",
    2:  "Cristiano B.",
    3:  "Rodrigo J.",
    4:  "Marcos C.",
    5:  "Cristiano dos S.",
    6:  "João P.",
    7:  "Loran H.",
    8:  "Bruno H.",
    9:  "Gabriel",
    10: "Tiago Bardu",
    11: "Felipe",
    12: "Gustavo",
    13: "Vanderlei",
    14: "Claudecir",
    15: "Rodrigo M.",
    16: "José",
}


# ─────────────────────────────────────────────
#  GRAPH API — lista AT_Teston_OS
# ─────────────────────────────────────────────

def _cfg(chave: str, padrao=None):
    valor = st.secrets.get(chave, padrao)
    if valor is None:
        raise RuntimeError(
            f"Secret '{chave}' não configurado. Adicione ao secrets.toml "
            "do app de checklist (veja o cabeçalho de integracao_os.py)."
        )
    return valor


@st.cache_data(ttl=3000)
def _token() -> str:
    r = requests.post(
        f"https://login.microsoftonline.com/{_cfg('TENANT_ID')}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     _cfg("CLIENT_ID"),
            "client_secret": _cfg("CLIENT_SECRET"),
            "scope":         "https://graph.microsoft.com/.default",
        },
        timeout=15,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {_token()}",
        "Content-Type":  "application/json",
        "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly",
    }


def _base_url() -> str:
    return (
        "https://graph.microsoft.com/v1.0"
        f"/sites/{_cfg('OS_SITE_ID')}"
        f"/lists/{_cfg('OS_LIST_ID')}"
        "/items"
    )


def _checar(r: requests.Response) -> None:
    if not r.ok:
        try:
            detalhe = r.json().get("error", {}).get("message", r.text[:300])
        except Exception:
            detalhe = r.text[:300]
        raise RuntimeError(f"Graph API {r.status_code}: {detalhe}")


def _escapar(valor: str) -> str:
    return str(valor).replace("'", "''")


def _itens_por_titulo(numero_os: str) -> list[dict]:
    url = _base_url() + f"?$expand=fields&$filter=fields/Title eq '{_escapar(numero_os)}'"
    r = requests.get(url, headers=_headers(), timeout=15)
    _checar(r)
    return r.json().get("value", [])


def _itens_por_prefixo(prefixo: str) -> list[dict]:
    url = (_base_url() + "?$expand=fields&$top=999"
           + f"&$filter=startswith(fields/Title,'{_escapar(prefixo)}')")
    itens = []
    try:
        while url:
            r = requests.get(url, headers=_headers(), timeout=20)
            r.raise_for_status()
            data = r.json()
            itens.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return itens
    except requests.HTTPError:
        url = _base_url() + "?$expand=fields&$top=999"
        itens = []
        while url:
            r = requests.get(url, headers=_headers(), timeout=20)
            _checar(r)
            data = r.json()
            itens.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return [i for i in itens
                if str(i.get("fields", {}).get("Title", "")).startswith(prefixo)]


def _post(os_dict: dict) -> str:
    payload = {"fields": {
        "Title":   os_dict["numero_os"],
        "DadosOS": json.dumps(os_dict, ensure_ascii=False),
        "Status":  os_dict.get("status", ""),
    }}
    r = requests.post(_base_url(), headers=_headers(), json=payload, timeout=15)
    _checar(r)
    return r.json()["id"]


def _patch(sp_id: str, os_dict: dict) -> None:
    payload = {
        "Title":   os_dict["numero_os"],
        "DadosOS": json.dumps(os_dict, ensure_ascii=False),
        "Status":  os_dict.get("status", ""),
    }
    r = requests.patch(_base_url() + f"/{sp_id}/fields",
                       headers=_headers(), json=payload, timeout=15)
    _checar(r)


def _proximo_numero_livre(ano: int) -> str:
    prefixo = f"OS-{ano}-"
    max_n = 0
    for item in _itens_por_prefixo(prefixo):
        titulo = str(item.get("fields", {}).get("Title", ""))
        try:
            max_n = max(max_n, int(titulo.rsplit("-", 1)[-1]))
        except ValueError:
            pass
    return f"{prefixo}{(max_n + 1):04d}"


def _post_com_numero_unico(os_dict: dict, tentativas: int = 8) -> str:
    """
    Mesmo algoritmo do dados.py da oficina — precisa ser idêntico nos dois
    lados, senão o desempate não converge. Quem tem o menor id do SharePoint
    fica com o número; o outro assume o próximo livre e regrava.
    """
    ano = datetime.now().year
    os_dict["numero_os"] = _proximo_numero_livre(ano)
    sp_id = _post(os_dict)

    for _ in range(tentativas):
        conflitantes = _itens_por_titulo(os_dict["numero_os"])
        if len(conflitantes) <= 1:
            return os_dict["numero_os"]
        vencedor = min(conflitantes, key=lambda i: int(i["id"]))
        if str(vencedor["id"]) == str(sp_id):
            return os_dict["numero_os"]
        os_dict["numero_os"] = _proximo_numero_livre(ano)
        _patch(sp_id, os_dict)

    raise RuntimeError("Não foi possível obter um número de OS único.")


# ─────────────────────────────────────────────
#  RESOLUÇÃO DE CENTRO DE CUSTO
# ─────────────────────────────────────────────

def _normalizar(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt or ""))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return " ".join(txt.upper().split())


def resolver_cod_cc(cc_nome: str) -> tuple[int | None, str]:
    """
    Converte o CCNome do Kanban (texto) no cod_cc inteiro da oficina.

    Estratégia, em ordem — a comparação por NOME vem primeiro de propósito:
    o prefixo numérico do Kanban não é necessariamente o mesmo código do
    centro de custo da oficina (no cadastro de frotas, '46-PATIO TESTON' é
    frente 46, enquanto o CC 46 é 'SERVIÇOS DE PREPARO DE SOLO').

      1. Nome bate com algum valor de CENTROS_CUSTO → usa esse código.
      2. Prefixo numérico existe em CENTROS_CUSTO E o nome confere → usa.
      3. Não resolveu → (None, texto_original). A OS é criada mesmo assim,
         marcada com cc_pendente, e o supervisor corrige na oficina.
    """
    bruto = str(cc_nome or "").strip()
    if not bruto:
        return None, ""

    alvo = _normalizar(bruto)

    # 1 — casamento por nome (com ou sem prefixo numérico)
    sem_prefixo = alvo.split("-", 1)[1].strip() if "-" in alvo and alvo.split("-", 1)[0].strip().isdigit() else alvo
    for cod, nome in CENTROS_CUSTO.items():
        if _normalizar(nome) in (alvo, sem_prefixo):
            return cod, nome

    # 2 — prefixo numérico, só se o nome também for compatível
    primeiro = alvo.split("-", 1)[0].strip()
    if primeiro.isdigit():
        cod = int(primeiro)
        nome_oficial = CENTROS_CUSTO.get(cod)
        if nome_oficial and _normalizar(nome_oficial)[:8] == sem_prefixo[:8]:
            return cod, nome_oficial

    return None, bruto


def resolver_cod_tecnico(nome_inspetor: str) -> int | None:
    """Nome do inspetor → cod_tecnico. Retorna None se não bater exatamente."""
    alvo = _normalizar(nome_inspetor)
    for cod, nome in TECNICOS_OFICINA.items():
        if _normalizar(nome) == alvo:
            return cod
    return None


# ─────────────────────────────────────────────
#  MONTAGEM DA AVALIAÇÃO
# ─────────────────────────────────────────────

def montar_avaliacao(itens_state: dict, observacoes_gerais: str = "",
                     motivo_entrada: str = "", km: int | None = None) -> str:
    """
    Monta o texto do campo Avaliação da OS a partir do checklist.
    Só entram itens marcados como Regular ou Ruim — o resto é ruído para
    quem vai executar o serviço.
    """
    linhas = []
    if motivo_entrada:
        linhas.append(f"Motivo da entrada: {motivo_entrada}")
    if km:
        linhas.append(f"KM na entrada: {km}")
    if linhas:
        linhas.append("")

    problemas = []
    for chave, aval in itens_state.items():
        if chave.endswith("__obs"):
            continue
        if aval in ("✅ Bom", "➖ Não tem", ""):
            continue
        categoria, item = (chave.split("||", 1) + [""])[:2]
        obs = itens_state.get(f"{chave}__obs", "").strip()
        marca = "RUIM" if "Ruim" in aval else "REGULAR"
        texto = f"[{marca}] {item}"
        if obs:
            texto += f" — {obs}"
        problemas.append(texto)

    if problemas:
        linhas.append(f"Itens com problema no checklist ({len(problemas)}):")
        linhas.extend(problemas)
    else:
        linhas.append("Checklist sem itens reprovados.")

    if observacoes_gerais.strip():
        linhas.append("")
        linhas.append(f"Observações do inspetor: {observacoes_gerais.strip()}")

    return "\n".join(linhas)


# ─────────────────────────────────────────────
#  API PÚBLICA
# ─────────────────────────────────────────────

def abrir_os_do_checklist(*, frota_nome: str, cc_nome: str, chassi: str = "",
                          nome_inspetor: str, cod_tecnico: int | None,
                          itens_state: dict, observacoes_gerais: str = "",
                          motivo_entrada: str = "", km: int | None = None,
                          checklist_id: str) -> dict:
    """
    Cria a OS interna na lista AT_Teston_OS e devolve o dict criado.

    Não faz controle de duplicidade por conta própria — quem chama deve
    verificar o campo OSNumero do checklist antes (ver o app_checklist.py).
    """
    cod_cc, cliente_texto = resolver_cod_cc(cc_nome)
    if cod_tecnico is None:
        cod_tecnico = resolver_cod_tecnico(nome_inspetor)

    frota_codigo = str(frota_nome).split("-", 1)[0].strip() or str(frota_nome).strip()

    nova = {
        "numero_os":                   None,
        "data_abertura":               datetime.now().strftime("%Y-%m-%d"),
        "frota":                       frota_codigo,
        "equipamento":                 str(frota_nome).strip(),
        "cod_cc":                      cod_cc,
        "cliente":                     CENTROS_CUSTO.get(cod_cc, cliente_texto) if cod_cc else cliente_texto,
        "cc_pendente":                 cod_cc is None,
        "tipo_os":                     "interna",
        "origem":                      "checklist",
        "checklist_id":                str(checklist_id),
        "avaliacao":                   montar_avaliacao(itens_state, observacoes_gerais,
                                                        motivo_entrada, km),
        "chassi":                      chassi,
        "status":                      "aberta",
        "aberto_por":                  f"checklist:{nome_inspetor}",
        "aberto_em":                   datetime.now().isoformat(),
        "tecnico_designado":           cod_tecnico,
        "fechado_para_aprovacao_por":  None,
        "fechado_para_aprovacao_em":   None,
        "validado_por":                None,
        "validado_em":                 None,
        "observacao_validacao":        "",
        "procedimentos":               [],
    }

    numero = _post_com_numero_unico(nova)
    nova["numero_os"] = numero
    return nova


# ──────────────────────────────────────────────────────────────────────
#  NOTA SOBRE DUPLICAÇÃO DE CONSTANTES
#
#  CENTROS_CUSTO e TECNICOS_OFICINA existem aqui e no dados.py da oficina.
#  Enquanto os dois apps forem repositórios separados, essa duplicação é o
#  preço da escrita direta. Se divergirem, a OS nasce com cliente errado.
#
#  Quando quiser eliminar isso, o caminho é mover as duas tabelas para uma
#  lista do SharePoint (ex: TabelasOficina) e os dois apps lerem de lá —
#  mesma solução que você já pensou para as frotas com o HistoricoFrenteFrotas.
# ──────────────────────────────────────────────────────────────────────
