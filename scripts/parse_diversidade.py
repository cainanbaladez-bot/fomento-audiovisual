"""
parse_diversidade.py — Lê raw/Posição_20260324.xlsx e computa métricas de
diversidade (raça/gênero) para editais COM e SEM Política Afirmativa.

Estrutura do xlsx (cabeçalho multi-nível, 3 linhas):
  Linha 0: Função (DIRETOR | ROTEIRISTA | PRODUTOR)
  Linha 1: Subcategoria (Branca, Parda, Preta, Indígena, Homem cis, Mulher cis etc.)
  Linha 2: Status (Contratados/em contratação | Não contratados)
  Dados a partir da linha 4 (linha 3 é vazia)

Colunas (para cada função, 22 colunas):
  Raça: Branca(c,n), Amarela(c,n), Parda(c,n), Preta(c,n), Indígena(c,n)
  Gênero: Homem cis(c,n), Mulher cis(c,n), Outro(c,n), Não declarar(c,n),
          Homem trans(c,n), Mulher trans/travesti(c,n)

Fonte: BRDE/FSA — posição consolidada dos editais seletivos.

Classificação PA vs Controle
-----------------------------
Grupo PA (7 editais): editais com critérios formais de diversidade na seleção,
confirmados por leitura dos regulamentos publicados pelo BRDE/FSA.

Grupo Controle (8 editais): editais seletivos de produção Cinema/TV-VOD do
mesmo período (2022-2024), sem critérios de diversidade, mas com processo
seletivo competitivo comparável (comissão de seleção, nota de mérito).

Editais excluídos da comparação PA vs Controle:
  - Arranjos Regionais (aprovação quase-automática, sem seleção competitiva)
  - Fluxo Contínuo / Suporte Automático (sem comissão de seleção)
  - Comercialização / Complementação (mecanismo de distribuição, não produção)
  - Desempenho Comercial / Artístico (automático por métricas, sem seleção)
  - Coprodução Internacional (seleção bilateral, critérios distintos)
  - SAV/2018 e RioFilme/2025 (não confirmada aplicação formal de PA)
  - PRODAV / PRODECINE (mecanismos diferentes, TV pública)
"""

import os
import pandas as pd

_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_XLSX = os.path.join(_BASE, "raw", "Posição_20260324.xlsx")

# ── Editais COM Política Afirmativa confirmada (7 editais) ──────────────────
# Critério: regulamento publicado pelo BRDE/FSA prevê critérios explícitos
# de raça e/ou gênero na avaliação e seleção dos projetos.
_PA_EXATOS = [
    "BRDE/FSA - CINEMA - NOVOS REALIZADORES 2022",
    "BRDE/FSA - PRODUÇÃO SELETIVO CINEMA 2024",
    "BRDE/FSA - PRODUÇÃO SELETIVO TV VOD 2024",
    "TV-VOD NOVOS REALIZADORES 2022",
    "Produção para cinema 2018 - Modalidade A",
    "Produção para cinema 2018 - Modalidade B",
]
_PA_PARCIAL = ["RUTH DE SOUZA"]  # match parcial (nome varia entre edições)

# ── Grupo de Controle (8 editais) ──────────────────────────────────────────
# Critério: editais seletivos de produção Cinema/TV-VOD, mesmo período
# (2022-2024), processo competitivo comparável, SEM critérios de diversidade.
_CONTROLE_PARCIAL = [
    "BRDE/FSA CINEMA PRODUÇÃO 2023",       # Nacional + Regional
    "BRDE/FSA PRODUÇÃO CINEMA - 2022",      # Nacional + Regional
    "PRODUÇÃO TV-VOD 2022",                 # Nacional + Regional
    "PRODUÇÃO TV-VOD 2023",                 # Regional
    "TV-VOD – 2023",                        # Nacional (traço longo no nome)
]


def _is_pa(nome: str) -> bool:
    n = str(nome).strip()
    if n in _PA_EXATOS:
        return True
    nu = n.upper()
    for kw in _PA_PARCIAL:
        if kw in nu:
            return True
    return False


def _is_controle(nome: str) -> bool:
    nu = str(nome).upper().strip()
    for kw in _CONTROLE_PARCIAL:
        if kw.upper() in nu:
            return True
    return False


def _is_arranjo(nome: str) -> bool:
    return "ARRANJOS REGIONAIS" in str(nome).upper()


def load(xlsx_path: str = _XLSX):
    """Retorna DataFrame limpo: 1 linha por edital, colunas numéricas."""
    df = pd.read_excel(xlsx_path, header=None)
    data = df.iloc[4:].copy()
    data = data[data.iloc[:, 0].notna()].reset_index(drop=True)
    for c in range(1, min(68, data.shape[1])):
        data.iloc[:, c] = pd.to_numeric(data.iloc[:, c], errors="coerce").fillna(0)
    data["edital"] = data.iloc[:, 0].astype(str)
    data["tem_pa"] = data["edital"].apply(_is_pa)
    data["controle"] = data["edital"].apply(_is_controle)
    data["arranjo"] = data["edital"].apply(_is_arranjo)
    return data


def _metrics(sub):
    """Calcula métricas de raça e gênero para um subconjunto de editais (DIRETOR)."""
    # Raça — colunas 2-11
    bra_c, bra_n = sub.iloc[:, 2].sum(), sub.iloc[:, 3].sum()
    ama_c, ama_n = sub.iloc[:, 4].sum(), sub.iloc[:, 5].sum()
    par_c, par_n = sub.iloc[:, 6].sum(), sub.iloc[:, 7].sum()
    pre_c, pre_n = sub.iloc[:, 8].sum(), sub.iloc[:, 9].sum()
    ind_c, ind_n = sub.iloc[:, 10].sum(), sub.iloc[:, 11].sum()

    neg_c = par_c + pre_c + ind_c
    neg_n = par_n + pre_n + ind_n
    total_raca = bra_c + bra_n + ama_c + ama_n + neg_c + neg_n
    total_sel_raca = bra_c + ama_c + neg_c

    # Gênero — colunas 12-23
    hom_c, hom_n = sub.iloc[:, 12].sum(), sub.iloc[:, 13].sum()
    mul_c, mul_n = sub.iloc[:, 14].sum(), sub.iloc[:, 15].sum()
    out_c = sub.iloc[:, 16].sum() + sub.iloc[:, 18].sum() + sub.iloc[:, 20].sum() + sub.iloc[:, 22].sum()
    out_n = sub.iloc[:, 17].sum() + sub.iloc[:, 19].sum() + sub.iloc[:, 21].sum() + sub.iloc[:, 23].sum()
    mtrans_c, mtrans_n = sub.iloc[:, 22].sum(), sub.iloc[:, 23].sum()

    total_genero = hom_c + hom_n + mul_c + mul_n + out_c + out_n
    total_sel_genero = hom_c + mul_c + out_c

    def _pct(num, den):
        return round(num / den * 100, 1) if den > 0 else 0.0

    return {
        "n_editais": len(sub),
        "n_inscricoes": int(sub.iloc[:, 1].sum()),
        # Raça
        "pct_branco_inscritos": _pct(bra_c + bra_n, total_raca),
        "pct_branco_selecionados": _pct(bra_c, total_sel_raca),
        "pct_negro_inscritos": _pct(neg_c + neg_n, total_raca),
        "pct_negro_selecionados": _pct(neg_c, total_sel_raca),
        "taxa_selecao_negro": _pct(neg_c, neg_c + neg_n),
        "taxa_selecao_branco": _pct(bra_c, bra_c + bra_n),
        "n_negro_contratados": int(neg_c),
        "n_negro_total": int(neg_c + neg_n),
        "n_branco_contratados": int(bra_c),
        "n_branco_total": int(bra_c + bra_n),
        # Gênero
        "pct_mulher_inscritas": _pct(mul_c + mul_n, total_genero),
        "pct_mulher_selecionadas": _pct(mul_c, total_sel_genero),
        "taxa_selecao_mulher": _pct(mul_c, mul_c + mul_n),
        "taxa_selecao_homem": _pct(hom_c, hom_c + hom_n),
        "n_mulher_contratadas": int(mul_c),
        "n_mulher_total": int(mul_c + mul_n),
        "n_homem_contratados": int(hom_c),
        "n_homem_total": int(hom_c + hom_n),
        # Trans
        "pct_mtrans_selecao": _pct(mtrans_c, mtrans_c + mtrans_n),
        # Produção executiva (PRODUTOR) — colunas 46-67, gênero nas 56-67
        "pct_mulher_prod_inscritas": _pct(
            sub.iloc[:, 58].sum() + sub.iloc[:, 59].sum(),
            sum(sub.iloc[:, c].sum() for c in range(56, 68))
        ),
        # Direção global (% mulher entre inscritas)
        "pct_mulher_dir_global": _pct(mul_c + mul_n, total_genero),
    }


def compute():
    """Retorna dict com métricas COM PA, Controle, e TODOS (excl. Arranjos)."""
    data = load()
    sem_arranjo = data[~data["arranjo"]]
    pa = sem_arranjo[sem_arranjo["tem_pa"]]
    ctrl = sem_arranjo[sem_arranjo["controle"]]
    return {
        "com_pa": _metrics(pa),
        "sem_pa": _metrics(ctrl),          # grupo de controle
        "todos": _metrics(sem_arranjo),
        "editais_pa": list(pa["edital"].values),
        "editais_controle": list(ctrl["edital"].values),
        # Alias para compatibilidade
        "editais_sem_pa": list(ctrl["edital"].values),
    }


def spotlight():
    """Retorna métricas por edital individual (para tabela spotlight)."""
    data = load()
    pa = data[data["tem_pa"] & ~data["arranjo"]]
    rows = []
    for _, r in pa.iterrows():
        neg_c = r.iloc[6] + r.iloc[8] + r.iloc[10]
        neg_n = r.iloc[7] + r.iloc[9] + r.iloc[11]
        bra_c, bra_n = r.iloc[2], r.iloc[3]
        ama_c, ama_n = r.iloc[4], r.iloc[5]
        total_r = bra_c + bra_n + ama_c + ama_n + neg_c + neg_n
        total_sel_r = bra_c + ama_c + neg_c

        mul_c, mul_n = r.iloc[14], r.iloc[15]
        hom_c = r.iloc[12]
        out_c = r.iloc[16] + r.iloc[18] + r.iloc[20] + r.iloc[22]
        total_sel_g = hom_c + mul_c + out_c

        def _p(n, d):
            return round(n / d * 100, 1) if d > 0 else 0.0

        rows.append({
            "edital": r["edital"],
            "n_inscricoes": int(r.iloc[1]),
            "pct_negro_insc": _p(neg_c + neg_n, total_r),
            "pct_negro_sel": _p(neg_c, total_sel_r),
            "delta_negro": round(_p(neg_c, total_sel_r) - _p(neg_c + neg_n, total_r), 1),
            "pct_mulher_insc": _p(mul_c + mul_n, sum(r.iloc[12:24])),
            "pct_mulher_sel": _p(mul_c, total_sel_g),
            "delta_mulher": round(_p(mul_c, total_sel_g) - _p(mul_c + mul_n, sum(r.iloc[12:24])), 1),
        })
    return rows


if __name__ == "__main__":
    m = compute()
    print("=== COM PA ===")
    for k, v in m["com_pa"].items():
        print(f"  {k}: {v}")
    print(f"\nEditais com PA ({len(m['editais_pa'])}):")
    for e in m["editais_pa"]:
        print(f"  {e}")
    print(f"\n=== CONTROLE ===")
    for k, v in m["sem_pa"].items():
        print(f"  {k}: {v}")
    print(f"\nEditais controle ({len(m['editais_controle'])}):")
    for e in m["editais_controle"]:
        print(f"  {e}")
    print(f"\n=== TODOS (excl. Arranjos) ===")
    for k, v in m["todos"].items():
        print(f"  {k}: {v}")
