"""
00b_fusao_festivais.py
======================
Cria o dataset único de pontuação de festivais a partir de duas Atas oficiais
BRDE/FSA (fontes primárias) complementadas pela planilha de pesquisa bibliográfica.

Fontes (ordem de prioridade):
  1. (PRIMÁRIA) dados/festivais_por_obra_ata_fsa2024.csv
       Ata Final BRDE/FSA — Desempenho Artístico 1ª Edição 2024
       (ANCINE, 26/09/2024). Gerado por 00_extrair_ata_fsa2024.py. 232 obras.
  2. (PRIMÁRIA) tabelas_apoio/Ata-Habilit-Final-Desempenho-Artistico-2024-2o-edicao_08.08.2025-publicacao.pdf
       Ata Final BRDE/FSA — Desempenho Artístico 2ª Edição 2024
       (ANCINE, 08/08/2025). 22 obras. Extraído diretamente aqui.
  3. (COMPLEMENTO) tabelas_apoio/Festivais_por_obra_pre_expansao.xlsx
       Pesquisa bibliográfica / Claude, mar/2026.
       Usada APENAS para obras não cobertas por nenhuma das Atas.

Saída:
  resultados/festivais_consolidado.csv
       Fonte única consumida pelos demais scripts do pipeline.

Uso:
    python scripts/00b_fusao_festivais.py
"""

import csv
import math
import os
import re
import unicodedata

import pandas as pd

ROOT      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ATA_CSV   = os.path.join(ROOT, "dados",         "festivais_por_obra_ata_fsa2024.csv")
ATA2_PDF  = os.path.join(ROOT, "tabelas_apoio", "Ata-Habilit-Final-Desempenho-Artistico-2024-2o-edicao_08.08.2025-publicacao.pdf")
XLSX_PRE  = os.path.join(ROOT, "tabelas_apoio", "Festivais_por_obra_pre_expansao.xlsx")
OUT_CSV   = os.path.join(ROOT, "resultados",    "festivais_consolidado.csv")

# Atas históricas (PRODAV 07 — Desempenho Artístico)
ATA_2014  = os.path.join(ROOT, "tabelas_apoio", "PRODAV-072014_resultado-final.pdf")
ATA_2015  = os.path.join(ROOT, "tabelas_apoio", "PRODAV-072015_resultadofinal.pdf")
ATA_2017  = os.path.join(ROOT, "tabelas_apoio", "Resultado-Final-PRODAV-07-2017.pdf")
ATA_2018  = os.path.join(ROOT, "tabelas_apoio", "Ata-Resultado-Final_DesempenhoArtístico_2018.pdf")

FONTE_ATA  = "Ata BRDE/FSA Desempenho Artístico 2024 (ANCINE, 26/09/2024)"
FONTE_ATA2 = "Ata BRDE/FSA Desempenho Artístico 2ª Edição 2024 (ANCINE, 08/08/2025)"
FONTE_2014 = "Ata BRDE/FSA PRODAV 07/2014 — PAQ (resultado final)"
FONTE_2015 = "Ata BRDE/FSA PRODAV 07/2015 — Desempenho Artístico (resultado final)"
FONTE_2017 = "Ata BRDE/FSA PRODAV 07/2017 — Desempenho Artístico (resultado final)"
FONTE_2018 = "Ata BRDE/FSA Desempenho Artístico 2018 (resultado final)"
FONTE_XLSX = "Festivais_por_obra_pre_expansao.xlsx (pesquisa bibliográfica / Claude, mar/2026)"

# Colunas individuais de festival presentes na planilha xlsx.
# São incluídas na saída para uso pelo pipeline (ex.: FEST_COLS no script 01).
# Para obras vindas da Ata (que não tem essa granularidade), ficam como 0.
FEST_COLS = [
    "Oscar", "Cannes", "Berlim", "Veneza", "Sundance", "Locarno", "TIFF",
    "San Seb.", "Rotterdam", "Annecy", "Outros Intl", "Havana", "NYFF",
    "BFI London", "BAFTA", "Globo de Ouro", "Brasília", "Gramado",
    "Fest.Rio", "Mostra SP",
]

# Mapeamento: nome bruto da subcoluna no xlsx → nome limpo no CSV de saída
_XLSX_RAW_TO_CLEAN = {
    "Oscar\n(Esp.)":      "Oscar",
    "Cannes\n(Esp.)":     "Cannes",
    "Berlim\n(Esp.)":     "Berlim",
    "Veneza\n(Esp.)":     "Veneza",
    "Sundance\n(AA)":     "Sundance",
    "Locarno\n(AA)":      "Locarno",
    "TIFF\n(AA)":         "TIFF",
    "San Seb.\n(AA)":     "San Seb.",
    "Rotterdam\n(AA)":    "Rotterdam",
    "Annecy\n(AA)":       "Annecy",
    "Outros\nIntl (A)":   "Outros Intl",
    "Havana\n(A)":        "Havana",
    "NYFF\n(A)":          "NYFF",
    "BFI Lond.\n(A)":     "BFI London",
    "BAFTA":              "BAFTA",
    "Globo de\nOuro":     "Globo de Ouro",
    "Brasília":           "Brasília",
    "Gramado":            "Gramado",
    "Fest.Rio":           "Fest.Rio",
    "Mostra SP":          "Mostra SP",
}

FIELDNAMES = [
    "cpb", "titulo", "titulo_norm", "pontuacao_total",
    "n_festivais", "lista_festivais", "fonte",
] + FEST_COLS


# ── helpers ──────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def _zero_fest() -> dict:
    return {c: 0 for c in FEST_COLS}


def _to_float(v, default=0.0) -> float:
    try:
        return float(str(v).replace(",", "."))
    except Exception:
        return default


# ── 1. Carregar Ata (fonte primária) ─────────────────────────────────────────

def load_ata() -> dict:
    """
    Retorna dict titulo_norm → registro completo.
    Filtra APENAS as entradas originárias da Ata PDF (exclui o complemento
    xlsx que o script 00 já incorporou — esse complemento é substituído aqui
    pelo Festivais_por_obra_pre_expansao.xlsx).
    """
    result = {}
    df = pd.read_csv(ATA_CSV, encoding="utf-8-sig", dtype=str).fillna("")
    for _, r in df.iterrows():
        tn = r.get("titulo_norm", "").strip()
        fonte = r.get("fonte", "")
        if not tn:
            continue
        # Aceita apenas entradas que vieram da Ata (não do complemento xlsx antigo)
        if "Ata BRDE" not in fonte:
            continue
        result[tn] = {
            "cpb":             r.get("cpb", "").strip(),
            "titulo":          r.get("titulo", "").strip(),
            "titulo_norm":     tn,
            "pontuacao_total": _to_float(r.get("pontuacao_total", "0")),
            "n_festivais":     int(_to_float(r.get("n_festivais", "0"))),
            "lista_festivais": r.get("lista_festivais", "").strip(),
            "fonte":           fonte.strip(),
            **_zero_fest(),
        }
    return result


# ── 2. Extrair Ata 2ª Edição direto do PDF ───────────────────────────────────

def _pdf_lines(pdf_path: str, page_start: int = 0) -> list[str]:
    """Extrai linhas de texto de um PDF usando posição das palavras."""
    try:
        import fitz
    except ImportError:
        raise RuntimeError("PyMuPDF não encontrado. Use: pip install pymupdf")
    doc = fitz.open(pdf_path)
    lines = []
    for i in range(page_start, len(doc)):
        words = doc[i].get_text("words")
        rows: dict = {}
        for w in words:
            y_key = round(w[1] / 4) * 4
            rows.setdefault(y_key, []).append((w[0], w[4]))
        for yk in sorted(rows):
            tokens = sorted(rows[yk], key=lambda t: t[0])
            line = " ".join(t[1] for t in tokens).strip()
            if re.match(r"^\d+$", line):
                continue
            if line in ("Informação Pública", "Informação  Pública"):
                continue
            if not line:
                continue
            lines.append(line)
    return lines


def _is_cpb(s: str) -> bool:
    return bool(re.match(r"^B\d{13}$", s.strip()))


def _parse_score(line: str) -> tuple[str, float | None]:
    parts = line.rsplit(" ", 1)
    if len(parts) == 2:
        try:
            return parts[0].strip(), float(parts[1].replace(",", "."))
        except ValueError:
            pass
    return line.strip(), None


def _parse_obra_lines(lines: list[str]) -> list[dict]:
    """Parseia linhas extraídas do PDF em lista de obras."""
    obras = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if i + 1 < len(lines) and _is_cpb(lines[i + 1]):
            titulo, total = _parse_score(line)
            cpb = lines[i + 1].strip()
            festivais = []
            j = i + 2
            while j < len(lines):
                if j + 1 < len(lines) and _is_cpb(lines[j + 1]):
                    break
                if _is_cpb(lines[j]):
                    j += 1
                    continue
                ft, fp = _parse_score(lines[j])
                if ft and fp is not None:
                    festivais.append((ft, fp))
                j += 1
            obras.append({"titulo": titulo, "cpb": cpb,
                          "total": total, "festivais": festivais})
            i = j
        else:
            i += 1
    return obras


# ── helpers PDF genérico ─────────────────────────────────────────────────────

_UFS = {'AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG',
        'PA','PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'}

_SKIP_OLD = re.compile(
    r'^(ATA DA|CHAMADA|SUPORTE|BRDE/FSA|RESULTADO|As |Às |'
    r'Na fase|Os demais|A Comiss|Sendo assim|A classif|'
    r'Após |Encerrado|Cada uma|O valor|EMPRESAS INABILITADAS|'
    r'Proponente|PRODUTORA|PROJETO|Pontuação|UF$|Projeto$|'
    r'__|Flávia|Aimê|Elizabeth|Amanda|Barbara|Daniel|Tatiana|Elaine|'
    r'Paula|André|Anna|ANCINE|processo|Portaria|nomeados)', re.IGNORECASE)

_NOISE_TITLE = re.compile(
    r'(\(Brasil\)|\(EUA\)|\(Arg|Festival |Mostra |'
    r'\bLTDA\b|\bS\.A\b|\bPRODUÇ|\bSERVIÇOS|\bFILMES\b)',
    re.IGNORECASE)


def _parse_old_pdf(pdf_path: str, fonte: str) -> dict:
    """
    Parseia Atas BRDE/FSA anteriores a 2024 (2014-2018).
    Estratégia: ancora no código UF da produtora e coleta
    título + pontuação total nos campos seguintes.
    """
    import fitz
    doc = fitz.open(pdf_path)

    all_lines: list[str] = []
    for page in doc:
        page_lines = [l.strip() for l in page.get_text("text").split("\n") if l.strip()]
        # Pular número de página (primeira linha não-vazia de cada página)
        if page_lines and re.match(r"^\d{1,2}$", page_lines[0]):
            page_lines = page_lines[1:]
        all_lines.extend(page_lines)

    lines = [l for l in all_lines if not _SKIP_OLD.match(l)]

    def _flt(s: str):
        try:
            v = float(s.replace(",", "."))
            return v if v >= 0 else None
        except Exception:
            return None

    result: dict = {}

    for i, line in enumerate(lines):
        # Âncora: linha que começa com código UF (sozinho ou + início do título)
        m = re.match(r"^([A-Z]{2})(?:\s+(.+))?$", line)
        if not m or m.group(1) not in _UFS:
            continue

        uf_rest = (m.group(2) or "").strip()

        # Coletar título e score avançando a partir da posição da UF
        title_parts: list[str] = [uf_rest] if uf_rest else []
        score = None
        j = i + 1
        while j < len(lines) and j < i + 8:
            nxt = lines[j]
            v = _flt(nxt)
            if v is not None and title_parts:
                score = v
                break
            elif v is None:
                title_parts.append(nxt)
            j += 1

        if not title_parts or score is None or score == 0.0:
            continue

        titulo = " ".join(title_parts).strip()

        # Filtrar títulos que são na verdade nomes de festival ou de empresa
        if _NOISE_TITLE.search(titulo):
            continue
        if len(titulo) > 70:   # fragmento longo → ruído
            continue

        tn = _norm(titulo)
        if tn and tn not in result:
            result[tn] = {
                "cpb":             "",
                "titulo":          titulo,
                "titulo_norm":     tn,
                "pontuacao_total": score,
                "n_festivais":     0,
                "lista_festivais": "",
                "fonte":           fonte,
                **_zero_fest(),
            }

    return result


def load_ata2() -> dict:
    """
    Extrai obras da Ata 2ª Edição diretamente do PDF.
    O Anexo I (scoring detalhado) começa na página 5 (índice 4).
    """
    if not os.path.exists(ATA2_PDF):
        print("  [AVISO] Ata 2ª Edição não encontrada, ignorando.")
        return {}

    # Anexo I começa na pág 5 (índice 4); páginas anteriores são texto intro
    lines = _pdf_lines(ATA2_PDF, page_start=4)
    obras = _parse_obra_lines(lines)

    result = {}
    for o in obras:
        tn = _norm(o["titulo"])
        lista = " | ".join(f[0] for f in o["festivais"])
        result[tn] = {
            "cpb":             o["cpb"],
            "titulo":          o["titulo"],
            "titulo_norm":     tn,
            "pontuacao_total": o["total"] or 0.0,
            "n_festivais":     len(o["festivais"]),
            "lista_festivais": lista,
            "fonte":           FONTE_ATA2,
            **_zero_fest(),
        }
    return result


# ── 3. Carregar xlsx (complemento) ───────────────────────────────────────────

def load_xlsx() -> dict:
    """
    Lê Festivais_por_obra_pre_expansao.xlsx.
    Retorna dict titulo_norm → registro, incluindo pontuação por festival.
    """
    df = pd.read_excel(XLSX_PRE, header=3)

    # Linha 0 (após header=3) contém os sub-cabeçalhos das colunas de festival
    sub = df.iloc[0]
    col_rename = {}
    for i, col in enumerate(df.columns):
        nm = str(sub.iloc[i]).strip()
        if nm not in ("nan", "NaN", ""):
            col_rename[col] = nm
    df = df.rename(columns=col_rename).iloc[1:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]

    titulo_col = next((c for c in df.columns if "tulo" in c.lower()), None)
    total_col  = next((c for c in df.columns if "TOTAL" in c.upper()), None)

    # Mapear colunas brutas do xlsx para nomes limpos do CSV de saída
    raw_to_clean = {}
    for col in df.columns:
        clean = _XLSX_RAW_TO_CLEAN.get(col)
        if clean:
            raw_to_clean[col] = clean

    result = {}
    for _, row in df.iterrows():
        t = row.get(titulo_col, "")
        if not isinstance(t, str) or not t.strip():
            continue
        tn = _norm(t)

        fest = _zero_fest()
        for raw_col, clean_col in raw_to_clean.items():
            v = pd.to_numeric(row.get(raw_col, 0), errors="coerce")
            fest[clean_col] = int(v) if (v and not math.isnan(float(v))) else 0

        if total_col:
            tv = pd.to_numeric(row.get(total_col, 0), errors="coerce")
            total = float(tv) if (tv and not math.isnan(float(tv))) else 0.0
        else:
            total = float(sum(fest.values()))

        # Ignorar linhas de legenda (total e todos os scores = 0)
        if total == 0.0 and sum(fest.values()) == 0:
            continue

        result[tn] = {
            "cpb":             "",
            "titulo":          t.strip(),
            "titulo_norm":     tn,
            "pontuacao_total": total,
            "n_festivais":     0,
            "lista_festivais": "",
            "fonte":           FONTE_XLSX,
            **fest,
        }
    return result


# ── 4. Fundir e salvar ────────────────────────────────────────────────────────

def main():
    print("=== 00b_fusao_festivais.py ===")

    print("  Carregando Ata 1ª Edição 2024 (primária)...")
    ata1 = load_ata()
    print(f"  Obras na Ata 1ª Ed.: {len(ata1)}")

    print("  Extraindo Ata 2ª Edição 2024 do PDF (primária)...")
    ata2 = load_ata2()
    print(f"  Obras na Ata 2ª Ed.: {len(ata2)}")
    overlap = set(ata1) & set(ata2)
    if overlap:
        print(f"  Sobreposição entre edições (2ª Ed. prevalece): {len(overlap)}")

    # Atas históricas (ordem cronológica — mais recente sobrescreve mais antiga)
    old_atas: dict = {}
    for pdf_path, fonte, ano in [
        (ATA_2014, FONTE_2014, "2014"),
        (ATA_2015, FONTE_2015, "2015"),
        (ATA_2017, FONTE_2017, "2017"),
        (ATA_2018, FONTE_2018, "2018"),
    ]:
        if not os.path.exists(pdf_path):
            print(f"  [AVISO] Ata {ano} não encontrada: {pdf_path}")
            continue
        print(f"  Extraindo Ata PRODAV {ano} do PDF...")
        ata_hist = _parse_old_pdf(pdf_path, fonte)
        print(f"  Obras na Ata {ano}: {len(ata_hist)}")
        old_atas.update(ata_hist)   # mais recente sobrescreve

    print("  Carregando Festivais_por_obra_pre_expansao.xlsx (complemento)...")
    xlsx = load_xlsx()
    print(f"  Obras no xlsx: {len(xlsx)}")

    # Merge (prioridade decrescente):
    #   Ata 2024 2ª Ed. > Ata 2024 1ª Ed. > Atas históricas (2018>2017>2015>2014) > xlsx
    merged = {}
    # xlsx como base (menor prioridade)
    for tn, rec in xlsx.items():
        merged[tn] = rec
    # Atas históricas sobrescrevem xlsx
    adicionados_hist = sum(1 for tn in old_atas if tn not in merged)
    merged.update(old_atas)
    # Atas 2024 têm prioridade máxima
    merged.update(ata1)
    merged.update(ata2)   # 2ª Ed. prevalece sobre 1ª

    adicionados_xlsx = sum(1 for tn in xlsx if tn not in {**old_atas, **ata1, **ata2})
    print(f"  Obras adicionadas das Atas históricas (novas): {adicionados_hist}")
    print(f"  Obras adicionadas do xlsx (não em nenhuma Ata): {adicionados_xlsx}")
    print(f"  Total consolidado: {len(merged)} obras")

    os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)

    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDNAMES)
        writer.writeheader()
        for rec in merged.values():
            writer.writerow({k: rec.get(k, 0) for k in FIELDNAMES})

    print(f"  Salvo: {OUT_CSV}")
    print("  Concluído.")


if __name__ == "__main__":
    main()
