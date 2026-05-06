"""
00_extrair_ata_fsa2024.py
=========================
Extrai tabela de pontuação de festivais da Ata Final BRDE/FSA - Desempenho
Artístico 2024 (publicada em 26/09/2024) e salva dois datasets:

  1. resultados/dataset/base_festivais_obras_ata.csv
     → Granular: uma linha por (obra × festival), com fonte
  2. dados/festivais_por_obra_ata_fsa2024.csv
     → Resumo: uma linha por obra (CPB, título, pontuação total, lista festivais, fonte)

Fonte primária: tabela_apoio/Ata-Habilit-Final-Desempenho-Artistico-2024_26.09.2024-publicacao.pdf
Complementada por: raw/Festivais_por_obra.xlsx (pesquisa Claude – obras não cobertas pela Ata)

Uso:
    python scripts/00_extrair_ata_fsa2024.py
"""

import os
import re
import sys
import unicodedata
import csv

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF_PATH = os.path.join(ROOT, "tabelas_apoio",
                        "Ata-Habilit-Final-Desempenho-Artistico-2024_26.09.2024-publicacao.pdf")
OUT_GRANULAR = os.path.join(ROOT, "resultados", "dataset",
                            "base_festivais_obras_ata.csv")
OUT_RESUMO   = os.path.join(ROOT, "dados",
                            "festivais_por_obra_ata_fsa2024.csv")

FONTE_LABEL  = "Ata BRDE/FSA Desempenho Artístico 2024 (ANCINE, 26/09/2024)"

# ─── helpers ────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()


def _parse_num(s: str):
    """Parse Brazilian decimal number (comma separator)."""
    s = s.strip().replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _is_cpb(s: str) -> bool:
    return bool(re.match(r"^B\d{13}$", s.strip()))


def _is_number(s: str) -> bool:
    return _parse_num(s) is not None


# ─── 1. Extrair texto do PDF via posição das palavras ───────────────────────

def extract_pdf_text() -> list[str]:
    """
    Extrai texto da seção de pontuação usando posições de palavras.
    O PDF tem layout de duas colunas (festival | pontuação), então
    simplesmente ler linhas de texto perde os scores. Usamos word-boxes
    para reconstruir cada linha agrupando palavras com mesmo y-offset.
    Retorna lista de linhas no formato "Festival Name\tScore" ou só "título".
    """
    try:
        import fitz
    except ImportError:
        sys.exit("[ERRO] PyMuPDF não encontrado. Use: pip install pymupdf")

    doc = fitz.open(PDF_PATH)
    result_lines: list[str] = []

    # Pontuação começa na página 22 (índice 21) e termina na 81 (índice 80)
    for page_idx in range(21, 81):
        page = doc[page_idx]
        words = page.get_text("words")
        # words: (x0, y0, x1, y1, text, block_no, line_no, word_no)

        # Agrupar palavras por linha (mesmo y0, tolerância 4 pts)
        rows: dict[int, list[tuple[float, str]]] = {}
        for w in words:
            x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
            # Quantizar y0 para agrupar palavras na mesma linha física
            y_key = round(y0 / 4) * 4
            if y_key not in rows:
                rows[y_key] = []
            rows[y_key].append((x0, text))

        # Ordenar por y, depois por x dentro de cada linha
        for y_key in sorted(rows.keys()):
            tokens = sorted(rows[y_key], key=lambda t: t[0])
            line = " ".join(t[1] for t in tokens).strip()

            # Filtrar cabeçalhos de página
            if re.match(r"^\d+$", line):  # número de página
                continue
            if line in ("Informação Pública", "Informação  Pública"):
                continue
            if not line:
                continue

            result_lines.append(line)

    return result_lines


# ─── 2. Parsear tabela de pontuação ─────────────────────────────────────────

def _split_text_score(line: str) -> tuple[str, float | None]:
    """
    Cada linha da extração por posição tem formato "Texto ... Numero".
    Retorna (texto_sem_score, score) ou (line, None).
    """
    parts = line.rsplit(" ", 1)
    if len(parts) == 2:
        score = _parse_num(parts[1])
        if score is not None:
            return parts[0].strip(), score
    return line.strip(), None


def parse_scoring_table(lines: list[str]) -> list[dict]:
    """
    Formato de entrada (após extração por posição):
      "A FEBRE 114,5"            ← título + total (próxima linha = CPB)
      "B1900486200000"            ← CPB
      "Academia Brasileira... 8"  ← festival + score
      ...
      "DIAMANTINO 88"             ← próximo título
      "B1800445100000"
      ...
    """
    # Localizar início da tabela ("TÍTULO - CPB - FESTIVAL PONTUAÇÃO")
    start = 0
    for idx, ln in enumerate(lines):
        if "TÍTULO" in ln and "CPB" in ln and "FESTIVAL" in ln:
            start = idx + 1
            break

    # Localizar fim
    end = len(lines)
    for idx in range(start, len(lines)):
        if "Divisão dos recursos financeiros" in lines[idx]:
            end = idx
            break

    lines = lines[start:end]

    films: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]

        # Padrão de início de obra: linha com score no final, SEGUIDA de CPB
        if (i + 1 < len(lines)
                and _is_cpb(lines[i + 1])):

            text, total = _split_text_score(line)
            cpb = lines[i + 1].strip()
            titulo = text  # pode conter o total junto ao texto — já separado

            # Coletar pares festival/score até próximo CPB ou fim
            festivais: list[tuple[str, float]] = []
            j = i + 2
            while j < len(lines):
                # Próxima obra: linha seguida de CPB
                if j + 1 < len(lines) and _is_cpb(lines[j + 1]):
                    break
                # CPB solitário (não deve acontecer, mas resguardar)
                if _is_cpb(lines[j]):
                    j += 1
                    continue
                # Fim da seção
                if "Divisão dos recursos financeiros" in lines[j]:
                    break

                fest_text, fest_pts = _split_text_score(lines[j])
                if fest_text and fest_pts is not None:
                    festivais.append((fest_text, fest_pts))
                j += 1

            films.append({
                "titulo":           titulo,
                "cpb":              cpb,
                "pontuacao_total":  total,
                "festivais":        festivais,
            })
            i = j
        else:
            i += 1

    return films


# ─── 3. Complementar com xlsx (Claude research) ─────────────────────────────

def load_xlsx_complement(norms_ja_cobertos: set[str]) -> list[dict]:
    """
    Lê Festivais_por_obra.xlsx via subprocess (sistema Python tem pandas ok).
    Retorna obras NÃO cobertas pela Ata, com seus scores originais.
    """
    xlsx_path = os.path.join(ROOT, "raw", "Festivais_por_obra.xlsx")
    extras: list[dict] = []

    # Python do sistema (3.10, tem pandas compatível)
    py_sys = r"C:\Users\INTEL\AppData\Local\Microsoft\WindowsApps\python.exe"
    if not os.path.exists(py_sys):
        py_sys = "python"

    import subprocess, json
    script = rf"""
import pandas as pd, json, sys
sys.stdout.reconfigure(encoding='utf-8')
df = pd.read_excel(r'{xlsx_path}', header=3)
sub = df.iloc[0]
remap = {{}}
for ci, col in enumerate(df.columns):
    nm = str(sub.iloc[ci]).strip()
    if nm not in ('nan', 'NaN', ''):
        remap[col] = nm
df = df.rename(columns=remap).iloc[1:].reset_index(drop=True)
df.columns = [str(c).strip() for c in df.columns]
titulo_col = next((c for c in df.columns if 'tulo' in c.lower()), None)
total_col  = next((c for c in df.columns if 'TOTAL' in c.upper()), None)
if not titulo_col or not total_col:
    print('[]')
    sys.exit()
out = []
for _, row in df.iterrows():
    t = row.get(titulo_col, '')
    if not isinstance(t, str) or not t.strip():
        continue
    try:
        total_v = float(str(row.get(total_col, 0) or 0).replace(',', '.'))
    except:
        total_v = 0.0
    out.append([t.strip(), total_v])
print(json.dumps(out, ensure_ascii=False))
"""
    try:
        result = subprocess.run(
            [py_sys, "-c", script],
            capture_output=True, text=True, encoding="utf-8", timeout=60
        )
        if result.returncode != 0:
            raise RuntimeError(result.stderr[:300])
        rows_data: list[list] = json.loads(result.stdout.strip())
        for titulo, total_v in rows_data:
            norm_t = _norm(titulo)
            extras.append({
                "titulo":          titulo,
                "cpb":             "",
                "pontuacao_total": total_v,
                "festivais":       [],
                "_titulo_norm":    norm_t,
                "_fonte":          "Festivais_por_obra.xlsx (pesquisa bibliográfica / Claude, mar/2026)",
            })
    except Exception as e:
        print(f"  [AVISO] xlsx complement: {e}")
    return extras


# ─── 4. Salvar datasets ──────────────────────────────────────────────────────

def save_datasets(films_ata: list[dict]):
    os.makedirs(os.path.dirname(OUT_GRANULAR), exist_ok=True)
    os.makedirs(os.path.dirname(OUT_RESUMO), exist_ok=True)

    # ── 4a. Granular: um row por (obra × festival) ───────────────────────────
    with open(OUT_GRANULAR, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=[
            "cpb", "titulo", "titulo_norm",
            "festival", "pontos_festival", "pontuacao_total_obra",
            "fonte",
        ])
        writer.writeheader()
        for film in films_ata:
            tn = _norm(film["titulo"])
            if film["festivais"]:
                for (fest, pts) in film["festivais"]:
                    writer.writerow({
                        "cpb":                   film["cpb"],
                        "titulo":                film["titulo"],
                        "titulo_norm":           tn,
                        "festival":              fest,
                        "pontos_festival":       pts,
                        "pontuacao_total_obra":  film["pontuacao_total"],
                        "fonte":                 FONTE_LABEL,
                    })
            else:
                # Obra sem detalhe de festival (apenas total)
                writer.writerow({
                    "cpb":                   film["cpb"],
                    "titulo":                film["titulo"],
                    "titulo_norm":           tn,
                    "festival":              "",
                    "pontos_festival":       "",
                    "pontuacao_total_obra":  film["pontuacao_total"],
                    "fonte":                 film.get("_fonte", FONTE_LABEL),
                })

    print(f"  Salvo: {OUT_GRANULAR}  ({sum(max(1, len(f['festivais'])) for f in films_ata)} linhas)")

    # ── 4b. Resumo: um row por obra ──────────────────────────────────────────
    with open(OUT_RESUMO, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=[
            "cpb", "titulo", "titulo_norm",
            "pontuacao_total", "n_festivais",
            "lista_festivais", "fonte",
        ])
        writer.writeheader()
        for film in films_ata:
            tn = _norm(film["titulo"])
            lista = " | ".join(f[0] for f in film["festivais"]) if film["festivais"] else ""
            writer.writerow({
                "cpb":             film["cpb"],
                "titulo":          film["titulo"],
                "titulo_norm":     tn,
                "pontuacao_total": film["pontuacao_total"],
                "n_festivais":     len(film["festivais"]),
                "lista_festivais": lista,
                "fonte":           film.get("_fonte", FONTE_LABEL),
            })

    print(f"  Salvo: {OUT_RESUMO}  ({len(films_ata)} obras)")


# ─── 5. Combinar Ata + xlsx (para o resumo completo) ────────────────────────

def build_combined(films_ata: list[dict]) -> list[dict]:
    """Ata é primária; xlsx complementa obras não presentes na Ata."""
    norms_cobertos = {_norm(f["titulo"]) for f in films_ata}
    n_ata = len(films_ata)

    extras = load_xlsx_complement(norms_cobertos)
    adicionados = 0
    for ex in extras:
        if ex["_titulo_norm"] not in norms_cobertos:
            films_ata.append(ex)
            norms_cobertos.add(ex["_titulo_norm"])
            adicionados += 1
    print(f"  Obras da Ata: {n_ata}")
    print(f"  Obras complementadas via xlsx: {adicionados}")
    return films_ata


# ─── main ────────────────────────────────────────────────────────────────────

def main():
    print("=== 00_extrair_ata_fsa2024.py ===")
    print("  Lendo PDF...")
    lines = extract_pdf_text()
    print(f"  Linhas extraídas: {len(lines)}")

    print("  Parseando tabela de pontuação...")
    films_ata = parse_scoring_table(lines)
    print(f"  Obras encontradas na Ata: {len(films_ata)}")

    # Sanity check: mostrar primeiras 5
    for f in films_ata[:5]:
        print(f"    {f['cpb']} | {f['titulo'][:40]:<40} | total={f['pontuacao_total']} | n_fest={len(f['festivais'])}")

    print("  Complementando com xlsx (pesquisa Claude)...")
    all_films = build_combined(films_ata)

    print("  Salvando datasets...")
    save_datasets(all_films)
    print("  Concluído.")


if __name__ == "__main__":
    main()
