from __future__ import annotations

from pathlib import Path
import textwrap

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import Rectangle
import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_FESTIVAL_CSV = BASE_DIR / "base_festivais_obras_ata.csv"
INPUT_PRODUCER_CSV = BASE_DIR.parent.parent / "raw" / "produtores-de-obras-nao-publicitarias-brasileiras.csv"
OUTPUT_PDF = BASE_DIR / "relatorio_por_produtora_festivais_tabela.pdf"


def format_points(value: float) -> str:
    return f"{value:.1f}".replace(".", ",")


def wrap_text(value: str, width: int) -> str:
    return "\n".join(textwrap.wrap(str(value), width=width, break_long_words=False))


def load_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    festivals = pd.read_csv(INPUT_FESTIVAL_CSV)
    festivals["pontos_festival"] = pd.to_numeric(festivals["pontos_festival"], errors="coerce").fillna(0.0)
    festivals["titulo"] = festivals["titulo"].fillna(festivals["titulo_norm"])
    festivals["titulo_norm"] = festivals["titulo_norm"].fillna(festivals["titulo"].str.lower())
    festivals = festivals[festivals["cpb"].notna()].copy()

    producers = pd.read_csv(
        INPUT_PRODUCER_CSV,
        sep=";",
        encoding="utf-8",
        engine="python",
        on_bad_lines="skip",
    )
    producers = producers[["CPB", "PRODUTOR", "CNPJ_PRODUTOR", "PAIS_PRODUTOR"]].drop_duplicates(subset=["CPB"])
    return festivals, producers


def build_work_summary(festivals: pd.DataFrame) -> pd.DataFrame:
    work = (
        festivals.groupby(["cpb", "titulo", "titulo_norm"], dropna=False)
        .agg(
            total_pontos=("pontos_festival", "sum"),
            qtd_festivais=("festival", "nunique"),
        )
        .reset_index()
    )
    return work


def build_festival_top_summary(festivals: pd.DataFrame) -> dict[str, str]:
    top = (
        festivals.groupby(["cpb", "festival"], dropna=False, as_index=False)
        .agg(pontos_festival=("pontos_festival", "sum"))
        .sort_values(["cpb", "pontos_festival", "festival"], ascending=[True, False, True])
    )

    summary: dict[str, str] = {}
    for cpb, group in top.groupby("cpb"):
        pieces: list[str] = []
        for _, row in group.head(3).iterrows():
            pieces.append(f"{row['festival']} ({format_points(row['pontos_festival'])})")
        summary[cpb] = " | ".join(pieces)
    return summary


def build_producer_rows(
    work_summary: pd.DataFrame,
    festivals: pd.DataFrame,
    producers: pd.DataFrame,
    top_festival_summary: dict[str, str],
) -> list[dict[str, str]]:
    work = work_summary.merge(producers, left_on="cpb", right_on="CPB", how="left")
    work["PRODUTOR"] = work["PRODUTOR"].fillna("SEM PRODUTORA IDENTIFICADA")
    work["CNPJ_PRODUTOR"] = work["CNPJ_PRODUTOR"].fillna("")

    producer_summary = (
        work.groupby(["PRODUTOR", "CNPJ_PRODUTOR"], dropna=False)
        .agg(
            pontos_totais=("total_pontos", "sum"),
            qtd_obras=("cpb", "nunique"),
        )
        .reset_index()
        .sort_values(["pontos_totais", "qtd_obras", "PRODUTOR"], ascending=[False, False, True])
    )

    rows: list[dict[str, str]] = []
    for _, producer in producer_summary.iterrows():
        producer_name = str(producer["PRODUTOR"])
        producer_cnpj = str(producer["CNPJ_PRODUTOR"])
        producer_points = format_points(producer["pontos_totais"])
        producer_works = str(int(producer["qtd_obras"]))

        rows.append(
            {
                "type": "producer",
                "producer": producer_name,
                "cnpj": producer_cnpj,
                "work": "",
                "cpb": "",
                "points": producer_points,
                "festivals": "",
                "works_count": producer_works,
            }
        )

        producer_works_df = (
            work.loc[work["PRODUTOR"] == producer_name]
            .sort_values(["total_pontos", "titulo", "cpb"], ascending=[False, True, True])
            .reset_index(drop=True)
        )

        for _, w in producer_works_df.iterrows():
            rows.append(
                {
                    "type": "work",
                    "producer": "",
                    "cnpj": "",
                    "work": str(w["titulo"]),
                    "cpb": str(w["cpb"]),
                    "points": format_points(w["total_pontos"]),
                    "festivals": top_festival_summary.get(str(w["cpb"]), ""),
                    "works_count": str(int(w["qtd_festivais"])),
                }
            )

    return rows


def row_height(row: dict[str, str]) -> float:
    if row["type"] == "producer":
        return 0.045
    work_lines = max(1, row["work"].count("\n") + 1)
    fest_lines = max(1, row["festivals"].count("\n") + 1) if row["festivals"] else 1
    line_count = max(work_lines, fest_lines)
    return 0.032 + 0.017 * line_count


def paginate(rows: list[dict[str, str]], available_height: float = 0.70) -> list[list[dict[str, str]]]:
    pages: list[list[dict[str, str]]] = []
    current: list[dict[str, str]] = []
    used = 0.0
    for row in rows:
        h = row_height(row)
        if current and used + h > available_height:
            pages.append(current)
            current = []
            used = 0.0
        current.append(row)
        used += h
    if current:
        pages.append(current)
    return pages


def draw_page(
    fig: plt.Figure,
    rows: list[dict[str, str]],
    page_num: int,
    total_pages: int,
    n_producers: int,
    n_works: int,
    n_festivals: int,
) -> None:
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    ax.add_patch(Rectangle((0, 0.905), 1, 0.095, facecolor="#111827", edgecolor="none"))
    ax.text(0.03, 0.968, "Listagem por produtora", fontsize=20, fontweight="bold", color="white", ha="left", va="top")
    ax.text(
        0.03,
        0.935,
        "Produtoras ordenadas por pontos totais; obras ordenadas por pontos dentro de cada produtora.",
        fontsize=10.5,
        color="#D1D5DB",
        ha="left",
        va="top",
    )
    ax.text(0.97, 0.968, f"Página {page_num}/{total_pages}", fontsize=10.5, color="#D1D5DB", ha="right", va="top")

    ax.text(0.03, 0.87, f"Produtoras: {n_producers}", fontsize=9.5, color="#374151", ha="left", va="top")
    ax.text(0.18, 0.87, f"Obras: {n_works}", fontsize=9.5, color="#374151", ha="left", va="top")
    ax.text(0.28, 0.87, f"Festivais: {n_festivals}", fontsize=9.5, color="#374151", ha="left", va="top")
    ax.text(0.42, 0.87, "Obras por produtora em ordem decrescente de pontos", fontsize=9.5, color="#6B7280", ha="left", va="top")

    x0 = 0.03
    x1 = 0.46
    x2 = 0.62
    x3 = 0.72
    x4 = 0.97
    header_y = 0.82
    top_y = 0.79
    bottom_y = 0.06

    ax.add_patch(Rectangle((x0, header_y), x4 - x0, 0.04, facecolor="#E5E7EB", edgecolor="#D1D5DB"))
    header_labels = [
        ("Produtora / Obra", x0 + 0.01),
        ("CNPJ / CPB", x1 + 0.01),
        ("Pontos", x2 + 0.01),
        ("Festivais principais", x3 + 0.01),
        ("Qtd.", x4 - 0.055),
    ]
    for label, x in header_labels:
        ax.text(x, header_y + 0.022, label, fontsize=9.3, fontweight="bold", color="#111827", ha="left", va="center")

    y = top_y
    odd = False
    for row in rows:
        h = row_height(row)
        if y - h < bottom_y:
            break
        if row["type"] == "producer":
            ax.add_patch(Rectangle((x0, y - h), x4 - x0, h, facecolor="#DCE7F7", edgecolor="#9CB7E2", linewidth=0.9))
            ax.text(x0 + 0.01, y - 0.010, row["producer"], fontsize=9.0, fontweight="bold", ha="left", va="top", color="#111827")
            if row["cnpj"]:
                ax.text(x1 + 0.01, y - 0.010, row["cnpj"], fontsize=8.3, ha="left", va="top", color="#111827")
            ax.text(x2 + 0.01, y - 0.010, row["points"], fontsize=9.0, fontweight="bold", ha="left", va="top", color="#111827")
            ax.text(x3 + 0.01, y - 0.010, f"{row['works_count']} obras", fontsize=8.8, ha="left", va="top", color="#111827")
        else:
            fill = "#FFFFFF" if not odd else "#F9FAFB"
            odd = not odd
            ax.add_patch(Rectangle((x0, y - h), x4 - x0, h, facecolor=fill, edgecolor="#D1D5DB", linewidth=0.7))
            ax.plot([x1, x1], [y - h, y], color="#D1D5DB", linewidth=0.8)
            ax.plot([x2, x2], [y - h, y], color="#D1D5DB", linewidth=0.8)
            ax.plot([x3, x3], [y - h, y], color="#D1D5DB", linewidth=0.8)
            ax.text(x0 + 0.015, y - 0.010, "  " + wrap_text(row["work"], 28), fontsize=8.2, ha="left", va="top", color="#111827", linespacing=1.1)
            ax.text(x1 + 0.01, y - 0.010, row["cpb"], fontsize=8.2, ha="left", va="top", color="#111827")
            ax.text(x2 + 0.01, y - 0.010, row["points"], fontsize=8.2, ha="left", va="top", color="#111827")
            ax.text(x3 + 0.01, y - 0.010, wrap_text(row["festivals"], 40), fontsize=8.0, ha="left", va="top", color="#111827", linespacing=1.07)
            ax.text(x4 - 0.055, y - 0.010, row["works_count"], fontsize=8.2, ha="left", va="top", color="#111827")

        y -= h

    ax.text(
        0.03,
        0.03,
        "Cada produtora aparece com seu total agregado; abaixo, as obras da mesma produtora já vêm ordenadas pela pontuação da obra.",
        fontsize=8.5,
        color="#6B7280",
        ha="left",
        va="bottom",
    )


def main() -> None:
    festivals, producers = load_data()
    work_summary = build_work_summary(festivals)
    top_festival_summary = build_festival_top_summary(festivals)
    rows = build_producer_rows(work_summary, festivals, producers, top_festival_summary)
    pages = paginate(rows)

    with PdfPages(OUTPUT_PDF) as pdf:
        for page_num, page_rows in enumerate(pages, start=1):
            fig = plt.figure(figsize=(11.69, 8.27))
            draw_page(
                fig,
                page_rows,
                page_num,
                len(pages),
                n_producers=work_summary.merge(producers, left_on="cpb", right_on="CPB", how="left")["PRODUTOR"].fillna("SEM PRODUTORA IDENTIFICADA").nunique(),
                n_works=work_summary["cpb"].nunique(),
                n_festivals=festivals["festival"].nunique(),
            )
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

    print(OUTPUT_PDF)


if __name__ == "__main__":
    main()
