"""
éco-log - Streamlit App
"""
import io
import os
import re
import hashlib
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import extra_streamlit_components as stx

# Professional environmental color palette
# Low impact = green tones, High impact = red/orange tones
ENV_COLORS = {
    "low": "#2D6A4F",       # Forest green (low impact)
    "medium": "#95D5B2",    # Light sage
    "high": "#E9C46A",      # Warm yellow
    "very_high": "#E76F51", # Terracotta (high impact)
    "critical": "#9B2226",  # Deep red (critical)

    # Chart accents
    "primary": "#264653",   # Deep teal
    "secondary": "#287271", # Ocean teal
    "accent": "#E9C46A",    # Golden yellow
    "line": "#E76F51",      # Accent line

    # Coating specific (blue spectrum)
    "coating_low": "#48CAE4",
    "coating_high": "#023E8A",
}

# Environmental impact gradient (low to high)
ENV_GRADIENT = ["#2D6A4F", "#40916C", "#74C69D", "#E9C46A", "#F4A261", "#E76F51", "#9B2226"]

from src.parser import parse_mengenliste, parse_oekobilanz
from src.matcher import MaterialMatcher
from src.calculator import UBPCalculator, CalculationResults

# Configurable path for Oekobilanz database (via environment variable or default)
OEKOBILANZ_PATH = Path(os.environ.get(
    "OEKOBILANZ_FILE",
    "Oekobilanzdaten_ Baubereich_Donne_ecobilans_construction_2009-1-2022_v7.0.xlsx"
))


APP_PASSWORD = "econstruct2025!"
AUTH_COOKIE_NAME = "eco_log_auth"
AUTH_TOKEN = hashlib.sha256(APP_PASSWORD.encode()).hexdigest()[:32]


def check_password() -> bool:
    """Password protection with cookie persistence."""
    cookie_manager = stx.CookieManager(key="eco_log_cookies")

    # Check session state first (fastest)
    if st.session_state.get("authenticated"):
        return True

    # Check cookie for persistent login
    auth_cookie = cookie_manager.get(AUTH_COOKIE_NAME)
    if auth_cookie == AUTH_TOKEN:
        st.session_state.authenticated = True
        return True

    # Show login form
    st.title("🔒 Anmeldung erforderlich")
    password = st.text_input("Passwort", type="password")
    if password:
        if password == APP_PASSWORD:
            st.session_state.authenticated = True
            # Set cookie for 30 days
            cookie_manager.set(AUTH_COOKIE_NAME, AUTH_TOKEN, max_age=30*24*60*60)
            st.rerun()
        else:
            st.error("Falsches Passwort")
    return False


# Seitenkonfiguration
st.set_page_config(
    page_title="éco-log",
    page_icon="♻️",
    layout="wide"
)


def format_number(n: float) -> str:
    """Formatiere grosse Zahlen mit Tausendertrennzeichen."""
    return f"{n:,.0f}".replace(",", "'")


def normalize_filename(filename: str) -> str:
    """Clean and normalize a filename for use in export."""
    # Remove extension
    name = Path(filename).stem
    # Replace spaces and special chars with underscores
    name = re.sub(r"[^\w\-]", "_", name)
    # Collapse multiple underscores
    name = re.sub(r"_+", "_", name)
    # Strip leading/trailing underscores
    return name.strip("_")


def create_material_chart(results: CalculationResults) -> go.Figure:
    """Erstelle horizontales Balkendiagramm der UBP nach Material."""
    if not results.by_material:
        return None

    # Sort by UBP (lowest at top for horizontal bars → highest impact at top visually)
    data = sorted(results.by_material.items(), key=lambda x: x[1]["ubp"])
    materials = [d[0] for d in data]
    ubps = [d[1]["ubp"] for d in data]
    weights = [d[1]["weight_kg"] for d in data]

    # Calculate percentages
    total_material_ubp = sum(ubps)
    percentages = [(u / total_material_ubp * 100) if total_material_ubp > 0 else 0 for u in ubps]

    # Create horizontal bar chart
    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=materials,
        x=ubps,
        orientation='h',
        marker=dict(
            color=ubps,
            colorscale=[[0, ENV_COLORS["low"]], [0.5, ENV_COLORS["accent"]], [1, ENV_COLORS["very_high"]]],
            line=dict(width=0),
        ),
        text=[f"<b>{format_number(u)}</b> UBP ({p:.0f}%)" for u, p in zip(ubps, percentages)],
        textposition='auto',
        textfont=dict(color='white', size=11),
        hovertemplate="<b>%{y}</b><br>" +
                      "Umweltbelastung: %{x:,.0f} UBP<br>" +
                      "<extra></extra>",
    ))

    # Find the highest impact material
    max_idx = len(ubps) - 1  # After sorting, highest is last
    max_material = materials[max_idx]
    max_pct = percentages[max_idx]

    fig.update_layout(
        title=dict(
            text=f"Umweltbelastung nach Material<br><sub>{max_material} verursacht {max_pct:.0f}% der Material-UBP</sub>",
            font=dict(size=16, color=ENV_COLORS["primary"]),
        ),
        xaxis=dict(
            title="Umweltbelastungspunkte (UBP)",
            showgrid=True,
            gridcolor='rgba(0,0,0,0.1)',
        ),
        yaxis=dict(
            title="",
            tickfont=dict(size=11),
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=200, r=20, t=80, b=40),
        height=max(350, len(materials) * 40 + 100),
        showlegend=False,
    )

    return fig


def create_coating_chart(results: CalculationResults) -> go.Figure:
    """Erstelle horizontales Balkendiagramm der UBP nach Beschichtung."""
    if not results.by_coating:
        return None

    # Sort by UBP (lowest at top → highest impact at top visually)
    data = sorted(results.by_coating.items(), key=lambda x: x[1]["ubp"])
    coatings = [d[0] for d in data]
    ubps = [d[1]["ubp"] for d in data]
    areas = [d[1]["area_m2"] for d in data]

    # Calculate percentages
    total_coating_ubp = sum(ubps)
    percentages = [(u / total_coating_ubp * 100) if total_coating_ubp > 0 else 0 for u in ubps]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=coatings,
        x=ubps,
        orientation='h',
        marker=dict(
            color=ubps,
            colorscale=[[0, ENV_COLORS["coating_low"]], [1, ENV_COLORS["coating_high"]]],
            line=dict(width=0),
        ),
        text=[f"<b>{format_number(u)}</b> UBP ({p:.0f}%)" for u, p in zip(ubps, percentages)],
        textposition='auto',
        textfont=dict(color='white', size=11),
        hovertemplate="<b>%{y}</b><br>" +
                      "Umweltbelastung: %{x:,.0f} UBP<br>" +
                      "<extra></extra>",
    ))

    # Find highest impact coating
    max_idx = len(ubps) - 1
    max_coating = coatings[max_idx] if coatings else "–"
    max_pct = percentages[max_idx] if percentages else 0

    fig.update_layout(
        title=dict(
            text=f"Umweltbelastung nach Beschichtung<br><sub>{max_coating} verursacht {max_pct:.0f}% der Beschichtungs-UBP</sub>",
            font=dict(size=16, color=ENV_COLORS["primary"]),
        ),
        xaxis=dict(
            title="Umweltbelastungspunkte (UBP)",
            showgrid=True,
            gridcolor='rgba(0,0,0,0.1)',
        ),
        yaxis=dict(
            title="",
            tickfont=dict(size=11),
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=220, r=20, t=80, b=40),
        height=max(300, len(coatings) * 45 + 100),
        showlegend=False,
    )

    return fig


def create_pareto_chart(results: CalculationResults) -> go.Figure:
    """Erstelle Pareto-Diagramm mit 80/20-Prinzip Hervorhebung."""
    top_components = sorted(
        results.components,
        key=lambda x: -x.ubp_total
    )[:15]

    if not top_components:
        return None

    def make_label(c):
        bez = c.bezeichnung[:20] + "…" if len(c.bezeichnung) > 20 else c.bezeichnung
        return f"Pos {int(c.pos)}: {bez} ({c.material})"

    labels = [make_label(c) for c in top_components]
    full_labels = [f"Pos {int(c.pos)}: {c.bezeichnung} ({c.material})" for c in top_components]
    ubps = [c.ubp_total for c in top_components]

    # Calculate cumulative percentage
    cumulative = []
    running = 0
    for ubp in ubps:
        running += ubp
        cumulative.append(running / results.total_ubp * 100 if results.total_ubp > 0 else 0)

    # Find where 80% threshold is crossed
    threshold_idx = next((i for i, c in enumerate(cumulative) if c >= 80), len(cumulative) - 1)

    # Create colors: highlight components that contribute to 80%
    bar_colors = [ENV_COLORS["very_high"] if i <= threshold_idx else ENV_COLORS["medium"]
                  for i in range(len(ubps))]

    fig = go.Figure()

    # Bars with gradient based on impact
    fig.add_trace(go.Bar(
        x=labels,
        y=ubps,
        name="UBP pro Bauteil",
        marker=dict(
            color=bar_colors,
            line=dict(width=0),
        ),
        text=[format_number(u) for u in ubps],
        textposition='outside',
        textfont=dict(size=10, color=ENV_COLORS["primary"]),
        hovertemplate="%{customdata}<br>UBP: %{y:,.0f}<extra></extra>",
        customdata=full_labels,
    ))

    # Cumulative line
    fig.add_trace(go.Scatter(
        x=labels,
        y=cumulative,
        name="Kumulierte Belastung",
        yaxis="y2",
        line=dict(color=ENV_COLORS["line"], width=3),
        marker=dict(size=8, color=ENV_COLORS["line"]),
        hovertemplate="%{y:.1f}% der Gesamtbelastung<extra></extra>",
    ))

    # 80% threshold line
    fig.add_hline(
        y=80, line_dash="dash", line_color=ENV_COLORS["primary"],
        line_width=2, yref="y2",
        annotation_text="80% Schwelle",
        annotation_position="right",
        annotation_font=dict(color=ENV_COLORS["primary"], size=11),
    )

    # Count components causing 80%
    components_for_80 = threshold_idx + 1
    pct_of_total = (components_for_80 / len(results.components) * 100) if results.components else 0

    fig.update_layout(
        title=dict(
            text=f"Pareto: Top {len(top_components)} Bauteile nach Umweltbelastung<br>" +
                 f"<sub><b>{components_for_80} Bauteile</b> ({pct_of_total:.0f}% aller Teile) verursachen <b>80%</b> der Belastung</sub>",
            font=dict(size=16, color=ENV_COLORS["primary"]),
        ),
        xaxis=dict(
            title="",
            tickangle=-45,
            tickfont=dict(size=10),
        ),
        yaxis=dict(
            title=dict(text="UBP pro Bauteil", font=dict(color=ENV_COLORS["primary"])),
            showgrid=True,
            gridcolor='rgba(0,0,0,0.08)',
        ),
        yaxis2=dict(
            title=dict(text="Kumuliert (%)", font=dict(color=ENV_COLORS["line"])),
            overlaying="y",
            side="right",
            range=[0, 105],
            showgrid=False,
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="center",
            x=0.5,
            font=dict(size=11),
        ),
        margin=dict(l=60, r=60, t=100, b=80),
        height=500,
        bargap=0.15,
    )

    return fig


def create_distribution_chart(results: CalculationResults) -> go.Figure:
    """Erstelle Treemap für hierarchische UBP-Verteilung."""
    if not results.by_material and not results.by_coating:
        return None

    # Calculate totals first
    total_material_ubp = sum(d["ubp"] for d in results.by_material.values())
    total_coating_ubp = sum(d["ubp"] for d in results.by_coating.values())
    total_ubp = total_material_ubp + total_coating_ubp

    # Build treemap data - root must equal sum of children for branchvalues="total"
    labels = ["Gesamt"]
    parents = [""]
    values = [total_ubp]
    colors = [ENV_COLORS["primary"]]

    if total_material_ubp > 0:
        labels.append("Materialien")
        parents.append("Gesamt")
        values.append(total_material_ubp)
        colors.append(ENV_COLORS["very_high"])

        # Add individual materials
        for mat, data in sorted(results.by_material.items(), key=lambda x: -x[1]["ubp"]):
            labels.append(mat)
            parents.append("Materialien")
            values.append(data["ubp"])
            # Gradient color based on relative impact
            pct = data["ubp"] / total_material_ubp if total_material_ubp > 0 else 0
            if pct > 0.3:
                colors.append(ENV_COLORS["critical"])
            elif pct > 0.15:
                colors.append(ENV_COLORS["very_high"])
            else:
                colors.append(ENV_COLORS["high"])

    if total_coating_ubp > 0:
        labels.append("Beschichtungen")
        parents.append("Gesamt")
        values.append(total_coating_ubp)
        colors.append(ENV_COLORS["coating_high"])

        # Add individual coatings
        for coat, data in sorted(results.by_coating.items(), key=lambda x: -x[1]["ubp"]):
            labels.append(coat)
            parents.append("Beschichtungen")
            values.append(data["ubp"])
            colors.append(ENV_COLORS["coating_low"])

    fig = go.Figure(go.Treemap(
        labels=labels,
        parents=parents,
        values=values,
        marker=dict(
            colors=colors,
            line=dict(width=2, color='white'),
        ),
        texttemplate="<b>%{label}</b><br>%{value:,.0f} UBP<br>%{percentRoot:.1%}",
        textfont=dict(size=12),
        hovertemplate="<b>%{label}</b><br>" +
                      "UBP: %{value:,.0f}<br>" +
                      "Anteil: %{percentRoot:.1%}<extra></extra>",
        branchvalues="total",
        pathbar=dict(visible=True),
    ))

    # Calculate material vs coating split
    mat_pct = (total_material_ubp / results.total_ubp * 100) if results.total_ubp > 0 else 0
    coat_pct = (total_coating_ubp / results.total_ubp * 100) if results.total_ubp > 0 else 0

    fig.update_layout(
        title=dict(
            text=f"UBP-Verteilung: Materialien vs. Beschichtungen<br>" +
                 f"<sub>Materialien: <b>{mat_pct:.0f}%</b> | Beschichtungen: <b>{coat_pct:.0f}%</b></sub>",
            font=dict(size=16, color=ENV_COLORS["primary"]),
        ),
        margin=dict(l=10, r=10, t=80, b=10),
        height=500,
    )

    return fig


def main():
    if not check_password():
        st.stop()

    st.title("éco-log")
    st.markdown("Berechne UBP (Umweltbelastungspunkte) für HiCAD-Modellexporte")

    # Seitenleiste für Datei-Upload und Einstellungen
    with st.sidebar:
        st.header("📁 Modell hochladen")
        uploaded_file = st.file_uploader(
            "HiCAD Excel-Export hochladen",
            type=["xlsx"],
            help="Lade die aus HiCAD exportierte Excel-Datei mit dem Blatt 'Mengenliste' hoch"
        )

        st.divider()
        st.header("⚙️ Einstellungen")

        # Prüfe ob Oekobilanz-Datei vorhanden
        if OEKOBILANZ_PATH.exists():
            st.success("✅ KBOB Ökobilanzdaten geladen")
        else:
            st.warning("⚠️ KBOB Ökobilanzdaten nicht gefunden")
            st.markdown("Platziere die KBOB Excel-Datei im App-Verzeichnis")

    # Hauptinhalt
    if uploaded_file is None:
        st.info("👈 Lade einen HiCAD Excel-Export hoch, um zu beginnen")

        st.markdown("### So funktioniert's")
        st.markdown("""
        1. **Lade** deinen HiCAD-Modellexport hoch (.xlsx)
        2. Der Rechner **verknüpft** Materialien mit den KBOB Ökobilanzdaten im Baubereich (2009/1:2022, V7.0)
        3. **Sieh** die Umweltbelastung aufgeschlüsselt nach Material und Bauteil
        4. **Lade** die detaillierten Ergebnisse als Excel herunter
        """)

        # Zeige unterstützte Materialien
        with st.expander("📋 Unterstützte Materialien"):
            matcher = MaterialMatcher()
            mappings = matcher.get_all_mappings()

            st.markdown("**Basismaterialien:**")
            for mat, data in mappings["materials"].items():
                st.markdown(f"- `{mat}` → {data['oeko_name']} ({data['ubp_per_kg']} UBP/kg)")

            st.markdown("\n**Beschichtungen:**")
            for coat, data in mappings["coatings"].items():
                st.markdown(f"- `{coat}` → {data['oeko_name']} ({data['ubp_per_m2']} UBP/m²)")

        return

    # Verarbeite hochgeladene Datei
    try:
        with st.spinner("Excel-Datei wird analysiert..."):
            df = parse_mengenliste(uploaded_file)
            st.success(f"✅ {len(df)} Bauteile aus Mengenliste geladen")

        with st.spinner("UBP wird berechnet..."):
            calculator = UBPCalculator()
            results = calculator.calculate(df)
            summary = calculator.summary_to_dict(results)

    except FileNotFoundError as e:
        st.error(f"Datei nicht gefunden: {e}")
        return
    except ValueError as e:
        st.error(f"Ungültiges Datenformat: {e}")
        return
    except KeyError as e:
        st.error(f"Erforderliche Spalte fehlt: {e}")
        return
    except pd.errors.EmptyDataError:
        st.error("Die Excel-Datei ist leer oder enthält keine gültigen Daten")
        return
    except Exception as e:
        st.error(f"Unerwarteter Fehler bei der Verarbeitung: {e}")
        return

    # Zusammenfassung
    st.header("📊 Zusammenfassung")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Total UBP", format_number(summary["total_ubp"]))
    with col2:
        st.metric("Gesamtgewicht", f"{summary['total_weight_kg']:,.1f} kg")
    with col3:
        st.metric("Trefferquote", f"{summary['match_rate']}%")
    with col4:
        st.metric("Nicht zugeordnet", str(summary["unmatched_count"]))

    # Diagramme
    st.header("📈 Umweltanalyse")

    tab1, tab2, tab3 = st.tabs([
        "🔩 Material-Impact", "🎨 Beschichtungs-Impact", "📊 Pareto (80/20)"
    ])

    with tab1:
        fig = create_material_chart(results)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Keine Materialdaten vorhanden")

    with tab2:
        fig = create_coating_chart(results)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Keine Beschichtungsdaten vorhanden")

    with tab3:
        fig = create_pareto_chart(results)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Keine Bauteil-Daten vorhanden")

    # Detaillierte Ergebnistabelle
    st.header("📋 Detaillierte Ergebnisse")

    results_df = calculator.results_to_dataframe(results)

    # Filteroptionen
    col1, col2 = st.columns(2)
    with col1:
        show_matched = st.checkbox("Nur zugeordnete anzeigen", value=False)
    with col2:
        sort_by = st.selectbox(
            "Sortieren nach",
            ["UBP Total", "Pos", "Gewicht (kg)"],
            index=0
        )

    display_df = results_df.copy()
    if show_matched:
        display_df = display_df[display_df["Zugeordnet"] == "Ja"]

    display_df = display_df.sort_values(sort_by, ascending=False)

    st.dataframe(
        display_df,
        use_container_width=True,
        height=400
    )

    # Download-Button (Excel mit Admin-Review-Spalten)
    export_df = calculator.results_to_dataframe(results, include_review_columns=True)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="UBP Ergebnisse")
        # Auto-adjust column widths
        worksheet = writer.sheets["UBP Ergebnisse"]
        for idx, col in enumerate(export_df.columns, 1):
            # Calculate width based on column name and content
            max_length = len(str(col))
            for value in export_df[col].astype(str):
                max_length = max(max_length, len(value))
            # Set width with padding, min 10, max 40
            adjusted_width = min(max(max_length + 2, 10), 40)
            worksheet.column_dimensions[worksheet.cell(1, idx).column_letter].width = adjusted_width
    excel_buffer.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    source_name = normalize_filename(uploaded_file.name)
    st.download_button(
        label="📥 Ergebnisse als Excel herunterladen",
        data=excel_buffer,
        file_name=f"{timestamp}_eco-log_{source_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Warnung für nicht zugeordnete Materialien
    if results.unmatched:
        st.header("⚠️ Nicht zugeordnete Materialien")
        st.warning(f"{len(results.unmatched)} Bauteile konnten nicht mit den KBOB Ökobilanzdaten verknüpft werden")

        unmatched_df = pd.DataFrame(results.unmatched).rename(columns={
            "pos": "Pos",
            "bezeichnung": "Bezeichnung",
            "material": "Material",
            "typ": "Typ",
            "weight_kg": "Gewicht (kg)",
            "reason": "Grund"
        })
        st.dataframe(unmatched_df, use_container_width=True)

        # Gruppiere nach Material für Zusammenfassung
        material_counts = unmatched_df.groupby("Material").agg({
            "Pos": "count",
            "Gewicht (kg)": "sum"
        }).rename(columns={"Pos": "Anzahl"}).sort_values("Anzahl", ascending=False)

        st.markdown("**Zusammenfassung nicht zugeordneter Materialien:**")
        st.dataframe(material_counts, use_container_width=True)


if __name__ == "__main__":
    main()
