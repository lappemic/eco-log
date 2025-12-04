"""
econstruct Umweltbelastungsrechner - Streamlit App
"""
import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

from src.parser import parse_mengenliste, parse_oekobilanz
from src.matcher import MaterialMatcher
from src.calculator import UBPCalculator, CalculationResults

# Configurable path for Oekobilanz database (via environment variable or default)
OEKOBILANZ_PATH = Path(os.environ.get(
    "OEKOBILANZ_FILE",
    "Oekobilanzdaten_ Baubereich_Donne_ecobilans_construction_2009-1-2022_v7.0.xlsx"
))


# Seitenkonfiguration
st.set_page_config(
    page_title="econstruct UBP-Rechner",
    page_icon="🌿",
    layout="wide"
)


def format_number(n: float) -> str:
    """Formatiere grosse Zahlen mit Tausendertrennzeichen."""
    return f"{n:,.0f}".replace(",", "'")


def create_material_chart(results: CalculationResults) -> go.Figure:
    """Erstelle Balkendiagramm der UBP nach Material."""
    if not results.by_material:
        return None

    data = sorted(results.by_material.items(), key=lambda x: -x[1]["ubp"])
    materials = [d[0] for d in data]
    ubps = [d[1]["ubp"] for d in data]

    fig = px.bar(
        x=materials,
        y=ubps,
        labels={"x": "Material", "y": "UBP"},
        title="UBP nach Basismaterial",
        color=ubps,
        color_continuous_scale="Reds"
    )
    fig.update_layout(
        xaxis_tickangle=-45,
        showlegend=False,
        coloraxis_showscale=False
    )
    return fig


def create_coating_chart(results: CalculationResults) -> go.Figure:
    """Erstelle Balkendiagramm der UBP nach Beschichtung."""
    if not results.by_coating:
        return None

    data = sorted(results.by_coating.items(), key=lambda x: -x[1]["ubp"])
    coatings = [d[0] for d in data]
    ubps = [d[1]["ubp"] for d in data]

    fig = px.bar(
        x=coatings,
        y=ubps,
        labels={"x": "Beschichtung", "y": "UBP"},
        title="UBP nach Beschichtung",
        color=ubps,
        color_continuous_scale="Blues"
    )
    fig.update_layout(
        xaxis_tickangle=-45,
        showlegend=False,
        coloraxis_showscale=False
    )
    return fig


def create_pareto_chart(results: CalculationResults) -> go.Figure:
    """Erstelle Pareto-Diagramm der Top-Komponenten nach UBP."""
    top_components = sorted(
        results.components,
        key=lambda x: -x.ubp_total
    )[:15]

    if not top_components:
        return None

    labels = [f"Pos {int(c.pos)}: {c.bezeichnung[:20]}" for c in top_components]
    ubps = [c.ubp_total for c in top_components]
    cumulative = []
    running = 0
    for ubp in ubps:
        running += ubp
        cumulative.append(running / results.total_ubp * 100 if results.total_ubp > 0 else 0)

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=labels,
        y=ubps,
        name="UBP",
        marker_color="steelblue"
    ))

    fig.add_trace(go.Scatter(
        x=labels,
        y=cumulative,
        name="Kumuliert %",
        yaxis="y2",
        line=dict(color="red", width=2),
        marker=dict(size=8)
    ))

    fig.update_layout(
        title="Top 15 Bauteile nach UBP (Pareto)",
        xaxis_tickangle=-45,
        yaxis=dict(title="UBP"),
        yaxis2=dict(
            title="Kumuliert %",
            overlaying="y",
            side="right",
            range=[0, 100]
        ),
        legend=dict(x=0.8, y=1.1),
        height=500
    )
    return fig


def create_pie_chart(results: CalculationResults) -> go.Figure:
    """Erstelle Kreisdiagramm der Materialverteilung."""
    if not results.by_material:
        return None

    all_sources = {}

    for mat, data in results.by_material.items():
        all_sources[f"Material: {mat}"] = data["ubp"]

    for coat, data in results.by_coating.items():
        all_sources[f"Beschichtung: {coat}"] = data["ubp"]

    labels = list(all_sources.keys())
    values = list(all_sources.values())

    fig = px.pie(
        names=labels,
        values=values,
        title="UBP-Verteilung (Materialien + Beschichtungen)",
        hole=0.4
    )
    fig.update_traces(textposition="inside", textinfo="percent+label")
    return fig


def main():
    st.title("🌿 econstruct Umweltbelastungsrechner")
    st.markdown("Berechne UBP (Umweltbelastungspunkte) für HiCAD-Modellexporte")

    # Seitenleiste für Datei-Upload und Einstellungen
    with st.sidebar:
        st.header("📁 Modell hochladen")
        uploaded_file = st.file_uploader(
            "HiCAD Excel-Export hochladen",
            type=["xlsx"],
            help="Laden Sie die aus HiCAD exportierte Excel-Datei mit dem Blatt 'Mengenliste' hoch"
        )

        st.divider()
        st.header("⚙️ Einstellungen")

        # Prüfe ob Oekobilanz-Datei vorhanden
        if OEKOBILANZ_PATH.exists():
            st.success("✅ Oekobilanz-Datenbank geladen")
        else:
            st.warning("⚠️ Oekobilanz-Datenbank nicht gefunden")
            st.markdown("Platzieren Sie die Oekobilanz Excel-Datei im App-Verzeichnis")

    # Hauptinhalt
    if uploaded_file is None:
        st.info("👆 Laden Sie einen HiCAD Excel-Export hoch, um zu beginnen")

        st.markdown("### So funktioniert's")
        st.markdown("""
        1. **Hochladen** Sie Ihren HiCAD-Modellexport (.xlsx)
        2. Der Rechner **verknüpft** Materialien mit der Schweizer Oekobilanz-Datenbank
        3. **Ansehen** der Umweltbelastung aufgeschlüsselt nach Material und Bauteil
        4. **Herunterladen** der detaillierten Ergebnisse als CSV
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
    st.header("📈 Visualisierungen")

    tab1, tab2, tab3, tab4 = st.tabs([
        "Nach Material", "Nach Beschichtung", "Top Bauteile", "Verteilung"
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

    with tab4:
        fig = create_pie_chart(results)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Keine Daten vorhanden")

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

    # Download-Button
    csv = results_df.to_csv(index=False)
    st.download_button(
        label="📥 Ergebnisse als CSV herunterladen",
        data=csv,
        file_name="ubp_ergebnisse.csv",
        mime="text/csv"
    )

    # Warnung für nicht zugeordnete Materialien
    if results.unmatched:
        st.header("⚠️ Nicht zugeordnete Materialien")
        st.warning(f"{len(results.unmatched)} Bauteile konnten nicht mit der Oekobilanz-Datenbank verknüpft werden")

        unmatched_df = pd.DataFrame(results.unmatched).rename(columns={
            "pos": "Pos",
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
