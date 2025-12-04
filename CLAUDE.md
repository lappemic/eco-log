# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

econstruct Umweltbelastungsrechner (Environmental Impact Calculator) - A Streamlit web application that calculates UBP (Umweltbelastungspunkte / Environmental Impact Points) for HiCAD model exports by matching materials against the Swiss Oekobilanz construction database (KBOB/eco-bau/IPB).

## Commands

```bash
# Install dependencies
uv sync

# Run the Streamlit app
uv run streamlit run app.py

# Test parser/matcher modules directly
uv run python -m src.parser
uv run python -m src.matcher
uv run python -m src.calculator
```

## Architecture

### Data Flow
1. User uploads HiCAD Excel export (.xlsx with "Mengenliste" sheet)
2. `parser.py` extracts component data (material, type, weight, area, coating)
3. `matcher.py` maps HiCAD material codes to Oekobilanz database entries
4. `calculator.py` computes UBP for each component and aggregates results
5. `app.py` displays results with Plotly visualizations

### Key Modules

- **app.py**: Streamlit UI with charts (Pareto, pie, bar) and data tables
- **src/parser.py**: Excel parsing for both HiCAD exports and Oekobilanz database
- **src/matcher.py**: Material/coating matching with type overrides (profile vs sheet)
- **src/calculator.py**: UBP calculation engine with dataclass result models
- **data/material_map.json**: HiCAD-to-Oekobilanz mapping configuration

### Material Matching Logic

The matcher differentiates:
- **Base materials**: Steel (S235JR), aluminum (EN AW-6060), stainless steel (X5CrNi18-10)
- **Type overrides**: Same material has different UBP when used as sheet vs profile
- **Coatings**: Surface treatments (galvanized, powder-coated) with aluminum-specific overrides

### Data Sources

- **Input**: HiCAD Excel exports with "Mengenliste" sheet (header row 8)
- **Reference**: Swiss Oekobilanz database (Oekobilanzdaten Baubereich 2009/1:2022 v7.0)
- **UBP units**: Per kilogram for materials, per square meter for coatings

## Domain Terms

- **UBP**: Umweltbelastungspunkte (Environmental Impact Points using UBP'21 methodology)
- **Mengenliste**: Bill of materials/quantities from HiCAD
- **Oekobilanz**: Life cycle assessment/eco-balance database
- **Feuerverzinkt**: Hot-dip galvanized coating
- **Pulverbeschichtung**: Powder coating
