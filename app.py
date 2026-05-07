"""
Axians PTToC Configurator
Run: streamlit run app.py
Keep pttoc_pricing.xlsx in the same folder.
"""

import streamlit as st
import pandas as pd
import openpyxl
import plotly.graph_objects as go
import os

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "pttoc_pricing.xlsx")
SALES_PWD  = "axians2024"

PACKAGES = ["Essentials", "Standard", "Premium"]
PKG_COLOR = {"Essentials": "#4472C4", "Standard": "#ED7D31", "Premium": "#2E7D32"}

FEATURES = {
    "Devices":          {"Essentials": "Crosscall (fixed)", "Standard": "Crosscall (fixed)", "Premium": "CYOD"},
    "PTT Application":  {"Essentials": "Integration only",  "Standard": "Integration only",  "Premium": "Integration + Custom UI"},
    "SOTI MDM":         {"Essentials": "–",                 "Standard": "Mobicontrol",       "Premium": "Mobicontrol XS"},
    "SIM":              {"Essentials": "–",                 "Standard": "Voice + Data",      "Premium": "Voice + Data"},
    "KPI Monitoring":   {"Essentials": "✔", "Standard": "✔", "Premium": "✔"},
    "Digital Training": {"Essentials": "✔", "Standard": "✔", "Premium": "✔"},
    "Update Consult.":  {"Essentials": "✔", "Standard": "✔", "Premium": "✔"},
    "On-site Training": {"Essentials": "–", "Standard": "✔", "Premium": "✔"},
    "Support":          {"Essentials": "Business hours", "Standard": "Business hours", "Premium": "24/7"},
    "Swap time":        {"Essentials": "< 96 h", "Standard": "< 48 h", "Premium": "< 24 h"},
    "Audit Service":    {"Essentials": "–", "Standard": "–", "Premium": "✔"},
}

# Add-ons that are included natively in a higher package
INCLUDED_IN = {
    "SOTI Mobicontrol MDM":    "Standard",
    "SOTI Mobicontrol XS MDM": "Premium",
    "SIM Voice + Data":        "Standard",
    "Custom PTT UI":           "Premium",
    "24/7 support upgrade":    "Premium",
}


@st.cache_data(ttl=60)
def load_addons(path: str):
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb["Cost and pricing"]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2 or not row[0]:
            continue
        name, _, pkg_formula, cost, type_, notes, available, conflicts = (list(row) + [None]*8)[:8]
        cost = float(cost) if cost else 0.0
        # pkg_price formula is like =D3*1.3; extract multiplier
        mult = 1.3
        if isinstance(pkg_formula, str) and "*" in pkg_formula:
            try:
                mult = float(pkg_formula.split("*")[1])
            except Exception:
                mult = 1.3
        pkg_price   = round(cost * mult, 2)
        addon_price = round(pkg_price * 1.3, 2)
        rows.append({
            "name":        str(name),
            "cost":        cost,
            "pkg_price":   pkg_price,
            "addon_price": addon_price,
            "type":        str(type_ or "recurring"),
            "notes":       str(notes or ""),
            "available":   str(available or "All"),
            "conflicts":   str(conflicts or ""),
        })
    return rows


def get_available(addons, pkg):
    pkg_l = pkg.lower()
    return [
        a for a in addons
        if a["available"].lower() == "all" or pkg_l in a["available"].lower()
    ]


def main():
    st.set_page_config(
        page_title="Axians PTToC",
        page_icon="📡",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    addons = load_addons(EXCEL_PATH)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 📡 Axians PTToC")
        st.markdown("---")

        mode = st.radio("Interface", ["🏢 Client", "💼 Sales"])
        is_sales = False
        if "Sales" in mode:
            pwd = st.text_input("Password", type="password")
            if pwd == SALES_PWD:
                is_sales = True
                st.success("Unlocked ✓")
            elif pwd:
                st.error("Wrong password")

        st.markdown("---")
        st.markdown("**Configuration**")
        n_devices = st.number_input("Devices", min_value=1, max_value=9999, value=50, step=5)
        duration  = st.slider("Contract (years)", 1, 5, 3)
        pkg       = st.selectbox("Base package", PACKAGES, index=1)

        st.markdown("---")
        st.markdown("**Add-ons**")
        available = get_available(addons, pkg)

        blocked = set()
        selected = []
        for ao in available:
            disabled = ao["name"] in blocked
            price_str = (
                f"€ {ao['addon_price']:,.2f}/yr"
                if ao["type"] == "recurring"
                else f"€ {ao['addon_price']:,.2f} one-off"
            )
            conflict_note = (
                f"\n⚠️ Conflicts with: {ao['conflicts']}"
                if ao["conflicts"] not in ("", "None", "nan")
                else ""
            )
            checked = st.checkbox(
                ao["name"],
                disabled=disabled,
                help=f"{price_str} · {ao['notes']}{conflict_note}",
            )
            st.caption(f"_{price_str}_")
            if checked:
                selected.append(ao["name"])
                for c in ao["conflicts"].split(","):
                    c = c.strip()
                    if c and c not in ("None", "nan"):
                        blocked.add(c)

        st.markdown("---")
        st.caption("Prices sourced from `pttoc_pricing.xlsx`")

    # ── Route to view ─────────────────────────────────────────────────────────
    if is_sales:
        sales_view(pkg, n_devices, duration, addons, selected)
    else:
        client_view(pkg, n_devices, duration, addons, selected)




# ─────────────────────────────────────────────────────────────────────────────
# pricing function
# ─────────────────────────────────────────────────────────────────────────────
def calculate_package_price(pkg, n_devices, duration, addons, selected, apply_addons=True):
    """
    Simplified pricing model:
    - Recurring base per package (you can later connect to Excel)
    - Add-ons applied only to selected package
    """

    # TEMP base pricing (replace later with Excel if needed)
    BASE_RECURRING = {
        "Essentials": 150,
        "Standard": 300,
        "Premium": 550,
    }

    SWAP = {
        "Essentials": 40,
        "Standard": 35,
        "Premium": 30,
    }

    APEX = {
        "Essentials": 600,
        "Standard": 700,
        "Premium": 2000,
    }

    recurring = BASE_RECURRING[pkg]
    swap = SWAP[pkg] * 1.3  # risk adjusted
    apex = APEX[pkg] / duration

    addon_cost = 0
    if apply_addons:
        for a in addons:
            if a["name"] in selected and a["type"] == "recurring":
                addon_cost += a["addon_price"]

    ppd = recurring + swap + apex + addon_cost

    fleet_year = ppd * n_devices
    total_contract = fleet_year * duration

    return {
        "ppd": ppd,
        "fleet_year": fleet_year,
        "total": total_contract,
    }

# ─────────────────────────────────────────────────────────────────────────────
# CLIENT VIEW
# ─────────────────────────────────────────────────────────────────────────────

def client_view(pkg, n_devices, duration, addons, selected):
    st.markdown(f"<h2 style='color:#1F3864'>📋 Your Configuration — {pkg}</h2>",
                unsafe_allow_html=True)

    # ── Upgrade alert (most important: show first) ────────────────────────────
    upgrade_alert(pkg, selected, addons)

    st.markdown("---")

    # ── Package feature table ─────────────────────────────────────────────────
    st.subheader("Package comparison")
    rows = []
    for feat, vals in FEATURES.items():
        row = {"Feature": feat}
        for p in PACKAGES:
            v = vals[p]
            row[p] = "✅" if v == "✔" else ("" if v == "–" else v)
        rows.append(row)

    df = pd.DataFrame(rows).set_index("Feature")

    def hl(col):
        return (
            ["background-color:#EEF4FF; font-weight:bold"] * len(col)
            if col.name == pkg
            else [""] * len(col)
        )

    st.dataframe(df.style.apply(hl, axis=0), use_container_width=True)

    st.markdown("---")
    
    st.subheader("💰 Price comparison (all packages)")
    
    comparison = []
    
    for p in PACKAGES:
        # Only apply add-ons to selected package
        apply_addons = (p == pkg)
    
        result = calculate_package_price(
            p, n_devices, duration, addons, selected, apply_addons
        )
    
        comparison.append({
            "Package": p + (" (Selected)" if p == pkg else ""),
            "€/device/year": f"€ {result['ppd']:,.0f}",
            "Fleet / year": f"€ {result['fleet_year']:,.0f}",
            "Total contract": f"€ {result['total']:,.0f}",
        })
    
    df_compare = pd.DataFrame(comparison).set_index("Package")
    
    def highlight_selected(row):
        if "(Selected)" in row.name:
            return ["background-color:#EEF4FF; font-weight:bold"] * len(row)
        return [""] * len(row)
    
    st.dataframe(df_compare.style.apply(highlight_selected, axis=1),
                 use_container_width=True)
    
    st.markdown("### 📊 Total contract value comparison")

    fig = go.Figure()
    
    for p in PACKAGES:
        apply_addons = (p == pkg)
        result = calculate_package_price(
            p, n_devices, duration, addons, selected, apply_addons
        )
    
        fig.add_trace(go.Bar(
            name=p,
            x=[p],
            y=[result["total"]],
            marker_color=PKG_COLOR[p],
            text=[f"€ {result['total']:,.0f}"],
            textposition="outside"
        ))
    
    fig.update_layout(
        plot_bgcolor="white",
        yaxis=dict(title="Total contract (€)", gridcolor="#eee"),
        height=350,
    )
    
    st.plotly_chart(fig, use_container_width=True)

    # ── Add-on summary ────────────────────────────────────────────────────────
    ao_list = [a for a in addons if a["name"] in selected]
    recurring = [a for a in ao_list if a["type"] == "recurring"]
    oneoffs   = [a for a in ao_list if a["type"] == "oneoff"]

    rec_per_dev  = sum(a["addon_price"] for a in recurring)
    rec_fleet_yr = rec_per_dev * n_devices
    oo_total     = sum(a["addon_price"] for a in oneoffs)

    if selected:
        st.subheader("Your add-ons")

        c1, c2, c3 = st.columns(3)
        c1.metric("Recurring / device / yr", f"€ {rec_per_dev:,.2f}" if rec_per_dev else "–")
        c2.metric("Recurring / fleet / yr",  f"€ {rec_fleet_yr:,.0f}" if rec_fleet_yr else "–",
                  f"{n_devices} devices")
        c3.metric("One-off total",           f"€ {oo_total:,.2f}" if oo_total else "–")

        # Detail table
        rows_ao = []
        for a in ao_list:
            rows_ao.append({
                "Add-on": a["name"],
                "Price":  (f"€ {a['addon_price']:,.2f} / yr"
                           if a["type"] == "recurring"
                           else f"€ {a['addon_price']:,.2f} one-off"),
                "Fleet / yr": (f"€ {a['addon_price']*n_devices:,.0f}"
                               if a["type"] == "recurring" else "–"),
                "Notes": a["notes"],
            })
        st.dataframe(pd.DataFrame(rows_ao).set_index("Add-on"), use_container_width=True)

        # Cumulative cost chart
        if rec_fleet_yr:
            st.markdown("**Cumulative add-on spend over contract**")
            durations = list(range(1, 6))
            fig = go.Figure(go.Bar(
                x=[f"{d} yr" for d in durations],
                y=[rec_fleet_yr * d + oo_total for d in durations],
                marker_color=PKG_COLOR[pkg],
                text=[f"€ {rec_fleet_yr*d+oo_total:,.0f}" for d in durations],
                textposition="outside",
            ))
            fig.add_vline(x=duration - 1, line_dash="dash", line_color="#888",
                          annotation_text=f"Your contract")
            fig.update_layout(
                plot_bgcolor="white", height=300,
                yaxis=dict(gridcolor="#eee", title="Cumulative cost (€)"),
                showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No add-ons selected. Use the sidebar to add features to your package.")


def upgrade_alert(pkg, selected, addons):
    """Show a prominent alert when selected add-ons are cheaper if you upgrade."""
    if pkg == "Premium" or not selected:
        return

    next_pkg = "Standard" if pkg == "Essentials" else "Premium"

    # Add-ons selected that are already included in the next package
    redundant = [
        a["name"] for a in addons
        if a["name"] in selected and INCLUDED_IN.get(a["name"]) == next_pkg
    ]

    num_selected = len(selected)

    if redundant:
        st.error(
            f"🔼 **Upgrade to {next_pkg} recommended**  \n"
            f"You've added: **{', '.join(redundant)}**  \n"
            f"These are already included in **{next_pkg}** at a lower bundled price. "
            f"You're paying the add-on premium (×1.3) for something that comes standard in the next tier."
        )
    elif num_selected >= 3:
        st.warning(
            f"💡 **Consider upgrading to {next_pkg}**  \n"
            f"You have {num_selected} add-ons selected. "
            f"At this level of customisation, the **{next_pkg}** package "
            f"may offer better overall value — several of these features are included natively."
        )


# ─────────────────────────────────────────────────────────────────────────────
# SALES VIEW
# ─────────────────────────────────────────────────────────────────────────────

def sales_view(pkg, n_devices, duration, addons, selected):
    st.markdown("<h2 style='color:#1F3864'>💼 Sales — Margin Overview</h2>",
                unsafe_allow_html=True)

    # ── Fleet P&L for selected add-ons ───────────────────────────────────────
    ao_selected = [a for a in addons if a["name"] in selected]

    if ao_selected:
        st.subheader(f"Selected add-ons — {n_devices} devices, {duration}-year contract")

        rows_pl = []
        for a in ao_selected:
            if a["type"] == "recurring":
                rev  = a["addon_price"] * n_devices * duration
                cost = a["cost"]        * n_devices * duration
            else:
                rev  = a["addon_price"]
                cost = a["cost"]
            rows_pl.append({
                "Add-on":    a["name"],
                "Type":      a["type"],
                "Revenue":   rev,
                "Cost":      cost,
                "Margin €":  rev - cost,
                "Margin %":  (rev - cost) / rev * 100 if rev else 0,
            })

        df_pl = pd.DataFrame(rows_pl)
        total_rev    = df_pl["Revenue"].sum()
        total_cost   = df_pl["Cost"].sum()
        total_margin = df_pl["Margin €"].sum()
        total_pct    = total_margin / total_rev * 100 if total_rev else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total revenue",   f"€ {total_rev:,.0f}")
        c2.metric("Total cost",      f"€ {total_cost:,.0f}")
        c3.metric("Total margin",    f"€ {total_margin:,.0f}")
        c4.metric("Gross margin %",  f"{total_pct:.1f}%",
                  delta_color="normal" if total_pct >= 30 else "inverse")

        # Format for display
        df_display = df_pl.copy()
        df_display["Revenue"]  = df_display["Revenue"].apply(lambda x: f"€ {x:,.2f}")
        df_display["Cost"]     = df_display["Cost"].apply(lambda x: f"€ {x:,.2f}")
        df_display["Margin €"] = df_display["Margin €"].apply(lambda x: f"€ {x:,.2f}")
        df_display["Margin %"] = df_display["Margin %"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(df_display.set_index("Add-on"), use_container_width=True)

        st.markdown("---")
    else:
        st.info("Select add-ons in the sidebar to see your P&L.")
        st.markdown("---")

    # ── Full margin table — all add-ons ──────────────────────────────────────
    st.subheader("All add-ons — margin at package price vs add-on price")
    st.caption(
        "**Package price** = bundled rate (cost × multiplier) — lower margin, but drives package upgrades.  \n"
        "**Add-on price** = client adds on top of existing package (package price × 1.3) — higher margin."
    )

    rows_all = []
    for a in addons:
        m_pkg = (a["pkg_price"] - a["cost"]) / a["pkg_price"] * 100 if a["pkg_price"] else 0
        m_ao  = (a["addon_price"] - a["cost"]) / a["addon_price"] * 100 if a["addon_price"] else 0
        rows_all.append({
            "Add-on":              a["name"],
            "Type":                a["type"],
            "Cost":                a["cost"],
            "Pkg price":           a["pkg_price"],
            "Add-on price":        a["addon_price"],
            "Margin @ pkg":        f"{m_pkg:.1f}%",
            "Margin @ add-on":     f"{m_ao:.1f}%",
            "Selected":            "✔" if a["name"] in selected else "",
        })

    df_all = pd.DataFrame(rows_all).set_index("Add-on")

    def hl_selected(row):
        return (
            ["background-color:#E2EFDA"] * len(row)
            if row["Selected"] == "✔"
            else [""] * len(row)
        )

    st.dataframe(
        df_all.style
              .apply(hl_selected, axis=1)
              .format({"Cost": "€ {:.2f}", "Pkg price": "€ {:.2f}", "Add-on price": "€ {:.2f}"}),
        use_container_width=True,
        height=420,
    )

    st.markdown("---")

    # ── Chart: margin % per add-on at both price points ──────────────────────
    st.subheader("Margin % — package price vs add-on price")
    names   = [a["name"] for a in addons]
    m_pkg   = [(a["pkg_price"]-a["cost"])/a["pkg_price"]*100 if a["pkg_price"] else 0 for a in addons]
    m_ao    = [(a["addon_price"]-a["cost"])/a["addon_price"]*100 if a["addon_price"] else 0 for a in addons]

    fig = go.Figure()
    fig.add_trace(go.Bar(name="Package price", x=names, y=m_pkg,
                         marker_color="#4472C4",
                         text=[f"{v:.1f}%" for v in m_pkg], textposition="outside"))
    fig.add_trace(go.Bar(name="Add-on price", x=names, y=m_ao,
                         marker_color="#2E7D32",
                         text=[f"{v:.1f}%" for v in m_ao], textposition="outside"))
    fig.add_hline(y=30, line_dash="dash", line_color="red",
                  annotation_text="30% target", annotation_position="top right")
    fig.update_layout(
        barmode="group",
        plot_bgcolor="white",
        yaxis=dict(title="Gross margin %", gridcolor="#eee"),
        xaxis_tickangle=-30,
        height=400,
        legend=dict(orientation="h", y=1.1),
    )
    st.plotly_chart(fig, use_container_width=True)


if __name__ == "__main__":
    main()