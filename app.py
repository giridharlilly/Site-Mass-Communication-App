"""
app.py — Site Mass Communication v4.0
Editable roles/classifications, document selection, HTML email preview
"""
import os, time, json
import dash
from dash import dcc, html, Input, Output, State, callback, ctx, ALL, no_update
import dash_bootstrap_components as dbc
import pandas as pd
from datetime import date
from concurrent.futures import ThreadPoolExecutor

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from ad_access import enforce_access, get_current_user, get_user_display_name
from db_connection import read_table

app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME],
    suppress_callback_exceptions=True, title="Site Mass Communication")
server = app.server
enforce_access(server)

# Theme
P = "#1a237e"; PL = "#283593"; ACC = "#3B82F6"; SUC = "#10B981"; WRN = "#F59E0B"; DNG = "#EF4444"
BG = "#F5F7FA"; CBG = "#FFFFFF"; TD = "#1E293B"; TM = "#64748B"; TL = "#94A3B8"; BD = "#E2E8F0"

# Cache
_cache = {}; _cache_ts = {}

def gc(name, force=False):
    now = time.time()
    if not force and name in _cache and (now - _cache_ts.get(name, 0)) < 300:
        return _cache[name].copy()
    try:
        df = read_table(name); _cache[name] = df; _cache_ts[name] = now; return df.copy()
    except Exception as e:
        print(f"Error reading {name}: {e}"); return pd.DataFrame()

for t in ["sc_lookup", "study_country_site_lookup"]:
    try: _cache[t] = read_table(t); _cache_ts[t] = time.time(); print(f"DEBUG: {t} = {len(_cache[t])} rows")
    except Exception as e: print(f"DEBUG: {t} FAILED = {e}")

# Get all unique roles from personnel tables for dropdown options
def get_all_roles():
    roles = set()
    try:
        df = gc("study_sponsor_personnel_assignment")
        if not df.empty and "Study_Team_Role" in df.columns:
            roles.update(df["Study_Team_Role"].dropna().unique().tolist())
    except: pass
    try:
        df = gc("country_sponsor_personnel_assignment")
        if not df.empty and "Study_Team_Role" in df.columns:
            roles.update(df["Study_Team_Role"].dropna().unique().tolist())
    except: pass
    try:
        df = gc("study_site_sponsor_personnel_combined")
        if not df.empty and "Role" in df.columns:
            roles.update(df["Role"].dropna().unique().tolist())
    except: pass
    return sorted([r for r in roles if r])

def get_all_classifications():
    try:
        df = gc("documents")
        if not df.empty and "Classification" in df.columns:
            return sorted(df["Classification"].dropna().unique().tolist())
    except: pass
    return []

def _card(title, icon, color, children, flex=None):
    style = {"background": CBG, "borderRadius": "8px", "padding": "10px 14px", "boxShadow": "0 1px 3px rgba(0,0,0,0.05)"}
    if flex: style["flex"] = flex
    header = []
    if title:
        header = [html.Div([html.I(className=f"{icon} me-2", style={"color": color, "fontSize": "11px"}),
            html.Span(title, style={"fontWeight": "700", "fontSize": "11px", "color": color, "letterSpacing": "0.5px"})],
            className="d-flex align-items-center mb-2")]
    return html.Div(header + children, style=style)

# ═══════ LAYOUT ═══════
app.layout = html.Div([
    dcc.Store(id="s-page", data="home"),
    dcc.Store(id="s-tpl", data=None),
    dcc.Store(id="s-to", data=[]),
    dcc.Store(id="s-bcc", data=[]),
    dcc.Store(id="s-docs", data=[]),

    # HEADER
    html.Div([
        html.Div([
            html.I(className="fas fa-home", id="btn-home",
                style={"fontSize": "18px", "cursor": "pointer", "color": "rgba(255,255,255,0.8)", "marginRight": "16px"}),
            html.Span("Site Mass Communication", style={"fontWeight": "700", "fontSize": "18px", "color": "white"}),
            html.Span(id="breadcrumb", style={"color": "rgba(255,255,255,0.6)", "fontSize": "12px", "marginLeft": "16px"}),
        ], className="d-flex align-items-center"),
        html.Div([
            html.Span("v4.0", style={"color": "rgba(255,255,255,0.4)", "fontSize": "10px", "marginRight": "16px"}),
            html.Div(id="header-user", style={"background": "rgba(255,255,255,0.1)", "borderRadius": "20px",
                "padding": "4px 14px", "fontSize": "13px", "color": "white"}),
        ], className="d-flex align-items-center"),
    ], className="d-flex justify-content-between align-items-center px-4",
        style={"background": f"linear-gradient(135deg, {P}, {PL})", "height": "56px", "boxShadow": "0 2px 8px rgba(0,0,0,0.15)"}),

    # HOME PAGE
    html.Div(id="home-page", children=[
        html.Div([
            html.H4("Select a Communication Template", style={"fontWeight": "600", "color": TD, "marginBottom": "4px"}),
            html.P("Choose a template to compose and send mass communications",
                style={"color": TM, "fontSize": "13px", "marginBottom": "24px"}),
            html.Div(id="template-grid"),
        ], style={"padding": "24px 32px"}),
    ]),

    # COMPOSE PAGE
    html.Div(id="compose-page", style={"display": "none", "padding": "12px 16px"}, children=[
        html.Div([
            # ═══ LEFT: Email (50%) ═══
            html.Div([
                html.Div([
                    html.Div([html.I(className="fas fa-pen-to-square me-2", style={"color": P}),
                        html.Span("Email Composition", style={"fontWeight": "700", "fontSize": "14px", "color": P})],
                        className="d-flex align-items-center mb-3"),

                    html.Label("Subject", style={"fontSize": "11px", "fontWeight": "600", "color": TM}),
                    dbc.Input(id="inp-subject", value="", size="sm",
                        style={"fontSize": "13px", "marginBottom": "12px", "background": "#FAFBFC", "border": f"1px solid {BD}", "borderRadius": "6px"}),

                    html.Label("Email Body (click to edit)", style={"fontSize": "11px", "fontWeight": "600", "color": TM, "marginBottom": "4px"}),

                    html.Iframe(id="email-preview", srcDoc="", style={"width": "100%", "height": "420px",
                        "border": f"1px solid {BD}", "borderRadius": "6px", "background": "white"}),
                    # Hidden store for body HTML (synced via JS in iframe)
                    dcc.Store(id="inp-body-store", data=""),

                    html.Div([html.Span("Placeholders: ", style={"fontWeight": "600", "fontSize": "10px"}),
                        html.Span("{{STUDY_ALIAS}} {{COUNTRY}} {{SITE}} {{DATE}} {{DOC_IDS}}",
                            style={"fontSize": "10px", "fontFamily": "monospace"})],
                        style={"color": TL, "marginTop": "6px"}),
                ], style={"background": CBG, "borderRadius": "10px", "padding": "16px", "boxShadow": "0 1px 4px rgba(0,0,0,0.06)"}),
            ], style={"flex": "1", "marginRight": "12px"}),

            # ═══ RIGHT: Controls (50%) ═══
            html.Div([
                # FILTERS
                _card("FILTERS", "fas fa-filter", P, [
                    dbc.Row([
                        dbc.Col([html.Label("Study Alias", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-study", placeholder="Select study...", clearable=True, style={"fontSize": "12px"})], md=4),
                        dbc.Col([html.Label("Country", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-country", placeholder="Select...", clearable=True, disabled=True, style={"fontSize": "12px"})], md=4),
                        dbc.Col([html.Label("Site", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-site", placeholder="Select...", clearable=True, disabled=True, style={"fontSize": "12px"})], md=4),
                    ]),
                ]),

                # ROLE & CLASSIFICATION FILTERS
                _card("ROLE & CLASSIFICATION FILTERS", "fas fa-sliders", P, [
                    dbc.Row([
                        dbc.Col([html.Label("Lilly Groups", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-lilly", multi=True, placeholder="Select roles...",
                                style={"fontSize": "11px"})], md=6),
                        dbc.Col([html.Label("Non Lilly Roles", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-nonlilly", multi=True, placeholder="Select roles...",
                                style={"fontSize": "11px"})], md=6),
                    ], className="mb-2"),
                    dbc.Row([
                        dbc.Col([html.Label("Classifications", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-class", multi=True, placeholder="Select classifications...",
                                style={"fontSize": "11px"})], md=12),
                    ]),
                    dbc.Button([html.I(className="fas fa-sync me-1"), "Refresh Recipients & Docs"], id="btn-refresh-roles",
                        size="sm", className="mt-2 w-100", style={"background": P, "border": "none", "fontSize": "11px"}),
                ]),

                # TO RECIPIENTS
                _card("", "fas fa-users", P, [
                    html.Div(id="to-header", style={"fontWeight": "700", "fontSize": "12px", "color": P, "marginBottom": "6px"}),
                    dcc.Textarea(id="display-to", value="", readOnly=True,
                        style={"width": "100%", "height": "55px", "fontSize": "10px", "background": "#FAFBFC",
                            "border": f"1px solid {BD}", "borderRadius": "6px", "padding": "6px", "resize": "none"}),
                    html.Div([
                        dbc.Input(id="inp-add-to", placeholder="user@lilly.com; user2@lilly.com",
                            size="sm", style={"fontSize": "11px", "flex": "1", "borderRadius": "6px"}),
                        dbc.Button([html.I(className="fas fa-plus me-1"), "Add"], id="btn-add-to", size="sm",
                            style={"background": P, "border": "none", "fontSize": "11px", "marginLeft": "6px"}),
                    ], className="d-flex align-items-center mt-1"),
                ]),

                # BCC
                _card("", "fas fa-eye-slash", ACC, [
                    html.Div(id="bcc-header", style={"fontWeight": "700", "fontSize": "12px", "color": ACC, "marginBottom": "4px"}),
                    html.Div("Not visible to other recipients", style={"color": TL, "fontSize": "9px", "marginBottom": "4px"}),
                    html.Div([
                        dbc.Input(id="inp-add-bcc", placeholder="external@partner.com",
                            size="sm", style={"fontSize": "11px", "flex": "1", "background": "#F0F7FF", "borderColor": "#B0D0F0"}),
                        dbc.Button([html.I(className="fas fa-plus me-1"), "BCC"], id="btn-add-bcc", size="sm",
                            style={"background": ACC, "border": "none", "fontSize": "11px", "marginLeft": "6px"}),
                    ], className="d-flex align-items-center"),
                    html.Div(id="bcc-display", style={"fontSize": "9px", "color": ACC, "marginTop": "4px"}),
                ]),

                # DOCUMENTS (with checkboxes)
                _card("", "fas fa-file-alt", P, [
                    html.Div(id="doc-header", style={"fontWeight": "700", "fontSize": "12px", "color": P, "marginBottom": "2px"}),
                    html.Div("Check documents to include VV-TMF IDs in email body",
                        style={"color": TL, "fontSize": "9px", "marginBottom": "6px"}),
                    dcc.Store(id="s-doc-data", data=[]),  # Store doc info for checkbox lookup
                    html.Div(id="doc-list", style={"maxHeight": "160px", "overflowY": "auto",
                        "borderRadius": "6px", "border": f"1px solid {BD}", "background": "#FAFBFC"}),
                ], flex="1"),

                # SEND
                _card("Send & Save", "fas fa-paper-plane", P, [
                    html.Div(id="send-summary", style={"fontSize": "11px", "color": TM, "marginBottom": "6px"}),
                    dbc.Button([html.I(className="fas fa-paper-plane me-2"), "Send Email & Save Communication"],
                        id="btn-send", size="sm", disabled=True,
                        style={"background": f"linear-gradient(135deg, {P}, {PL})", "border": "none",
                            "width": "100%", "height": "40px", "fontSize": "12px", "fontWeight": "600", "borderRadius": "8px"}),
                    html.Div(id="send-status", style={"marginTop": "6px"}),
                ]),
            ], style={"flex": "1", "display": "flex", "flexDirection": "column", "gap": "8px"}),
        ], style={"display": "flex", "alignItems": "stretch"}),
    ]),
], style={"fontFamily": "'Segoe UI', -apple-system, sans-serif", "minHeight": "100vh", "background": BG})


# ═══════════════════════════════════════════════════════════════════════
#  CALLBACKS
# ═══════════════════════════════════════════════════════════════════════

@callback(Output("header-user", "children"), Input("s-page", "data"))
def show_user(_):
    uid = get_current_user(); name = get_user_display_name(uid)
    return [html.I(className="fas fa-user me-2", style={"fontSize": "11px"}), html.Span(name or uid)]

@callback(Output("breadcrumb", "children"), [Input("s-page", "data"), Input("s-tpl", "data")])
def bc(page, tpl):
    return f"Home  ›  {tpl.get('Title', '')}" if page == "compose" and tpl else ""

@callback([Output("home-page", "style"), Output("compose-page", "style")], Input("s-page", "data"))
def toggle(page):
    if page == "compose": return {"display": "none"}, {"display": "block", "padding": "12px 16px"}
    return {"display": "block"}, {"display": "none", "padding": "12px 16px"}

@callback([Output("s-page", "data", allow_duplicate=True), Output("s-tpl", "data", allow_duplicate=True),
    Output("s-to", "data", allow_duplicate=True), Output("s-bcc", "data", allow_duplicate=True),
    Output("s-docs", "data", allow_duplicate=True),
    Output("dd-study", "value"), Output("dd-country", "value"), Output("dd-site", "value"),
    Output("dd-lilly", "value"), Output("dd-nonlilly", "value"), Output("dd-class", "value")],
    Input("btn-home", "n_clicks"), prevent_initial_call=True)
def go_home(n):
    return "home", None, [], [], [], None, None, None, [], [], []

@callback(Output("template-grid", "children"), Input("s-page", "data"))
def grid(page):
    df = gc("sc_lookup")
    if df.empty: return dbc.Alert("No templates found.", color="warning")
    cards = []
    for _, r in df.iterrows():
        title, tname, tid = str(r.get("Title","")), str(r.get("Template_Name","")), str(r.get("ID",""))
        if not title or tid == "nan": continue
        cards.append(dbc.Col(html.Button(html.Div([
            html.I(className="fas fa-envelope-open-text", style={"fontSize": "18px", "marginBottom": "6px", "opacity": "0.9"}),
            html.Div(title, style={"fontWeight": "600", "fontSize": "11px", "lineHeight": "1.3"}),
            html.Div(tname, style={"fontSize": "8px", "opacity": "0.7", "marginTop": "3px", "maxHeight": "20px", "overflow": "hidden"}),
        ], style={"textAlign": "center", "padding": "8px 4px"}),
            id={"type": "tpl-btn", "index": tid},
            style={"background": f"linear-gradient(135deg, {P}, {PL})", "color": "white", "border": "none",
                "borderRadius": "10px", "width": "100%", "height": "90px", "cursor": "pointer",
                "boxShadow": "0 2px 6px rgba(26,35,126,0.2)"}),
            xs=6, sm=4, md=3, lg=2, className="mb-3"))
    return dbc.Row(cards)

# Template click → populate compose page
@callback([Output("s-page", "data", allow_duplicate=True), Output("s-tpl", "data", allow_duplicate=True),
    Output("s-to", "data", allow_duplicate=True), Output("s-bcc", "data", allow_duplicate=True), Output("s-docs", "data", allow_duplicate=True),
    Output("inp-subject", "value"), Output("inp-body-store", "data"), Output("email-preview", "srcDoc"),
    Output("dd-lilly", "options"), Output("dd-lilly", "value", allow_duplicate=True),
    Output("dd-nonlilly", "options"), Output("dd-nonlilly", "value", allow_duplicate=True),
    Output("dd-class", "options"), Output("dd-class", "value", allow_duplicate=True)],
    Input({"type": "tpl-btn", "index": ALL}, "n_clicks"), prevent_initial_call=True)
def select_tpl(clicks):
    if not ctx.triggered_id or not any(c for c in clicks if c): return [no_update] * 14
    idx = str(ctx.triggered_id["index"])
    df = gc("sc_lookup")
    row = df[df["ID"].astype(str) == idx]
    if row.empty: return [no_update] * 14
    r = row.iloc[0]
    tpl = {col: str(r.get(col, "") or "") for col in
        ["ID", "Title", "Template_Name", "EmailSubject", "EmailBody", "Classifications", "Lilly_Groups", "Non_Lilly_Roles", "DocURL"]}

    body = tpl["EmailBody"]; subject = tpl["EmailSubject"]
    if subject == "nan": subject = ""
    preview = _wrap_html(body)

    lilly_vals = list(dict.fromkeys([g.strip() for g in tpl["Lilly_Groups"].replace("\n", ";").split(";") if g.strip()]))
    nonlilly_vals = list(dict.fromkeys([g.strip() for g in tpl["Non_Lilly_Roles"].replace("\n", ";").split(";") if g.strip()]))
    class_vals = list(dict.fromkeys([c.strip() for c in tpl["Classifications"].replace("\n", ";").split(";") if c.strip()]))

    all_roles = get_all_roles()
    all_class = get_all_classifications()
    lilly_all = list(dict.fromkeys(lilly_vals + [r for r in all_roles if r not in lilly_vals]))
    nonlilly_all = list(dict.fromkeys(nonlilly_vals + [r for r in all_roles if r not in nonlilly_vals]))
    class_all = list(dict.fromkeys(class_vals + [c for c in all_class if c not in class_vals]))

    lilly_opts = [{"label": r, "value": r} for r in lilly_all]
    nonlilly_opts = [{"label": r, "value": r} for r in nonlilly_all]
    class_opts = [{"label": c, "value": c} for c in class_all]

    return ("compose", tpl, [], [], [], subject, body, preview,
        lilly_opts, lilly_vals, nonlilly_opts, nonlilly_vals, class_opts, class_vals)


def _wrap_html(body):
    return f"""<html><head><style>
body{{font-family:'Segoe UI',sans-serif;font-size:13px;padding:12px;color:#1E293B;line-height:1.6;outline:none;}}
body:focus{{outline:none;}}
</style></head><body contenteditable="true">{body}</body></html>"""



# ═══════ FILTER CASCADE ═══════

@callback(Output("dd-study", "options"), Input("s-page", "data"))
def load_studies(page):
    df = gc("study_country_site_lookup")
    if df.empty or "Study_Alias" not in df.columns: return []
    return [{"label": s, "value": s} for s in sorted(df["Study_Alias"].dropna().unique().tolist()) if s]

@callback([Output("dd-country", "options"), Output("dd-country", "disabled"),
    Output("dd-country", "value", allow_duplicate=True), Output("dd-site", "value", allow_duplicate=True), Output("dd-site", "disabled")],
    Input("dd-study", "value"), prevent_initial_call=True)
def cascade_c(study):
    if not study: return [], True, None, None, True
    df = gc("study_country_site_lookup")
    f = df[df["Study_Alias"] == study]
    c = sorted(f["Country_Name"].dropna().unique().tolist()) if "Country_Name" in f.columns else []
    return [{"label": x, "value": x} for x in c if x], False, None, None, True

@callback([Output("dd-site", "options"), Output("dd-site", "disabled", allow_duplicate=True)],
    Input("dd-country", "value"), State("dd-study", "value"), prevent_initial_call=True)
def cascade_s(country, study):
    if not country or not study: return [], True
    df = gc("study_country_site_lookup")
    f = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country)]
    s = sorted(f["Site"].dropna().unique().tolist()) if "Site" in f.columns else []
    return [{"label": x, "value": x} for x in s if x], False


# ═══════ AUTO-POPULATE ON FILTER/ROLE CHANGE ═══════

@callback([Output("s-bcc", "data", allow_duplicate=True),
    Output("inp-subject", "value", allow_duplicate=True),
    Output("inp-body-store", "data", allow_duplicate=True),
    Output("email-preview", "srcDoc", allow_duplicate=True),
    Output("doc-list", "children"), Output("doc-header", "children"),
    Output("s-doc-data", "data")],
    [Input("dd-study", "value"), Input("dd-country", "value"), Input("dd-site", "value"),
     Input("btn-refresh-roles", "n_clicks")],
    [State("s-tpl", "data"), State("dd-lilly", "value"), State("dd-nonlilly", "value"),
     State("dd-class", "value")],
    prevent_initial_call=True)
def on_filter(study, country, site, refresh_n, tpl, lilly_sel, nonlilly_sel, class_sel):
    if not tpl: return no_update, no_update, no_update, no_update, no_update, no_update, no_update

    body = tpl.get("EmailBody", "")
    body = body.replace("{{STUDY_ALIAS}}", study or "").replace("{{COUNTRY}}", country or "")
    body = body.replace("{{SITE}}", site or "").replace("{{DATE}}", date.today().strftime("%d %B %Y"))
    subject = tpl.get("EmailSubject", "").replace("{{STUDY_ALIAS}}", study or "")
    if subject == "nan": subject = ""

    lg = lilly_sel or []
    nlr = nonlilly_sel or []
    cls = class_sel or []

    # Get recipients for current filter level (replaces previous, not merge)
    bcc_list = _auto_recips(study, country, site, lg, nlr)

    docs_html, doc_count, doc_data = _render_docs(study, country, site, cls)

    doc_header = [html.I(className="fas fa-file-alt me-2", style={"color": P, "fontSize": "11px"}),
        f"eTMF & Non-TMF Documents ({doc_count} found)"]

    return bcc_list, subject, body, _wrap_html(body), docs_html, doc_header, doc_data


def _auto_recips(study, country, site, lg, nlr):
    """Get recipients based on CURRENT filter level only:
    - Site selected → query Site personnel table only
    - Country selected (no site) → query Country personnel table only
    - Study only → query Study personnel table only
    """
    emails = set()
    if not study: return sorted(list(emails))

    if site and country:
        # Site level — use Lilly Groups + Non Lilly Roles from site table
        all_r = list(set(lg + nlr))
        if all_r:
            try:
                df = gc("study_site_sponsor_personnel_combined")
                if not df.empty:
                    m = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country) & (df["Site"] == site) & (df["Role"].isin(all_r)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                    emails.update(m["Email_Address"].str.strip().tolist())
            except: pass

    elif country:
        # Country level — use Lilly Groups from country table
        if lg:
            try:
                df = gc("country_sponsor_personnel_assignment")
                if not df.empty:
                    m = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country) & (df["Study_Team_Role"].isin(lg)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                    emails.update(m["Email_Address"].str.strip().tolist())
            except: pass

    else:
        # Study level only — use Lilly Groups from study table
        if lg:
            try:
                df = gc("study_sponsor_personnel_assignment")
                if not df.empty:
                    m = df[(df["Study_Alias"] == study) & (df["Study_Team_Role"].isin(lg)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                    emails.update(m["Email_Address"].str.strip().tolist())
            except: pass

    return sorted([e for e in emails if e])


def _render_docs(study, country, site, class_filter):
    if not study:
        return html.Div("Select a study to load documents.", style={"color": TL, "fontSize": "11px", "padding": "10px", "textAlign": "center"}), 0, []
    try:
        df = gc("documents")
        if df.empty:
            return html.Div("No documents.", style={"color": TL, "fontSize": "11px", "padding": "10px", "textAlign": "center"}), 0, []

        mask = df["Study_Alias"] == study
        s_lvl = mask & (df["Site"].isna() | (df["Site"] == "")) & (df["Country_Code"].isna() | (df["Country_Code"] == ""))
        c_lvl = pd.Series(False, index=df.index)
        if country: c_lvl = mask & (df["Country_Code"] == country) & (df["Site"].isna() | (df["Site"] == ""))
        st_lvl = pd.Series(False, index=df.index)
        if site: st_lvl = mask & (df["Site"] == site)
        filtered = df[s_lvl | c_lvl | st_lvl]

        if class_filter:
            filtered = filtered[filtered["Classification"].isin(class_filter)]

        doc_count = len(filtered)
        filtered = filtered.head(150)

        if filtered.empty:
            return html.Div("No documents found.", style={"color": TL, "fontSize": "11px", "padding": "10px", "textAlign": "center"}), 0, []

        items = []
        doc_data = []  # Store doc info for checkbox callback
        for i, (_, doc) in enumerate(filtered.iterrows()):
            dn = str(doc.get("Document_Number", "")); dname = str(doc.get("Document_Name", ""))
            dl = str(doc.get("DocLink", "")); cl = str(doc.get("Classification", ""))
            df_dt = str(doc.get("Date_Document_Finalized", ""))[:10]
            sv, cv = doc.get("Site", ""), doc.get("Country_Code", "")
            if not sv and not cv: lbl, lc = "Study", P
            elif cv and not sv: lbl, lc = "Country", SUC
            else: lbl, lc = "Site", WRN

            doc_data.append({"num": dn, "name": dname, "link": dl})

            items.append(html.Div([
                dbc.Checkbox(id={"type": "doc-chk", "index": i}, value=False,
                    className="me-2", style={"marginTop": "2px"}),
                html.Div([
                    html.A(f"{dn} - {dname}", href=dl, target="_blank",
                        style={"color": P, "fontSize": "10px", "fontWeight": "600", "textDecoration": "none"}),
                    html.Div([html.Span(f"{cl} | {df_dt}", style={"color": TL, "fontSize": "8px"}),
                        html.Span(lbl, style={"fontSize": "8px", "fontWeight": "600", "color": lc,
                            "marginLeft": "6px", "background": f"{lc}15", "padding": "0 4px", "borderRadius": "3px"})]),
                ], style={"flex": "1"}),
            ], className="d-flex", style={"padding": "4px 8px", "borderBottom": "1px solid #F0F0F0"}))

        return html.Div(items), doc_count, doc_data
    except Exception as e:
        return html.Div(f"Error: {e}", style={"color": DNG, "fontSize": "11px", "padding": "10px"}), 0, []


# ═══════ DOC SELECTION → INSERT INTO EMAIL BODY ═══════

@callback([Output("s-docs", "data", allow_duplicate=True),
    Output("inp-body-store", "data", allow_duplicate=True),
    Output("email-preview", "srcDoc", allow_duplicate=True)],
    Input({"type": "doc-chk", "index": ALL}, "value"),
    [State("s-doc-data", "data"), State("inp-body-store", "data")],
    prevent_initial_call=True)
def on_doc_select(checks, doc_data, current_body):
    if not checks or not doc_data:
        return no_update, no_update, no_update

    selected = []
    for i, checked in enumerate(checks):
        if checked and i < len(doc_data):
            d = doc_data[i]
            selected.append(f"{d['num']} - {d['name']}")

    # Build bullet list HTML
    if selected:
        bullet_html = "<ul>" + "".join(
            f"<li><strong>{d.split(' - ')[0]}</strong> – {' - '.join(d.split(' - ')[1:])}</li>"
            for d in selected
        ) + "</ul>"
    else:
        bullet_html = ""

    body = current_body or ""
    # Remove any previous doc list
    if "<!-- DOC_LIST_START -->" in body:
        start = body.index("<!-- DOC_LIST_START -->")
        end = body.index("<!-- DOC_LIST_END -->") + len("<!-- DOC_LIST_END -->") if "<!-- DOC_LIST_END -->" in body else len(body)
        body = body[:start] + body[end:]

    # Replace placeholder or append
    if "{{DOC_IDS}}" in body:
        body = body.replace("{{DOC_IDS}}", bullet_html)
    elif bullet_html:
        body = body + f"\n<!-- DOC_LIST_START -->\n<h4>Selected Documents:</h4>\n{bullet_html}\n<!-- DOC_LIST_END -->"
    
    return selected, body, _wrap_html(body)


# ═══════ UPDATE DISPLAYS ═══════

@callback([Output("display-to", "value"), Output("to-header", "children"), Output("btn-send", "disabled"), Output("send-summary", "children")],
    [Input("s-to", "data"), Input("s-bcc", "data"), Input("s-docs", "data")])
def upd_to(to_r, bcc_r, docs):
    to_r = to_r or []; bcc_r = bcc_r or []; docs = docs or []
    # Enable send if either TO or BCC has recipients
    has_recipients = len(to_r) > 0 or len(bcc_r) > 0
    return ("; ".join(to_r),
        [html.I(className="fas fa-users me-2", style={"color": P, "fontSize": "11px"}), f"TO Recipients ({len(to_r)})"],
        not has_recipients,
        f"TO: {len(to_r)}  |  BCC: {len(bcc_r)}  |  Docs: {len(docs)}")

@callback([Output("bcc-display", "children"), Output("bcc-header", "children")], Input("s-bcc", "data"))
def upd_bcc(bcc_r):
    bcc_r = bcc_r or []
    display = dcc.Textarea(value="; ".join(bcc_r), readOnly=True,
        style={"width": "100%", "height": "60px", "fontSize": "10px", "background": "#F0F7FF",
            "border": f"1px solid #B0D0F0", "borderRadius": "6px", "padding": "6px", "resize": "none",
            "color": ACC}) if bcc_r else ""
    return (display,
        [html.I(className="fas fa-eye-slash me-2", style={"color": ACC, "fontSize": "11px"}), f"BCC — Auto + External ({len(bcc_r)})"])


# ═══════ ADD TO / BCC ═══════

@callback([Output("s-to", "data", allow_duplicate=True), Output("inp-add-to", "value")],
    Input("btn-add-to", "n_clicks"), [State("inp-add-to", "value"), State("s-to", "data")], prevent_initial_call=True)
def add_to(n, text, cur):
    if not text: return no_update, no_update
    cur = cur or []; new = [e.strip() for e in text.replace(",", ";").split(";") if e.strip()]
    ex = set(cur)
    for e in new:
        if e not in ex: cur.append(e); ex.add(e)
    return cur, ""

@callback([Output("s-bcc", "data", allow_duplicate=True), Output("inp-add-bcc", "value")],
    Input("btn-add-bcc", "n_clicks"), [State("inp-add-bcc", "value"), State("s-bcc", "data")], prevent_initial_call=True)
def add_bcc(n, text, cur):
    if not text: return no_update, no_update
    cur = cur or []; new = [e.strip() for e in text.replace(",", ";").split(";") if e.strip()]
    ex = set(cur)
    for e in new:
        if e not in ex: cur.append(e); ex.add(e)
    return cur, ""


# ═══════ SEND (placeholder) ═══════

@callback(Output("send-status", "children"), Input("btn-send", "n_clicks"),
    [State("s-to", "data"), State("s-bcc", "data"), State("s-docs", "data")], prevent_initial_call=True)
def send(n, to_r, bcc_r, docs):
    return dbc.Alert([html.I(className="fas fa-info-circle me-2"),
        f"Ready to send to {len(to_r or [])} TO and {len(bcc_r or [])} BCC. {len(docs or [])} docs selected. ",
        "Integrate email service to enable sending."], color="info", dismissable=True, style={"fontSize": "10px"})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8050)
