"""
app.py — Site Mass Communication v3.0
Dash app for Posit Connect
Fixed: Persistent layout (no re-render), HTML email body, 50/50 split
"""

import os, json, time
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

# ═══════ THEME ═══════
P = "#1a237e"
PL = "#283593"
ACC = "#3B82F6"
SUC = "#10B981"
WRN = "#F59E0B"
DNG = "#EF4444"
BG = "#F5F7FA"
CBG = "#FFFFFF"
TD = "#1E293B"
TM = "#64748B"
TL = "#94A3B8"
BD = "#E2E8F0"

# ═══════ CACHE ═══════
_cache = {}
_cache_ts = {}

def get_cached(name, force=False):
    now = time.time()
    if not force and name in _cache and (now - _cache_ts.get(name, 0)) < 300:
        return _cache[name].copy()
    try:
        df = read_table(name)
        _cache[name] = df
        _cache_ts[name] = now
        return df.copy()
    except Exception as e:
        print(f"Error reading {name}: {e}")
        return pd.DataFrame()

# Preload
for t in ["sc_lookup", "study_country_site_lookup"]:
    try:
        _cache[t] = read_table(t)
        _cache_ts[t] = time.time()
        print(f"DEBUG: {t} loaded = {len(_cache[t])} rows")
    except Exception as e:
        print(f"DEBUG: {t} FAILED = {e}")

# ═══════ LAYOUT (persistent — never re-rendered) ═══════
app.layout = html.Div([
    dcc.Store(id="s-page", data="home"),
    dcc.Store(id="s-tpl", data=None),
    dcc.Store(id="s-to", data=[]),
    dcc.Store(id="s-bcc", data=[]),

    # HEADER
    html.Div([
        html.Div([
            html.I(className="fas fa-home", id="btn-home",
                style={"fontSize": "18px", "cursor": "pointer", "color": "rgba(255,255,255,0.8)", "marginRight": "16px"}),
            html.Span("Site Mass Communication", style={"fontWeight": "700", "fontSize": "18px", "color": "white"}),
            html.Span(id="breadcrumb", style={"color": "rgba(255,255,255,0.6)", "fontSize": "12px", "marginLeft": "16px"}),
        ], className="d-flex align-items-center"),
        html.Div([
            html.Span("v3.0", style={"color": "rgba(255,255,255,0.4)", "fontSize": "10px", "marginRight": "16px"}),
            html.Div(id="header-user", style={"background": "rgba(255,255,255,0.1)", "borderRadius": "20px",
                "padding": "4px 14px", "fontSize": "13px", "color": "white"}),
        ], className="d-flex align-items-center"),
    ], className="d-flex justify-content-between align-items-center px-4",
        style={"background": f"linear-gradient(135deg, {P}, {PL})", "height": "56px", "boxShadow": "0 2px 8px rgba(0,0,0,0.15)"}),

    # ═══════ HOME PAGE ═══════
    html.Div(id="home-page", children=[
        html.Div([
            html.H4("Select a Communication Template", style={"fontWeight": "600", "color": TD, "marginBottom": "4px"}),
            html.P("Choose a template to compose and send mass communications",
                style={"color": TM, "fontSize": "13px", "marginBottom": "24px"}),
            html.Div(id="template-grid"),
        ], style={"padding": "24px 32px"}),
    ]),

    # ═══════ COMPOSE PAGE ═══════
    html.Div(id="compose-page", style={"display": "none"}, children=[
        html.Div([
            # LEFT: Email (50%)
            html.Div([
                html.Div([
                    html.Div([html.I(className="fas fa-pen-to-square me-2", style={"color": P}),
                        html.Span("Email Composition", style={"fontWeight": "700", "fontSize": "14px", "color": P})],
                        className="d-flex align-items-center mb-3"),

                    html.Label("Subject", style={"fontSize": "11px", "fontWeight": "600", "color": TM}),
                    dbc.Input(id="inp-subject", value="", size="sm",
                        style={"fontSize": "13px", "marginBottom": "14px", "background": "#FAFBFC",
                            "border": f"1px solid {BD}", "borderRadius": "6px"}),

                    html.Div([
                        html.Label("Email Body", style={"fontSize": "11px", "fontWeight": "600", "color": TM}),
                        dbc.Button("Edit Source", id="btn-toggle-edit", size="sm", outline=True, color="secondary",
                            style={"fontSize": "10px", "marginLeft": "auto"}),
                    ], className="d-flex align-items-center mb-1"),

                    # HTML preview (default view)
                    html.Iframe(id="email-preview", style={"width": "100%", "height": "450px",
                        "border": f"1px solid {BD}", "borderRadius": "6px", "background": "white"}),

                    # Source editor (hidden by default)
                    dcc.Textarea(id="inp-body", value="", style={"width": "100%", "height": "450px",
                        "fontSize": "12px", "border": f"1px solid {BD}", "borderRadius": "6px",
                        "padding": "12px", "fontFamily": "monospace", "resize": "vertical",
                        "background": "#FAFBFC", "display": "none"}),

                    html.Div([html.Span("Placeholders: ", style={"fontWeight": "600", "fontSize": "10px"}),
                        html.Span("{{STUDY_ALIAS}}  {{COUNTRY}}  {{SITE}}  {{DATE}}  {{DOC_IDS}}",
                            style={"fontSize": "10px", "fontFamily": "monospace"})],
                        style={"color": TL, "marginTop": "8px"}),

                ], style={"background": CBG, "borderRadius": "10px", "padding": "20px",
                    "boxShadow": "0 1px 4px rgba(0,0,0,0.06)"}),
            ], style={"flex": "1", "marginRight": "16px"}),

            # RIGHT: Filters + Recipients + Docs + Send (50%)
            html.Div([
                # FILTERS
                html.Div([
                    html.Div([html.I(className="fas fa-filter me-2", style={"color": P, "fontSize": "12px"}),
                        html.Span("FILTERS", style={"fontWeight": "700", "fontSize": "11px", "color": P, "letterSpacing": "1px"})],
                        className="d-flex align-items-center mb-2"),
                    dbc.Row([
                        dbc.Col([html.Label("Study Alias", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-study", placeholder="Select study...", clearable=True, style={"fontSize": "12px"})], md=4),
                        dbc.Col([html.Label("Country", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-country", placeholder="Select...", clearable=True, disabled=True, style={"fontSize": "12px"})], md=4),
                        dbc.Col([html.Label("Site", style={"fontSize": "10px", "color": TM, "fontWeight": "600"}),
                            dcc.Dropdown(id="dd-site", placeholder="Select...", clearable=True, disabled=True, style={"fontSize": "12px"})], md=4),
                    ]),
                ], style={"background": CBG, "borderRadius": "10px", "padding": "14px 16px",
                    "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px"}),

                # TO RECIPIENTS
                html.Div([
                    html.Div(id="to-header", style={"fontWeight": "700", "fontSize": "12px", "color": P, "marginBottom": "6px"}),
                    dcc.Textarea(id="display-to", value="", readOnly=True,
                        style={"width": "100%", "height": "60px", "fontSize": "10px", "background": "#FAFBFC",
                            "border": f"1px solid {BD}", "borderRadius": "6px", "padding": "6px", "resize": "none"}),
                    html.Div([
                        dbc.Input(id="inp-add-to", placeholder="user@lilly.com; user2@lilly.com",
                            size="sm", style={"fontSize": "11px", "flex": "1", "borderRadius": "6px"}),
                        dbc.Button([html.I(className="fas fa-plus me-1"), "Add"], id="btn-add-to", size="sm",
                            style={"background": P, "border": "none", "fontSize": "11px", "marginLeft": "6px", "borderRadius": "6px"}),
                    ], className="d-flex align-items-center mt-2"),
                ], style={"background": CBG, "borderRadius": "10px", "padding": "14px 16px",
                    "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px"}),

                # BCC
                html.Div([
                    html.Div(id="bcc-header", style={"fontWeight": "700", "fontSize": "12px", "color": ACC, "marginBottom": "4px"}),
                    html.Div("Not visible to other recipients", style={"color": TL, "fontSize": "10px", "marginBottom": "6px"}),
                    html.Div([
                        dbc.Input(id="inp-add-bcc", placeholder="external@partner.com",
                            size="sm", style={"fontSize": "11px", "flex": "1", "borderRadius": "6px", "background": "#F0F7FF", "borderColor": "#B0D0F0"}),
                        dbc.Button([html.I(className="fas fa-plus me-1"), "BCC"], id="btn-add-bcc", size="sm",
                            style={"background": ACC, "border": "none", "fontSize": "11px", "marginLeft": "6px", "borderRadius": "6px"}),
                    ], className="d-flex align-items-center"),
                    html.Div(id="bcc-display", style={"fontSize": "10px", "color": ACC, "marginTop": "4px"}),
                ], style={"background": CBG, "borderRadius": "10px", "padding": "12px 16px",
                    "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px"}),

                # DOCUMENTS
                html.Div([
                    html.Div(id="doc-header", style={"fontWeight": "700", "fontSize": "12px", "color": P, "marginBottom": "4px"}),
                    html.Div("Select documents to include VV-TMF IDs in email",
                        style={"color": TL, "fontSize": "10px", "marginBottom": "6px"}),
                    html.Div(id="doc-list", style={"maxHeight": "180px", "overflowY": "auto",
                        "borderRadius": "6px", "border": f"1px solid {BD}", "background": "#FAFBFC"}),
                ], style={"background": CBG, "borderRadius": "10px", "padding": "12px 16px",
                    "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px", "flex": "1"}),

                # SEND
                html.Div([
                    html.Div([html.I(className="fas fa-paper-plane me-2", style={"color": P, "fontSize": "12px"}),
                        html.Span("Send & Save", style={"fontWeight": "700", "fontSize": "12px", "color": P})],
                        className="d-flex align-items-center mb-2"),
                    html.Div(id="send-summary", style={"fontSize": "12px", "color": TM, "marginBottom": "8px"}),
                    dbc.Button([html.I(className="fas fa-paper-plane me-2"), "Send Email & Save Communication"],
                        id="btn-send", size="sm", disabled=True,
                        style={"background": f"linear-gradient(135deg, {P}, {PL})", "border": "none",
                            "width": "100%", "height": "42px", "fontSize": "13px", "fontWeight": "600", "borderRadius": "8px"}),
                    html.Div(id="send-status", style={"marginTop": "8px"}),
                ], style={"background": CBG, "borderRadius": "10px", "padding": "14px 16px",
                    "boxShadow": "0 1px 4px rgba(0,0,0,0.06)"}),
            ], style={"flex": "1", "display": "flex", "flexDirection": "column"}),
        ], style={"display": "flex", "alignItems": "stretch"}),
    ], style={"padding": "16px 20px"}),

], style={"fontFamily": "'Segoe UI', -apple-system, sans-serif", "minHeight": "100vh", "background": BG})


# ═══════════════════════════════════════════════════════════════════════
#  CALLBACKS
# ═══════════════════════════════════════════════════════════════════════

# Header user
@callback(Output("header-user", "children"), Input("s-page", "data"))
def show_user(_):
    uid = get_current_user()
    name = get_user_display_name(uid)
    return [html.I(className="fas fa-user me-2", style={"fontSize": "11px"}), html.Span(name or uid)]

# Breadcrumb
@callback(Output("breadcrumb", "children"), [Input("s-page", "data"), Input("s-tpl", "data")])
def breadcrumb(page, tpl):
    return f"Home  ›  {tpl.get('Title', '')}" if page == "compose" and tpl else ""

# Toggle pages
@callback([Output("home-page", "style"), Output("compose-page", "style")], Input("s-page", "data"))
def toggle_page(page):
    if page == "compose":
        return {"display": "none"}, {"display": "block"}
    return {"display": "block"}, {"display": "none"}

# Home button
@callback([Output("s-page", "data", allow_duplicate=True), Output("s-tpl", "data", allow_duplicate=True),
    Output("s-to", "data", allow_duplicate=True), Output("s-bcc", "data", allow_duplicate=True),
    Output("dd-study", "value"), Output("dd-country", "value"), Output("dd-site", "value")],
    Input("btn-home", "n_clicks"), prevent_initial_call=True)
def go_home(n):
    return "home", None, [], [], None, None, None

# Template grid
@callback(Output("template-grid", "children"), Input("s-page", "data"))
def render_grid(page):
    df = get_cached("sc_lookup")
    if df.empty:
        return dbc.Alert("No templates found.", color="warning")
    cards = []
    for _, r in df.iterrows():
        title = str(r.get("Title", ""))
        tname = str(r.get("Template_Name", ""))
        tid = str(r.get("ID", ""))
        if not title or tid == "nan":
            continue
        cards.append(dbc.Col(
            html.Button(html.Div([
                html.I(className="fas fa-envelope-open-text", style={"fontSize": "20px", "marginBottom": "8px", "opacity": "0.9"}),
                html.Div(title, style={"fontWeight": "600", "fontSize": "12px", "lineHeight": "1.3"}),
                html.Div(tname, style={"fontSize": "9px", "opacity": "0.7", "marginTop": "4px", "maxHeight": "24px", "overflow": "hidden"}),
            ], style={"textAlign": "center", "padding": "10px 6px"}),
                id={"type": "tpl-btn", "index": tid},
                style={"background": f"linear-gradient(135deg, {P}, {PL})", "color": "white", "border": "none",
                    "borderRadius": "10px", "width": "100%", "height": "100px", "cursor": "pointer",
                    "boxShadow": "0 2px 6px rgba(26,35,126,0.2)"}),
            xs=6, sm=4, md=3, lg=2, className="mb-3"))
    return dbc.Row(cards)

# Template click
@callback([Output("s-page", "data", allow_duplicate=True), Output("s-tpl", "data", allow_duplicate=True),
    Output("s-to", "data", allow_duplicate=True), Output("s-bcc", "data", allow_duplicate=True),
    Output("inp-subject", "value"), Output("inp-body", "value"), Output("email-preview", "srcDoc")],
    Input({"type": "tpl-btn", "index": ALL}, "n_clicks"), prevent_initial_call=True)
def select_tpl(clicks):
    if not ctx.triggered_id or not any(c for c in clicks if c):
        return [no_update] * 7
    idx = str(ctx.triggered_id["index"])
    df = get_cached("sc_lookup")
    row = df[df["ID"].astype(str) == idx]
    if row.empty:
        return [no_update] * 7
    r = row.iloc[0]
    tpl = {col: str(r.get(col, "") or "") for col in
        ["ID", "Title", "Template_Name", "EmailSubject", "EmailBody",
         "Classifications", "Lilly_Groups", "Non_Lilly_Roles", "DocURL"]}
    body = tpl["EmailBody"]
    subject = tpl["EmailSubject"]
    # Wrap HTML body for preview
    preview = f"""<html><head><style>body{{font-family:'Segoe UI',sans-serif;font-size:13px;padding:12px;color:#1E293B;line-height:1.6;}}</style></head><body>{body}</body></html>"""
    return "compose", tpl, [], [], subject, body, preview

# Toggle source/preview
@callback([Output("email-preview", "style"), Output("inp-body", "style"), Output("btn-toggle-edit", "children")],
    Input("btn-toggle-edit", "n_clicks"), State("btn-toggle-edit", "children"), prevent_initial_call=True)
def toggle_edit(n, label):
    if label == "Edit Source":
        return ({"width": "100%", "height": "450px", "border": f"1px solid {BD}", "borderRadius": "6px", "background": "white", "display": "none"},
            {"width": "100%", "height": "450px", "fontSize": "12px", "border": f"1px solid {BD}", "borderRadius": "6px",
                "padding": "12px", "fontFamily": "monospace", "resize": "vertical", "background": "#FAFBFC", "display": "block"}, "Preview")
    return ({"width": "100%", "height": "450px", "border": f"1px solid {BD}", "borderRadius": "6px", "background": "white", "display": "block"},
        {"width": "100%", "height": "450px", "fontSize": "12px", "border": f"1px solid {BD}", "borderRadius": "6px",
            "padding": "12px", "fontFamily": "monospace", "resize": "vertical", "background": "#FAFBFC", "display": "none"}, "Edit Source")

# Update preview when source changes
@callback(Output("email-preview", "srcDoc", allow_duplicate=True),
    Input("inp-body", "value"), prevent_initial_call=True)
def update_preview(body):
    if not body: return ""
    return f"""<html><head><style>body{{font-family:'Segoe UI',sans-serif;font-size:13px;padding:12px;color:#1E293B;line-height:1.6;}}</style></head><body>{body}</body></html>"""


# ═══════ FILTER CASCADE ═══════

@callback(Output("dd-study", "options"), Input("s-page", "data"))
def load_studies(page):
    df = get_cached("study_country_site_lookup")
    if df.empty or "Study_Alias" not in df.columns:
        return []
    return [{"label": s, "value": s} for s in sorted(df["Study_Alias"].dropna().unique().tolist()) if s]

@callback([Output("dd-country", "options"), Output("dd-country", "disabled"),
    Output("dd-country", "value", allow_duplicate=True), Output("dd-site", "value", allow_duplicate=True),
    Output("dd-site", "disabled")],
    Input("dd-study", "value"), prevent_initial_call=True)
def cascade_country(study):
    if not study:
        return [], True, None, None, True
    df = get_cached("study_country_site_lookup")
    f = df[df["Study_Alias"] == study]
    if "Country_Name" in f.columns:
        c = sorted(f["Country_Name"].dropna().unique().tolist())
        return [{"label": x, "value": x} for x in c if x], False, None, None, True
    return [], True, None, None, True

@callback([Output("dd-site", "options"), Output("dd-site", "disabled", allow_duplicate=True)],
    Input("dd-country", "value"), State("dd-study", "value"), prevent_initial_call=True)
def cascade_site(country, study):
    if not country or not study:
        return [], True
    df = get_cached("study_country_site_lookup")
    f = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country)]
    if "Site" in f.columns:
        return [{"label": x, "value": x} for x in sorted(f["Site"].dropna().unique().tolist()) if x], False
    return [], True


# ═══════ AUTO-POPULATE RECIPIENTS + DOCS ON FILTER CHANGE ═══════

@callback([Output("s-to", "data", allow_duplicate=True),
    Output("inp-subject", "value", allow_duplicate=True),
    Output("inp-body", "value", allow_duplicate=True),
    Output("email-preview", "srcDoc", allow_duplicate=True),
    Output("doc-list", "children")],
    [Input("dd-study", "value"), Input("dd-country", "value"), Input("dd-site", "value")],
    State("s-tpl", "data"), prevent_initial_call=True)
def on_filter(study, country, site, tpl):
    if not tpl:
        return no_update, no_update, no_update, no_update, no_update

    body = tpl.get("EmailBody", "")
    body = body.replace("{{STUDY_ALIAS}}", study or "").replace("{{COUNTRY}}", country or "")
    body = body.replace("{{SITE}}", site or "").replace("{{DATE}}", date.today().strftime("%d %B %Y"))
    subject = tpl.get("EmailSubject", "").replace("{{STUDY_ALIAS}}", study or "")
    preview = f"""<html><head><style>body{{font-family:'Segoe UI',sans-serif;font-size:13px;padding:12px;color:#1E293B;line-height:1.6;}}</style></head><body>{body}</body></html>"""

    recipients = _auto_recips(tpl, study, country, site)
    docs = _render_docs(tpl, study, country, site)

    return recipients, subject, body, preview, docs


def _auto_recips(tpl, study, country, site):
    lg = [g.strip() for g in tpl.get("Lilly_Groups", "").replace("\n", ";").split(";") if g.strip()]
    nlr = [g.strip() for g in tpl.get("Non_Lilly_Roles", "").replace("\n", ";").split(";") if g.strip()]
    emails = set()
    if not study: return sorted(list(emails))

    if lg:
        try:
            df = get_cached("study_sponsor_personnel_assignment")
            if not df.empty:
                m = df[(df["Study_Alias"] == study) & (df["Study_Team_Role"].isin(lg)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                emails.update(m["Email_Address"].str.strip().tolist())
        except: pass

    if country and lg:
        try:
            df = get_cached("country_sponsor_personnel_assignment")
            if not df.empty:
                m = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country) & (df["Study_Team_Role"].isin(lg)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                emails.update(m["Email_Address"].str.strip().tolist())
        except: pass

    if site:
        all_r = list(set(lg + nlr))
        if all_r:
            try:
                df = get_cached("study_site_sponsor_personnel_combined")
                if not df.empty:
                    m = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country) & (df["Site"] == site) & (df["Role"].isin(all_r)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                    emails.update(m["Email_Address"].str.strip().tolist())
            except: pass

    return sorted([e for e in emails if e])


def _render_docs(tpl, study, country, site):
    if not study:
        return html.Div("Select a study to load documents.", style={"color": TL, "fontSize": "11px", "padding": "12px", "textAlign": "center"})
    try:
        df = get_cached("documents")
        if df.empty:
            return html.Div("No documents available.", style={"color": TL, "fontSize": "11px", "padding": "12px", "textAlign": "center"})

        mask = df["Study_Alias"] == study
        s_lvl = mask & (df["Site"].isna() | (df["Site"] == "")) & (df["Country_Code"].isna() | (df["Country_Code"] == ""))
        c_lvl = pd.Series(False, index=df.index)
        if country: c_lvl = mask & (df["Country_Code"] == country) & (df["Site"].isna() | (df["Site"] == ""))
        st_lvl = pd.Series(False, index=df.index)
        if site: st_lvl = mask & (df["Site"] == site)

        filtered = df[s_lvl | c_lvl | st_lvl]
        cls = tpl.get("Classifications", "")
        if cls:
            cl = [c.strip() for c in cls.replace("\n", ";").split(";") if c.strip()]
            if cl: filtered = filtered[filtered["Classification"].isin(cl)]

        filtered = filtered.head(150)
        if filtered.empty:
            return html.Div("No documents found.", style={"color": TL, "fontSize": "11px", "padding": "12px", "textAlign": "center"})

        items = []
        for _, doc in filtered.iterrows():
            dn, dname, dl = str(doc.get("Document_Number","")), str(doc.get("Document_Name","")), str(doc.get("DocLink",""))
            cl, df_dt = str(doc.get("Classification","")), str(doc.get("Date_Document_Finalized",""))[:10]
            sv, cv = doc.get("Site",""), doc.get("Country_Code","")
            if not sv and not cv: lbl, lc = "Study", P
            elif cv and not sv: lbl, lc = "Country", SUC
            else: lbl, lc = "Site", WRN

            items.append(html.Div([html.Div([
                html.A(f"{dn} - {dname}", href=dl, target="_blank",
                    style={"color": P, "fontSize": "11px", "fontWeight": "600", "textDecoration": "none"}),
                html.Div([html.Span(f"{cl} | {df_dt}", style={"color": TL, "fontSize": "9px"}),
                    html.Span(lbl, style={"fontSize": "9px", "fontWeight": "600", "color": lc,
                        "marginLeft": "8px", "background": f"{lc}15", "padding": "1px 6px", "borderRadius": "3px"})]),
            ])], style={"padding": "6px 10px", "borderBottom": "1px solid #F0F0F0"}))
        return html.Div(items)
    except Exception as e:
        return html.Div(f"Error: {e}", style={"color": DNG, "fontSize": "11px", "padding": "12px"})


# ═══════ UPDATE DISPLAYS FROM STORES ═══════

@callback([Output("display-to", "value"), Output("to-header", "children"), Output("btn-send", "disabled"),
    Output("send-summary", "children")],
    [Input("s-to", "data"), Input("s-bcc", "data")])
def update_to_display(to_r, bcc_r):
    to_r = to_r or []
    bcc_r = bcc_r or []
    return ("; ".join(to_r),
        [html.I(className="fas fa-users me-2", style={"color": P, "fontSize": "12px"}), f"TO Recipients ({len(to_r)})"],
        len(to_r) == 0,
        f"TO: {len(to_r)}  |  BCC: {len(bcc_r)}  |  Docs: 0")

@callback([Output("bcc-display", "children"), Output("bcc-header", "children")], Input("s-bcc", "data"))
def update_bcc_display(bcc_r):
    bcc_r = bcc_r or []
    return ("; ".join(bcc_r) if bcc_r else "",
        [html.I(className="fas fa-eye-slash me-2", style={"color": ACC, "fontSize": "12px"}), f"BCC — External ({len(bcc_r)})"])

@callback(Output("doc-header", "children"), Input("doc-list", "children"))
def update_doc_header(_):
    return [html.I(className="fas fa-file-alt me-2", style={"color": P, "fontSize": "12px"}), "eTMF & Non-TMF Documents"]


# ═══════ ADD TO / BCC ═══════

@callback([Output("s-to", "data", allow_duplicate=True), Output("inp-add-to", "value")],
    Input("btn-add-to", "n_clicks"), [State("inp-add-to", "value"), State("s-to", "data")], prevent_initial_call=True)
def add_to(n, text, cur):
    if not text: return no_update, no_update
    cur = cur or []
    new = [e.strip() for e in text.replace(",", ";").split(";") if e.strip()]
    ex = set(cur)
    for e in new:
        if e not in ex: cur.append(e); ex.add(e)
    return cur, ""

@callback([Output("s-bcc", "data", allow_duplicate=True), Output("inp-add-bcc", "value")],
    Input("btn-add-bcc", "n_clicks"), [State("inp-add-bcc", "value"), State("s-bcc", "data")], prevent_initial_call=True)
def add_bcc(n, text, cur):
    if not text: return no_update, no_update
    cur = cur or []
    new = [e.strip() for e in text.replace(",", ";").split(";") if e.strip()]
    ex = set(cur)
    for e in new:
        if e not in ex: cur.append(e); ex.add(e)
    return cur, ""


# ═══════ SEND (placeholder) ═══════

@callback(Output("send-status", "children"), Input("btn-send", "n_clicks"),
    [State("s-to", "data"), State("s-bcc", "data")], prevent_initial_call=True)
def send_email(n, to_r, bcc_r):
    return dbc.Alert([html.I(className="fas fa-info-circle me-2"),
        f"Ready to send to {len(to_r or [])} TO and {len(bcc_r or [])} BCC recipients. ",
        "Integrate email service to enable sending."], color="info", dismissable=True, style={"fontSize": "11px"})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8050)
