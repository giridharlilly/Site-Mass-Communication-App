"""
app.py — Site Mass Communication v3.0
Dash app for Posit Connect
Layout: Option B — Home (template grid) → Compose (50/50 split)
Theme: Dark blue (matching FlaskMeetingApp)
Data: Fabric Lakehouse (OneLake HTTPS) + AD group access control
"""

import os
import json
import time
import dash
from dash import dcc, html, dash_table, Input, Output, State, callback, ctx, ALL, no_update
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

# ═══════════════════════════════════════════════════════════════════════
#  APP SETUP
# ═══════════════════════════════════════════════════════════════════════
app = dash.Dash(__name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP, dbc.icons.FONT_AWESOME],
    suppress_callback_exceptions=True,
    title="Site Mass Communication")
server = app.server
enforce_access(server)

# ═══════════════════════════════════════════════════════════════════════
#  THEME (dark blue — matching FlaskMeetingApp)
# ═══════════════════════════════════════════════════════════════════════
PRIMARY = "#1a237e"
PRIMARY_LIGHT = "#283593"
PRIMARY_LIGHTER = "#E8EAF6"
ACCENT = "#3B82F6"
SUCCESS = "#10B981"
DANGER = "#EF4444"
WARNING = "#F59E0B"
PAGE_BG = "#F5F7FA"
CARD_BG = "#FFFFFF"
TEXT_DARK = "#1E293B"
TEXT_MUTED = "#64748B"
TEXT_LIGHT = "#94A3B8"
BORDER = "#E2E8F0"

# ═══════════════════════════════════════════════════════════════════════
#  DATA CACHE
# ═══════════════════════════════════════════════════════════════════════
_cache = {}
_cache_ts = {}
CACHE_TTL = 300

def get_cached(table_name, force=False):
    now = time.time()
    if not force and table_name in _cache and (now - _cache_ts.get(table_name, 0)) < CACHE_TTL:
        return _cache[table_name].copy()
    try:
        df = read_table(table_name)
        _cache[table_name] = df
        _cache_ts[table_name] = now
        return df.copy()
    except Exception as e:
        print(f"Error reading {table_name}: {e}")
        return pd.DataFrame()

def preload_startup():
    tables = ["SC_LookUp", "study_country_site_lookup"]
    print(f"DEBUG ABFSS_BASE: {os.getenv('ABFSS_BASE', 'NOT SET')}")
    print(f"DEBUG FABRIC_CLIENT_ID: {os.getenv('FABRIC_CLIENT_ID', 'NOT SET')[:10]}...")
    for name in tables:
        try:
            df = read_table(name)
            _cache[name] = df
            _cache_ts[name] = time.time()
            print(f"DEBUG: {name} loaded = {len(df)} rows")
        except Exception as e:
            print(f"DEBUG: {name} FAILED = {e}")
            _cache[name] = pd.DataFrame()
        except:
            pass

preload_startup()

# ═══════════════════════════════════════════════════════════════════════
#  LAYOUT
# ═══════════════════════════════════════════════════════════════════════
app.layout = html.Div([
    dcc.Store(id="store-page", data="home"),
    dcc.Store(id="store-template", data=None),
    dcc.Store(id="store-to", data=[]),
    dcc.Store(id="store-bcc", data=[]),
    dcc.Store(id="store-docs", data=[]),

    # HEADER
    html.Div([
        html.Div([
            html.I(className="fas fa-home", id="btn-home",
                style={"fontSize": "18px", "cursor": "pointer", "color": "rgba(255,255,255,0.8)", "marginRight": "16px"}),
            html.Span("Site Mass Communication",
                style={"fontWeight": "700", "fontSize": "18px", "color": "white", "letterSpacing": "0.3px"}),
            html.Span(id="breadcrumb", style={"color": "rgba(255,255,255,0.6)", "fontSize": "12px", "marginLeft": "16px"}),
        ], className="d-flex align-items-center"),
        html.Div([
            html.Span("v3.0", style={"color": "rgba(255,255,255,0.4)", "fontSize": "10px", "marginRight": "16px"}),
            html.Div(id="header-user",
                style={"background": "rgba(255,255,255,0.1)", "borderRadius": "20px",
                    "padding": "4px 14px", "fontSize": "13px", "color": "white"}),
        ], className="d-flex align-items-center"),
    ], className="d-flex justify-content-between align-items-center px-4",
        style={"background": f"linear-gradient(135deg, {PRIMARY}, {PRIMARY_LIGHT})",
            "height": "56px", "boxShadow": "0 2px 8px rgba(0,0,0,0.15)"}),

    html.Div(id="page-content", style={"minHeight": "calc(100vh - 56px)", "background": PAGE_BG}),
    html.Div(id="toast-msg"),
], style={"fontFamily": "'Segoe UI', -apple-system, sans-serif"})


# ═══════════════════════════════════════════════════════════════════════
#  CALLBACKS
# ═══════════════════════════════════════════════════════════════════════

@callback(Output("header-user", "children"), Input("store-page", "data"))
def show_user(_):
    uid = get_current_user()
    name = get_user_display_name(uid)
    return [html.I(className="fas fa-user me-2", style={"fontSize": "11px"}), html.Span(name or uid)]

@callback(Output("breadcrumb", "children"), [Input("store-page", "data"), Input("store-template", "data")])
def breadcrumb(page, tpl):
    if page == "compose" and tpl:
        return f"Home  ›  {tpl.get('Title', '')}"
    return ""

@callback([Output("store-page", "data", allow_duplicate=True), Output("store-template", "data", allow_duplicate=True),
    Output("store-to", "data", allow_duplicate=True), Output("store-bcc", "data", allow_duplicate=True),
    Output("store-docs", "data", allow_duplicate=True)],
    Input("btn-home", "n_clicks"), prevent_initial_call=True)
def go_home(n):
    return "home", None, [], [], []

@callback([Output("store-page", "data", allow_duplicate=True), Output("store-template", "data", allow_duplicate=True),
    Output("store-to", "data", allow_duplicate=True), Output("store-bcc", "data", allow_duplicate=True),
    Output("store-docs", "data", allow_duplicate=True)],
    Input({"type": "tpl-btn", "index": ALL}, "n_clicks"), prevent_initial_call=True)
def select_template(clicks):
    if not ctx.triggered_id or not any(c for c in clicks if c):
        return [no_update] * 5
    idx = str(ctx.triggered_id["index"])
    df = get_cached("SC_LookUp")
    if df.empty:
        return [no_update] * 5
    row = df[df["ID"].astype(str) == idx]
    if row.empty:
        return [no_update] * 5
    r = row.iloc[0]
    tpl = {col: str(r.get(col, "") or "") for col in
        ["ID", "Title", "Template_Name", "EmailSubject", "EmailBody",
         "Classifications", "Lilly_Groups", "Non_Lilly_Roles", "DocURL"]}
    return "compose", tpl, [], [], []

# ── Render page ──
@callback(Output("page-content", "children"),
    [Input("store-page", "data"), Input("store-template", "data"),
     Input("store-to", "data"), Input("store-bcc", "data"), Input("store-docs", "data")])
def render_page(page, tpl, to_r, bcc_r, docs):
    if page == "compose" and tpl:
        return render_compose(tpl, to_r or [], bcc_r or [], docs or [])
    return render_home()


def render_home():
    df = get_cached("SC_LookUp")
    if df.empty:
        return html.Div(dbc.Alert("No templates found.", color="warning"), style={"padding": "24px"})
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
                html.Div(tname, style={"fontSize": "9px", "opacity": "0.7", "marginTop": "4px",
                    "lineHeight": "1.2", "maxHeight": "24px", "overflow": "hidden"}),
            ], style={"textAlign": "center", "padding": "10px 6px"}),
                id={"type": "tpl-btn", "index": tid},
                style={"background": f"linear-gradient(135deg, {PRIMARY}, {PRIMARY_LIGHT})",
                    "color": "white", "border": "none", "borderRadius": "10px",
                    "width": "100%", "height": "100px", "cursor": "pointer",
                    "boxShadow": "0 2px 6px rgba(26,35,126,0.2)"}),
            xs=6, sm=4, md=3, lg=2, className="mb-3"))

    return html.Div([
        html.H4("Select a Communication Template",
            style={"fontWeight": "600", "color": TEXT_DARK, "marginBottom": "4px"}),
        html.P("Choose a template to compose and send mass communications",
            style={"color": TEXT_MUTED, "fontSize": "13px", "marginBottom": "24px"}),
        dbc.Row(cards),
    ], style={"padding": "24px 32px"})


def render_compose(tpl, to_r, bcc_r, docs):
    return html.Div([html.Div([
        # ════ LEFT: Email Compose (50%) ════
        html.Div([html.Div([
            html.Div([html.I(className="fas fa-pen-to-square me-2", style={"color": PRIMARY}),
                html.Span("Email Composition", style={"fontWeight": "700", "fontSize": "14px", "color": PRIMARY})],
                className="d-flex align-items-center mb-3"),
            html.Label("Subject", style={"fontSize": "11px", "fontWeight": "600", "color": TEXT_MUTED, "marginBottom": "4px"}),
            dbc.Input(id="inp-subject", value=tpl.get("EmailSubject", ""), size="sm",
                style={"fontSize": "13px", "marginBottom": "14px", "background": "#FAFBFC",
                    "border": f"1px solid {BORDER}", "borderRadius": "6px"}),
            html.Label("Email Body", style={"fontSize": "11px", "fontWeight": "600", "color": TEXT_MUTED, "marginBottom": "4px"}),
            dcc.Textarea(id="inp-body", value=tpl.get("EmailBody", ""),
                style={"width": "100%", "minHeight": "420px", "fontSize": "13px",
                    "border": f"1px solid {BORDER}", "borderRadius": "6px",
                    "padding": "12px", "fontFamily": "'Segoe UI', sans-serif",
                    "resize": "vertical", "background": "#FAFBFC", "lineHeight": "1.6"}),
            html.Div([html.Span("Placeholders: ", style={"fontWeight": "600", "fontSize": "10px"}),
                html.Span("{{STUDY_ALIAS}}  {{COUNTRY}}  {{SITE}}  {{DATE}}  {{DOC_IDS}}",
                    style={"fontSize": "10px", "fontFamily": "monospace"})],
                style={"color": TEXT_LIGHT, "marginTop": "8px"}),
        ], style={"background": CARD_BG, "borderRadius": "10px", "padding": "20px",
            "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "height": "100%"})],
            style={"flex": "1", "marginRight": "16px", "display": "flex"}),

        # ════ RIGHT: Filters + Recipients + Docs + Send (50%) ════
        html.Div([
            # FILTERS
            html.Div([
                html.Div([html.I(className="fas fa-filter me-2", style={"color": PRIMARY, "fontSize": "12px"}),
                    html.Span("FILTERS", style={"fontWeight": "700", "fontSize": "11px", "color": PRIMARY, "letterSpacing": "1px"})],
                    className="d-flex align-items-center mb-2"),
                dbc.Row([
                    dbc.Col([html.Label("Study Alias", style={"fontSize": "10px", "color": TEXT_MUTED, "fontWeight": "600"}),
                        dcc.Dropdown(id="dd-study", placeholder="Select study...", clearable=True, style={"fontSize": "12px"})], md=4),
                    dbc.Col([html.Label("Country", style={"fontSize": "10px", "color": TEXT_MUTED, "fontWeight": "600"}),
                        dcc.Dropdown(id="dd-country", placeholder="Select...", clearable=True, disabled=True, style={"fontSize": "12px"})], md=4),
                    dbc.Col([html.Label("Site", style={"fontSize": "10px", "color": TEXT_MUTED, "fontWeight": "600"}),
                        dcc.Dropdown(id="dd-site", placeholder="Select...", clearable=True, disabled=True, style={"fontSize": "12px"})], md=4),
                ]),
            ], style={"background": CARD_BG, "borderRadius": "10px", "padding": "14px 16px",
                "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px"}),

            # TO RECIPIENTS
            html.Div([
                html.Div([html.I(className="fas fa-users me-2", style={"color": PRIMARY, "fontSize": "12px"}),
                    html.Span(f"TO Recipients ({len(to_r)})", style={"fontWeight": "700", "fontSize": "12px", "color": PRIMARY})],
                    className="d-flex align-items-center mb-2"),
                dcc.Textarea(id="display-to", value="; ".join(to_r), readOnly=True,
                    style={"width": "100%", "height": "60px", "fontSize": "10px", "background": "#FAFBFC",
                        "border": f"1px solid {BORDER}", "borderRadius": "6px", "padding": "6px", "resize": "none"}),
                html.Div([
                    dbc.Input(id="inp-add-to", placeholder="user@lilly.com; user2@lilly.com",
                        size="sm", style={"fontSize": "11px", "flex": "1", "borderRadius": "6px"}),
                    dbc.Button([html.I(className="fas fa-plus me-1"), "Add"], id="btn-add-to", size="sm",
                        style={"background": PRIMARY, "border": "none", "fontSize": "11px", "marginLeft": "6px", "borderRadius": "6px"}),
                ], className="d-flex align-items-center mt-2"),
            ], style={"background": CARD_BG, "borderRadius": "10px", "padding": "14px 16px",
                "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px"}),

            # BCC
            html.Div([
                html.Div([html.I(className="fas fa-eye-slash me-2", style={"color": ACCENT, "fontSize": "12px"}),
                    html.Span(f"BCC — External ({len(bcc_r)})", style={"fontWeight": "700", "fontSize": "12px", "color": ACCENT})],
                    className="d-flex align-items-center mb-1"),
                html.Div("Not visible to other recipients", style={"color": TEXT_LIGHT, "fontSize": "10px", "marginBottom": "6px"}),
                html.Div([
                    dbc.Input(id="inp-add-bcc", placeholder="external@partner.com",
                        size="sm", style={"fontSize": "11px", "flex": "1", "borderRadius": "6px", "background": "#F0F7FF", "borderColor": "#B0D0F0"}),
                    dbc.Button([html.I(className="fas fa-plus me-1"), "BCC"], id="btn-add-bcc", size="sm",
                        style={"background": ACCENT, "border": "none", "fontSize": "11px", "marginLeft": "6px", "borderRadius": "6px"}),
                ], className="d-flex align-items-center"),
                html.Div("; ".join(bcc_r) if bcc_r else "", style={"fontSize": "10px", "color": ACCENT, "marginTop": "4px", "maxHeight": "36px", "overflow": "auto"}),
            ], style={"background": CARD_BG, "borderRadius": "10px", "padding": "12px 16px",
                "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px"}),

            # DOCUMENTS
            html.Div([
                html.Div([html.I(className="fas fa-file-alt me-2", style={"color": PRIMARY, "fontSize": "12px"}),
                    html.Span(f"eTMF & Non-TMF Documents ({len(docs)} selected)",
                        style={"fontWeight": "700", "fontSize": "12px", "color": PRIMARY})],
                    className="d-flex align-items-center mb-1"),
                html.Div("Select documents to include VV-TMF IDs in email",
                    style={"color": TEXT_LIGHT, "fontSize": "10px", "marginBottom": "6px"}),
                html.Div(id="doc-list", style={"maxHeight": "180px", "overflowY": "auto",
                    "borderRadius": "6px", "border": f"1px solid {BORDER}", "background": "#FAFBFC"}),
            ], style={"background": CARD_BG, "borderRadius": "10px", "padding": "12px 16px",
                "boxShadow": "0 1px 4px rgba(0,0,0,0.06)", "marginBottom": "10px", "flex": "1"}),

            # SEND
            html.Div([
                html.Div([html.I(className="fas fa-paper-plane me-2", style={"color": PRIMARY, "fontSize": "12px"}),
                    html.Span("Send & Save", style={"fontWeight": "700", "fontSize": "12px", "color": PRIMARY})],
                    className="d-flex align-items-center mb-2"),
                html.Div(f"TO: {len(to_r)}  |  BCC: {len(bcc_r)}  |  Docs: {len(docs)}",
                    style={"fontSize": "12px", "color": TEXT_MUTED, "marginBottom": "8px"}),
                dbc.Button([html.I(className="fas fa-paper-plane me-2"), "Send Email & Save Communication"],
                    id="btn-send", size="sm", disabled=len(to_r) == 0,
                    style={"background": f"linear-gradient(135deg, {PRIMARY}, {PRIMARY_LIGHT})",
                        "border": "none", "width": "100%", "height": "42px",
                        "fontSize": "13px", "fontWeight": "600", "borderRadius": "8px"}),
                html.Div(id="send-status", style={"marginTop": "8px"}),
            ], style={"background": CARD_BG, "borderRadius": "10px", "padding": "14px 16px",
                "boxShadow": "0 1px 4px rgba(0,0,0,0.06)"}),
        ], style={"flex": "1", "display": "flex", "flexDirection": "column"}),
    ], style={"display": "flex", "alignItems": "stretch"})], style={"padding": "16px 20px"})


# ═══════════════════════════════════════════════════════════════════════
#  FILTER CASCADING
# ═══════════════════════════════════════════════════════════════════════

@callback(Output("dd-study", "options"), Input("store-page", "data"))
def load_studies(page):
    if page != "compose":
        return []
    df = get_cached("study_country_site_lookup")
    if df.empty or "Study_Alias" not in df.columns:
        return []
    return [{"label": s, "value": s} for s in sorted(df["Study_Alias"].dropna().unique().tolist()) if s]

@callback([Output("dd-country", "options"), Output("dd-country", "disabled"),
    Output("dd-country", "value"), Output("dd-site", "value"), Output("dd-site", "disabled")],
    Input("dd-study", "value"))
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
        s = sorted(f["Site"].dropna().unique().tolist())
        return [{"label": x, "value": x} for x in s if x], False
    return [], True


# ═══════════════════════════════════════════════════════════════════════
#  AUTO-POPULATE ON FILTER CHANGE
# ═══════════════════════════════════════════════════════════════════════

@callback([Output("store-to", "data", allow_duplicate=True), Output("inp-subject", "value"),
    Output("inp-body", "value"), Output("doc-list", "children")],
    [Input("dd-study", "value"), Input("dd-country", "value"), Input("dd-site", "value")],
    State("store-template", "data"), prevent_initial_call=True)
def on_filter_change(study, country, site, tpl):
    if not tpl:
        return no_update, no_update, no_update, no_update

    body = tpl.get("EmailBody", "")
    body = body.replace("{{STUDY_ALIAS}}", study or "").replace("{{COUNTRY}}", country or "")
    body = body.replace("{{SITE}}", site or "").replace("{{DATE}}", date.today().strftime("%d %B %Y"))

    subject = tpl.get("EmailSubject", "").replace("{{STUDY_ALIAS}}", study or "")

    recipients = _auto_recipients(tpl, study, country, site)
    doc_html = _render_docs(tpl, study, country, site)

    return recipients, subject, body, doc_html


def _auto_recipients(tpl, study, country, site):
    lg = [g.strip() for g in tpl.get("Lilly_Groups", "").replace("\n", ";").split(";") if g.strip()]
    nlr = [g.strip() for g in tpl.get("Non_Lilly_Roles", "").replace("\n", ";").split(";") if g.strip()]
    emails = set()
    if not study:
        return sorted(list(emails))

    if lg:
        try:
            df = get_cached("Study_Sponsor_Personnel_Assignment")
            if not df.empty:
                m = df[(df["Study_Alias"] == study) & (df["Study_Team_Role"].isin(lg)) &
                    (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                emails.update(m["Email_Address"].str.strip().tolist())
        except:
            pass

    if country and lg:
        try:
            df = get_cached("Country_Sponsor_Personnel_Assignment")
            if not df.empty:
                m = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country) &
                    (df["Study_Team_Role"].isin(lg)) & (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                emails.update(m["Email_Address"].str.strip().tolist())
        except:
            pass

    if site:
        all_r = list(set(lg + nlr))
        if all_r:
            try:
                df = get_cached("Study_Site_Sponsor_Personnel_Combined")
                if not df.empty:
                    m = df[(df["Study_Alias"] == study) & (df["Country_Name"] == country) &
                        (df["Site"] == site) & (df["Role"].isin(all_r)) &
                        (df["Email_Address"].notna()) & (df["Email_Address"] != "")]
                    emails.update(m["Email_Address"].str.strip().tolist())
            except:
                pass

    return sorted([e for e in emails if e])


def _render_docs(tpl, study, country, site):
    if not study:
        return html.Div("Select a study to load documents.",
            style={"color": TEXT_LIGHT, "fontSize": "11px", "padding": "12px", "textAlign": "center"})
    try:
        df = get_cached("Documents")
        if df.empty:
            return html.Div("No documents available.",
                style={"color": TEXT_LIGHT, "fontSize": "11px", "padding": "12px", "textAlign": "center"})

        mask = df["Study_Alias"] == study
        s_lvl = mask & (df["Site"].isna() | (df["Site"] == "")) & (df["Country_Code"].isna() | (df["Country_Code"] == ""))
        c_lvl = pd.Series(False, index=df.index)
        if country:
            c_lvl = mask & (df["Country_Code"] == country) & (df["Site"].isna() | (df["Site"] == ""))
        st_lvl = pd.Series(False, index=df.index)
        if site:
            st_lvl = mask & (df["Site"] == site)

        filtered = df[s_lvl | c_lvl | st_lvl]

        cls = tpl.get("Classifications", "")
        if cls:
            cls_list = [c.strip() for c in cls.replace("\n", ";").split(";") if c.strip()]
            if cls_list:
                filtered = filtered[filtered["Classification"].isin(cls_list)]

        filtered = filtered.head(150)
        if filtered.empty:
            return html.Div("No documents found for this selection.",
                style={"color": TEXT_LIGHT, "fontSize": "11px", "padding": "12px", "textAlign": "center"})

        items = []
        for _, doc in filtered.iterrows():
            dn = str(doc.get("Document_Number", ""))
            dname = str(doc.get("Document_Name", ""))
            dl = str(doc.get("DocLink", ""))
            cl = str(doc.get("Classification", ""))
            df_date = str(doc.get("Date_Document_Finalized", ""))[:10]
            sv = doc.get("Site", "")
            cv = doc.get("Country_Code", "")
            if not sv and not cv:
                lbl, lc = "Study", PRIMARY
            elif cv and not sv:
                lbl, lc = "Country", SUCCESS
            else:
                lbl, lc = "Site", WARNING

            items.append(html.Div([html.Div([
                html.A(f"{dn} - {dname}", href=dl, target="_blank",
                    style={"color": PRIMARY, "fontSize": "11px", "fontWeight": "600", "textDecoration": "none"}),
                html.Div([html.Span(f"{cl}  |  {df_date}", style={"color": TEXT_LIGHT, "fontSize": "9px"}),
                    html.Span(lbl, style={"fontSize": "9px", "fontWeight": "600", "color": lc,
                        "marginLeft": "8px", "background": f"{lc}15", "padding": "1px 6px", "borderRadius": "3px"})]),
            ])], style={"padding": "6px 10px", "borderBottom": "1px solid #F0F0F0"}))

        return html.Div(items)
    except Exception as e:
        return html.Div(f"Error: {e}", style={"color": DANGER, "fontSize": "11px", "padding": "12px"})


# ═══════════════════════════════════════════════════════════════════════
#  ADD TO / BCC
# ═══════════════════════════════════════════════════════════════════════

@callback([Output("store-to", "data", allow_duplicate=True), Output("inp-add-to", "value")],
    Input("btn-add-to", "n_clicks"), [State("inp-add-to", "value"), State("store-to", "data")], prevent_initial_call=True)
def add_to(n, text, cur):
    if not text: return no_update, no_update
    cur = cur or []
    new = [e.strip() for e in text.replace(",", ";").split(";") if e.strip()]
    ex = set(cur)
    for e in new:
        if e not in ex: cur.append(e); ex.add(e)
    return cur, ""

@callback([Output("store-bcc", "data", allow_duplicate=True), Output("inp-add-bcc", "value")],
    Input("btn-add-bcc", "n_clicks"), [State("inp-add-bcc", "value"), State("store-bcc", "data")], prevent_initial_call=True)
def add_bcc(n, text, cur):
    if not text: return no_update, no_update
    cur = cur or []
    new = [e.strip() for e in text.replace(",", ";").split(";") if e.strip()]
    ex = set(cur)
    for e in new:
        if e not in ex: cur.append(e); ex.add(e)
    return cur, ""


# ═══════════════════════════════════════════════════════════════════════
#  SEND (placeholder)
# ═══════════════════════════════════════════════════════════════════════

@callback(Output("send-status", "children"), Input("btn-send", "n_clicks"),
    [State("store-to", "data"), State("store-bcc", "data"), State("store-docs", "data")],
    prevent_initial_call=True)
def send_email(n, to_r, bcc_r, docs):
    return dbc.Alert([html.I(className="fas fa-info-circle me-2"),
        f"Ready to send to {len(to_r or [])} TO and {len(bcc_r or [])} BCC recipients. ",
        "Integrate email service to enable sending."],
        color="info", dismissable=True, style={"fontSize": "11px"})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8050)
