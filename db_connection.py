"""
db_connection.py
================
Connects to Microsoft Fabric Lakehouse via Delta Lake (deltalake library).
Writes proper Delta tables to Tables/dbo/{table_name} — visible in SQL endpoint & Power BI.
Falls back to parquet in Files/ if deltalake write fails.
"""

import os
import io
import time
import logging

import requests
import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

logger = logging.getLogger(__name__)

# ── Configuration ─────────────────────────────────────────────────────
FABRIC_CLIENT_ID = os.getenv("FABRIC_CLIENT_ID")
FABRIC_CLIENT_SECRET = os.getenv("FABRIC_CLIENT_SECRET")
FABRIC_TENANT_ID = os.getenv("FABRIC_TENANT_ID")
WORKSPACE_NAME = os.getenv("FABRIC_WORKSPACE_NAME", "DPA_PowerPlatform")
LAKEHOUSE_NAME = os.getenv("FABRIC_LAKEHOUSE_NAME", "MC_ProjectManagement_LH")
APP_USER = os.getenv("APP_USER", "unknown")

ONELAKE_DFS = "https://onelake.dfs.fabric.microsoft.com"
ABFSS_BASE = os.getenv("ABFSS_BASE",
    f"abfss://{WORKSPACE_NAME}@onelake.dfs.fabric.microsoft.com/{LAKEHOUSE_NAME}.Lakehouse")

# ── Token Cache ───────────────────────────────────────────────────────
_token_cache = {}


def _get_token(scope):
    """Get Azure AD token with caching."""
    cached = _token_cache.get(scope)
    if cached and cached["expires"] > time.time():
        return cached["token"]

    resp = requests.post(
        f"https://login.microsoftonline.com/{FABRIC_TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": FABRIC_CLIENT_ID,
            "client_secret": FABRIC_CLIENT_SECRET,
            "scope": scope,
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    _token_cache[scope] = {
        "token": data["access_token"],
        "expires": time.time() + data.get("expires_in", 3600) - 60,
    }
    return data["access_token"]


def _storage_token():
    return _get_token("https://storage.azure.com/.default")


def _storage_options():
    return {"bearer_token": _storage_token(), "use_fabric_endpoint": "true"}


def _storage_headers():
    return {"Authorization": f"Bearer {_storage_token()}"}


def _onelake_base():
    return f"{ONELAKE_DFS}/{WORKSPACE_NAME}/{LAKEHOUSE_NAME}.Lakehouse"


# ═══════════════════════════════════════════════════════════════════════
#  READ — Delta Lake (with fallback to parquet in Files/)
# ═══════════════════════════════════════════════════════════════════════

def read_table(table_name):
    """
    Read a table. Tries Delta Lake first, then falls back to parquet.
    Schema folder configurable via FABRIC_SCHEMA env var (default: dbo).
    """
    schema = os.getenv("FABRIC_SCHEMA", "dbo")
    # Try Delta Lake first
    try:
        from deltalake import DeltaTable
        delta_path = f"{ABFSS_BASE}/Tables/{schema}/{table_name}"
        print(f"DEBUG read_table: trying Delta path = {delta_path}")
        dt = DeltaTable(delta_path, storage_options=_storage_options())
        df = dt.to_pandas()
        print(f"DEBUG read_table: {table_name} = {len(df)} rows from Delta")
        return df
    except Exception as e:
        print(f"DEBUG read_table: Delta FAILED for {table_name}: {e}")

    # Fallback to parquet in Files/
    try:
        url = f"{_onelake_base()}/Files/app_data/{table_name}.parquet"
        print(f"DEBUG read_table: trying parquet = {url}")
        resp = requests.get(url, headers=_storage_headers(), timeout=60)
        print(f"DEBUG read_table: parquet status = {resp.status_code}")
        if resp.status_code == 200:
            df = pd.read_parquet(io.BytesIO(resp.content))
            return df
    except Exception as e:
        print(f"DEBUG read_table: Parquet FAILED for {table_name}: {e}")

    return pd.DataFrame()


# ═══════════════════════════════════════════════════════════════════════
#  WRITE — Delta Lake (with fallback to parquet)
# ═══════════════════════════════════════════════════════════════════════

def write_table(table_name, df):
    """
    Write a DataFrame as a Delta Lake table to Tables/dbo/{table_name}.
    Falls back to parquet in Files/ if Delta write fails.
    """
    # Try Delta Lake write
    try:
        from deltalake import write_deltalake
        delta_path = f"{ABFSS_BASE}/Tables/dbo/{table_name}"
        write_deltalake(
            delta_path, df,
            mode="overwrite",
            storage_options=_storage_options(),
        )
        logger.info("Wrote %d rows to Delta table %s", len(df), table_name)
        return
    except Exception as e:
        logger.warning("Delta write failed for %s: %s, falling back to parquet", table_name, e)

    # Fallback to parquet
    _write_parquet_fallback(table_name, df)


def _write_parquet_fallback(table_name, df):
    """Write as parquet to Files/app_data/ (fallback)."""
    url = f"{_onelake_base()}/Files/app_data/{table_name}.parquet"
    headers = _storage_headers()
    buf = io.BytesIO()
    df.to_parquet(buf, index=False, engine="pyarrow")
    parquet_bytes = buf.getvalue()

    try:
        requests.delete(url, headers=headers, timeout=15)
    except Exception:
        pass

    requests.put(url, headers={**headers, "Content-Length": "0"},
        params={"resource": "file"}, timeout=15).raise_for_status()
    requests.patch(url, headers={**headers, "Content-Length": str(len(parquet_bytes)),
        "Content-Type": "application/octet-stream"},
        params={"action": "append", "position": "0"},
        data=parquet_bytes, timeout=30).raise_for_status()
    requests.patch(url, headers=headers,
        params={"action": "flush", "position": str(len(parquet_bytes))},
        timeout=15).raise_for_status()
    logger.info("Wrote %d rows to parquet %s (fallback)", len(df), table_name)


def append_row(table_name, row_dict):
    """Append a single row. Ensures consistent schema across all rows."""
    existing = read_table(table_name)
    new_row = pd.DataFrame([row_dict])

    if existing.empty:
        combined = new_row
    else:
        # Ensure both DataFrames have the same columns
        all_cols = list(dict.fromkeys(list(existing.columns) + list(new_row.columns)))
        for col in all_cols:
            if col not in existing.columns:
                existing[col] = ""
            if col not in new_row.columns:
                new_row[col] = ""
        # Reorder to match
        new_row = new_row[all_cols]
        existing = existing[all_cols]
        combined = pd.concat([existing, new_row], ignore_index=True)

    # Fill any NaN with empty string to prevent schema issues
    combined = combined.fillna("")
    write_table(table_name, combined)
    return len(combined)


def update_table(table_name, df):
    """Replace entire table."""
    write_table(table_name, df)


# ═══════════════════════════════════════════════════════════════════════
#  PARALLEL READ (for reading multiple tables at once)
# ═══════════════════════════════════════════════════════════════════════

from concurrent.futures import ThreadPoolExecutor, as_completed

def read_tables_parallel(table_names, max_workers=5):
    """Read multiple tables in parallel. Returns dict of {table_name: DataFrame}.

    5 tables in ~1-2 sec instead of ~5 sec sequential.

    Args:
        table_names: List of table names to read.
        max_workers: Max concurrent reads (default 5).

    Returns:
        dict: {table_name: DataFrame}

    Example:
        tables = read_tables_parallel(["Study", "Country", "Site"])
        df_study = tables["Study"]
        df_country = tables["Country"]
    """
    results = {}

    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(read_table, name): name for name in table_names}
        for future in as_completed(futures):
            name = futures[future]
            try:
                results[name] = future.result()
                logger.info("Parallel read %s: %d rows", name, len(results[name]))
            except Exception as e:
                logger.error("Parallel read failed for %s: %s", name, e)
                results[name] = pd.DataFrame()

    return results


def read_tables_parallel_cached(table_names, cache_dict=None, cache_ts_dict=None, cache_ttl=300, max_workers=5):
    """Read multiple tables in parallel with caching.

    Only reads tables whose cache has expired. Returns dict of {table_name: DataFrame}.

    Args:
        table_names: List of table names.
        cache_dict: Dict to store cached DataFrames (pass your app's cache).
        cache_ts_dict: Dict to store cache timestamps.
        cache_ttl: Cache time-to-live in seconds (default 300 = 5 min).
        max_workers: Max concurrent reads.

    Returns:
        dict: {table_name: DataFrame}

    Example:
        _cache = {}
        _cache_ts = {}
        tables = read_tables_parallel_cached(
            ["Study", "Country", "Site"],
            cache_dict=_cache, cache_ts_dict=_cache_ts
        )
    """
    import time
    now = time.time()

    if cache_dict is None:
        cache_dict = {}
    if cache_ts_dict is None:
        cache_ts_dict = {}

    # Determine which tables need refreshing
    stale = []
    for name in table_names:
        if name not in cache_dict or (now - cache_ts_dict.get(name, 0)) >= cache_ttl:
            stale.append(name)

    # Read stale tables in parallel
    if stale:
        fresh = read_tables_parallel(stale, max_workers)
        for name, df in fresh.items():
            cache_dict[name] = df
            cache_ts_dict[name] = now

    # Return all requested tables from cache
    return {name: cache_dict.get(name, pd.DataFrame()).copy() for name in table_names}


# ═══════════════════════════════════════════════════════════════════════
#  CONNECTION TEST
# ═══════════════════════════════════════════════════════════════════════

def test_connection():
    """Test OneLake connectivity."""
    try:
        url = f"{_onelake_base()}/Files"
        resp = requests.get(url, headers=_storage_headers(),
            params={"resource": "filesystem", "recursive": "false"}, timeout=15)
        return resp.status_code == 200
    except Exception as e:
        logger.error("Connection test failed: %s", e)
        return False


if __name__ == "__main__":
    print(f"Workspace:  {WORKSPACE_NAME}")
    print(f"Lakehouse:  {LAKEHOUSE_NAME}")
    print(f"Delta path: {ABFSS_BASE}/Tables/dbo/")
    print()
    if test_connection():
        print("SUCCESS - Connected!")
        df = read_table("Lookups")
        if not df.empty:
            print(f"Lookups: {len(df)} rows")
        else:
            print("Lookups: empty or not created yet")
    else:
        print("FAILED")
