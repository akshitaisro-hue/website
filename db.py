import pyodbc
from contextlib import contextmanager

CONN_STR = (
    "DSN=python;"
    "SERVER=localhost;"
    "DATABASE=Pahal-1;"
    "UID=root;"
    "PWD=;"
)

KEY_COLS_TC = ["src_card_type", "advantek_tc_address", "word_pin"]  # card type + ip + pin
KEY_COLS_TM = ["rt_addres", "sa_addres", "tm_channel_position", "advantek_tm_ip","tm_length"]  # real col names (typos preserved)


@contextmanager
def get_conn():
    conn = pyodbc.connect(CONN_STR)
    try:
        yield conn
    finally:
        conn.close()


# ---------------- PROJECTS ----------------

def list_projects():
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT p_id, project_name FROM pahal_projects ORDER BY p_id")
        return [{"p_id": r.p_id, "project_name": r.project_name} for r in cur.fetchall()]


def get_project_by_name(project_name):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT p_id, project_name FROM pahal_projects WHERE project_name = ?", project_name)
        row = cur.fetchone()
        return {"p_id": row.p_id, "project_name": row.project_name} if row else None


def get_project_by_id(p_id):
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT p_id, project_name FROM pahal_projects WHERE p_id = ?", p_id)
        row = cur.fetchone()
        return {"p_id": row.p_id, "project_name": row.project_name} if row else None


def create_project(project_name):
    if get_project_by_name(project_name):
        return None  # 409
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("INSERT INTO pahal_projects (project_name) VALUES (?)", project_name)
        conn.commit()
        cur.execute("SELECT p_id FROM pahal_projects WHERE project_name = ?", project_name)
        return {"p_id": cur.fetchone().p_id, "project_name": project_name}


# ---------------- TC / TM RECORDS ----------------

def build_match_key(row: dict, table: str) -> str:
    assert table in ("all_tc", "all_tm")
    cols = KEY_COLS_TC if table == "all_tc" else KEY_COLS_TM
    return "|".join(str(row.get(c, "")) for c in cols)


def fetch_records(proj_id, table):
    """proj_id: python var name only, matches SQL column p_id (project FK)"""
    assert table in ("all_tc", "all_tm")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute(f"SELECT * FROM {table} WHERE p_id = ?", proj_id)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]


def insert_records(proj_id, table, rows: list[dict], columns: list[str]):
    assert table in ("all_tc", "all_tm")
    if not rows:
        return 0
    col_list = ", ".join(["p_id", "match_key"] + columns)
    placeholders = ", ".join(["?"] * (len(columns) + 2))
    sql = f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})"
    with get_conn() as conn:
        cur = conn.cursor()
        for row in rows:
            values = [proj_id, row["match_key"]] + [row.get(c) for c in columns]
            cur.execute(sql, values)
        conn.commit()
    return len(rows)


def diff_table(proj_id, table, new_rows: list[dict]):
    """single-table diff, used by check_sync to build combined tc+tm response"""
    existing = fetch_records(proj_id, table)
    existing_keys = {r["match_key"] for r in existing}
    new_keys = {r["match_key"] for r in new_rows}

    if not existing:
        return {"status": "no_records", "matched": 0, "new": new_rows, "conflict": []}
    if new_keys == existing_keys:
        return {"status": "all_match", "matched": len(existing_keys), "new": [], "conflict": []}
    if new_keys.isdisjoint(existing_keys):
        # "new" stays populated so insert_all can force-insert; "conflict" = same rows, for UI diff preview only
        return {"status": "conflict", "matched": 0, "new": new_rows, "conflict": new_rows}

    to_append = [r for r in new_rows if r["match_key"] not in existing_keys]
    matched = len(new_keys & existing_keys)
    return {"status": "some_new", "matched": matched, "new": to_append, "conflict": []}


def check_sync(proj_id, tc_rows, tm_rows):
    """
    combines tc_records + tm_records diffs into single response matching:
    {status, matched_count, new_count, conflict_rows, new_rows:{tc,tm}}
    overall status priority: conflict > some_new > all_match > no_records
    """
    tc_diff = diff_table(proj_id, "all_tc", tc_rows)
    tm_diff = diff_table(proj_id, "all_tm", tm_rows)

    statuses = {tc_diff["status"], tm_diff["status"]}
    if "conflict" in statuses:
        overall = "conflict"
    elif "some_new" in statuses:
        overall = "some_new"
    elif "no_records" in statuses:
        overall = "no_records"
    else:
        overall = "all_match"

    return {
        "status": overall,
        "matched_count": tc_diff["matched"] + tm_diff["matched"],
        "new_count": len(tc_diff["new"]) + len(tm_diff["new"]),
        "conflict_rows": tc_diff["conflict"] + tm_diff["conflict"],
        "new_rows": {"tc": tc_diff["new"], "tm": tm_diff["new"]},
    }
