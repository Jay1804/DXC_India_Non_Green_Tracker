import io
import os
from datetime import datetime, timedelta

import pandas as pd
import pymysql
import streamlit as st
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

TITLE = "DXC INDIA - NON GREEN TRACKER"  # page title only

HEADERS = [
    "candidate name", "client name", "ars. no", "process name", "location",
    "received date", "candidate id", "requisition id", "candidate_email",
    "case_status", "check_created_on", "check unique name", "check status",
    "check severity", "check closure date", "verification_source_name",
    "latest closure comments", "pre hire or post hire", "package name",
    "latest insuff remarks",
]
DATE_COLS = ["received date", "check_created_on", "check closure date"]

QUERY = """
WITH final_output_cte AS (
    SELECT ecc.case_check_id AS CTE_id,
           GROUP_CONCAT(DISTINCT
               CASE WHEN ecc.family_id = 5 THEN emc.company_name
                    WHEN ecc.family_id = 4 THEN
                        CASE WHEN emei.institute_type = 2 THEN emei.institute_name
                             ELSE CONCAT(emei.institute_name, '/', COALESCE(NULLIF(emei1.institute_name, ''), ''))
                        END
                    ELSE emc1.city_name
               END SEPARATOR ', ') AS final_output
    FROM ec_case_checks ecc
    LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id = eccvs.case_check_id
    LEFT JOIN ec_master_company emc ON eccvs.org_id = emc.company_id
    LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id = emei.institute_id
    LEFT JOIN ec_master_educational_institute emei1 ON emei.university_id = emei1.institute_id
    LEFT JOIN ec_master_city emc1 ON eccvs.org_id = emc1.city_id
    WHERE ecc.check_closure_date BETWEEN %(from_dt)s AND %(to_dt)s
    GROUP BY ecc.case_check_id
)
SELECT
    CONCAT(COALESCE(ecc1.first_name, ''), ' ', COALESCE(ecc1.middle_name, ''), ' ', COALESCE(ecc1.last_name, '')) AS `candidate name`,
    emc.company_name AS `client name`,
    ecm.case_ars_no AS `ars. no`,
    ecp.process_name AS `process name`,
    emcl.office_name AS `location`,
    ecm.received_date AS `received date`,
    ecf.CASE_FLEX_FIELD1 AS `candidate id`,
    ecf.CASE_FLEX_FIELD10 AS `requisition id`,
    eccd.candidate_email AS `candidate_email`,
    fn_case_status(ecm.case_status) AS `case_status`,
    ecc.check_created_on AS `check_created_on`,
    ec.check_name AS `check unique name`,
    fn_check_status(ecc.check_status) AS `check status`,
    ecc.check_severity AS `check severity`,
    ecc.check_closure_date AS `check closure date`,
    foc.final_output AS `verification_source_name`,
    ecc.closure_comments AS `latest closure comments`,
    ecf.CASE_FLEX_FIELD17 AS `pre hire or post hire`,
    ecf.CASE_FLEX_FIELD13 AS `package name`,
    ecc.insuff_remarks AS `latest insuff remarks`
FROM ec_case_master ecm
LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_case_candidate_data eccd ON ecm.case_id = eccd.case_id
LEFT JOIN ec_client_process ecp ON ecm.client_id = ecp.client_id AND ecm.process_id = ecp.process_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id = ecm.candidate_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id = emcl.office_id
LEFT JOIN ec_case_fields ecf ON ecm.case_id = ecf.case_id
LEFT JOIN ec_checks ec ON ecc.check_id = ec.check_id
LEFT JOIN ec_client ec1 ON ecm.client_id = ec1.client_id
LEFT JOIN final_output_cte foc ON ecc.case_check_id = foc.CTE_id
WHERE ec1.client_external_id IN (2307)
  AND ecm.case_status NOT IN (8)
  AND ecc.check_severity NOT LIKE 'Green%%'
  AND ecc1.first_name NOT LIKE '%%dummy%%'
  AND ecc1.first_name NOT LIKE '%%test%%'
  AND ecc.check_id NOT IN (193)
  AND ecc.check_closure_date BETWEEN %(from_dt)s AND %(to_dt)s
ORDER BY ecc.check_closure_date, ecm.case_ars_no
"""


def get_window(now: datetime):
    """Return (from_dt, to_dt) or None on Sat/Sun."""
    wd = now.weekday()  # Mon=0 ... Sun=6
    if wd >= 5:
        return None
    back = 3 if wd == 0 else 1
    start = (now - timedelta(days=back)).replace(hour=0, minute=0, second=0, microsecond=0)
    end = now.replace(hour=23, minute=59, second=59, microsecond=0)
    return start, end


def fetch(from_dt, to_dt) -> pd.DataFrame:
    conn = pymysql.connect(
        host=os.environ["DB_HOST"],
        port=int(os.environ.get("DB_PORT", 3306)),
        user=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        database=os.environ["DB_NAME"],
        charset="utf8mb4",
        connect_timeout=20,
    )
    try:
        with conn.cursor() as cur:
            cur.execute(QUERY, {"from_dt": from_dt, "to_dt": to_dt})
            rows = cur.fetchall()
    finally:
        conn.close()
    df = pd.DataFrame(rows, columns=HEADERS)
    for c in DATE_COLS:
        # real dates (no time part) so Excel filters show the Year > Month > Day hierarchy
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    df = df.astype(object)
    return df.where(df.notna(), "")


def _as_number(v):
    """Requisition id as a real number (avoids Excel's 'number stored as text' warning)."""
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return v


def to_excel(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "sheet1"
    side = Side(style="thin", color="000000")
    border = Border(left=side, right=side, top=side, bottom=side)
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="left")
        c.border = border
    for r, row in enumerate(df.itertuples(index=False), 2):
        for i, v in enumerate(row, 1):
            if HEADERS[i - 1] == "requisition id":
                v = _as_number(v)
            cell = ws.cell(row=r, column=i, value=v)
            cell.border = border
            if HEADERS[i - 1] in DATE_COLS and v != "":
                cell.number_format = "mmm d, yyyy"
                cell.alignment = Alignment(horizontal="left")
    for i, h in enumerate(HEADERS, 1):
        width = max([len(h)] + [min(len(str(v)), 50) for v in df[h]]) + 2 if len(df) else len(h) + 2
        ws.column_dimensions[get_column_letter(i)].width = width
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


st.set_page_config(page_title="DXC India - Non Green Tracker", layout="wide")
st.title(TITLE.title())

now = datetime.now()
window = get_window(now)
today_str = now.strftime("%d-%m-%Y")
file_name = f"DXC_INDIA_-_NON_GREEN_Tracker {today_str}.xlsx"

if window is None:
    st.info(f"Today is {now.strftime('%A')} - no tracker is generated on Saturday and Sunday.")
    st.stop()

from_dt, to_dt = window
st.write(f"**{now.strftime('%A')}** — check closure date from `{from_dt}` to `{to_dt}`")

if st.button("Generate Tracker", type="primary"):
    try:
        with st.spinner("Fetching data from live DB..."):
            st.session_state["df"] = fetch(from_dt, to_dt)
    except Exception as e:  # noqa: BLE001
        st.error(f"Database error: {e}")

df = st.session_state.get("df")
if df is not None:
    st.success(f"{len(df)} records")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button(
        "Download Excel",
        data=to_excel(df),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
