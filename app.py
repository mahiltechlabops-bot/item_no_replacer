import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json
import os
import re
import numpy as np
from datetime import datetime
from io import BytesIO

HISTORY_FILE = "change_history.json"

st.set_page_config(page_title="Item No Replacer", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem; }
    .stDataFrame { border: 1px solid #e0e0e0; border-radius: 8px; }
    div[data-testid="stMetricValue"] { font-size: 1.1rem; }
    .change-row { padding: 6px 0; border-bottom: 1px solid #f0f0f0; font-size: 0.85rem; }
    .tag-inactive { background:#fff3cd; color:#856404; padding:2px 8px; border-radius:4px; font-size:0.78rem; }
    .tag-replaced { background:#d1e7dd; color:#0f5132; padding:2px 8px; border-radius:4px; font-size:0.78rem; }
</style>
""", unsafe_allow_html=True)

# ── session state init ──────────────────────────────────────────────────────
def init_state():
    defaults = {
        "df_a": None,
        # Snapshot of Excel A at the time it was loaded; used for change highlighting.
        "df_a_original": None,
        "df_b": None,
        "file_a_error": None,
        "file_b_error": None,
        "file_a_empty_message": None,
        "file_b_empty_message": None,
        "history_stack": [],   # list of df_a snapshots (undo)
        "redo_stack": [],
        "change_log": [],      # [{ts, productId, name, old_item, new_item, item_name}]
        "changed_product_ids": [],  # list of productIds modified in current Excel A session
        "selected_row": None,
        "file_a_name": "updated_excel_a.xlsx",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── persistence helpers ─────────────────────────────────────────────────────
def load_history_file():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r") as f:
            return json.load(f)
    return []

def save_history_file(log):
    with open(HISTORY_FILE, "w") as f:
        json.dump(log, f, indent=2)

if not st.session_state.change_log and os.path.exists(HISTORY_FILE):
    st.session_state.change_log = load_history_file()

# ── helpers ─────────────────────────────────────────────────────────────────
def df_to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def read_excel_file(file_obj, label):
    try:
        df = pd.read_excel(file_obj)
    except Exception as exc:
        return None, f"{label} could not be opened. Please upload a valid Excel file. Details: {exc}"

    df.columns = [str(c).strip() for c in df.columns]
    return df, None

def save_excel_file(df, out_path):
    try:
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        df.to_excel(out_path, index=False, engine="openpyxl")
    except PermissionError:
        return (
            False,
            "Could not save the Excel file because it is open in another program. "
            "Close it and try again."
        )
    except Exception as exc:
        return False, f"Could not save the Excel file. Details: {exc}"

    return True, None

def get_modified_product_list():
    current_df = st.session_state.df_a
    original_df = st.session_state.df_a_original

    if (
        current_df is None
        or original_df is None
        or "productId" not in current_df.columns
        or "item_No" not in current_df.columns
        or "productId" not in original_df.columns
        or "item_No" not in original_df.columns
    ):
        return pd.DataFrame(columns=current_df.columns if current_df is not None else [])

    original_item_map = (
        original_df[["productId", "item_No"]]
        .assign(_pid=lambda d: d["productId"].astype(str))
        .set_index("_pid")["item_No"]
    )
    current_pid = current_df["productId"].astype(str)
    changed_mask = (
        current_pid.map(original_item_map).fillna("").astype(str).str.strip()
        != current_df["item_No"].fillna("").astype(str).str.strip()
    )
    return current_df.loc[changed_mask].copy()

def build_change_signature_series(df):
    if df is None or df.empty:
        return pd.Series(dtype="string")

    if "productId" in df.columns and "item_No" in df.columns:
        return (
            df["productId"].fillna("").astype(str).str.strip()
            + "||"
            + df["item_No"].fillna("").astype(str).str.strip()
        )

    return df.fillna("").astype(str).agg("||".join, axis=1)

def build_product_id_series(df):
    if df is None or df.empty or "productId" not in df.columns:
        return pd.Series(dtype="string")

    return df["productId"].fillna("").astype(str).str.strip()

def merge_with_existing_modified_list(new_df, out_path):
    existing_df = pd.DataFrame(columns=new_df.columns if new_df is not None else [])

    if os.path.exists(out_path):
        try:
            existing_df = pd.read_excel(out_path)
        except Exception as exc:
            return None, None, (
                "Could not open the existing modified product list Excel file. "
                f"Details: {exc}"
            )

    empty_like_existing = pd.DataFrame(columns=existing_df.columns)

    if new_df is None or new_df.empty:
        return existing_df.reset_index(drop=True), empty_like_existing.copy(), empty_like_existing.copy(), existing_df.reset_index(drop=True), None

    new_signatures = build_change_signature_series(new_df)
    existing_signatures = set(build_change_signature_series(existing_df).tolist()) if not existing_df.empty else set()
    exact_match_mask = new_signatures.isin(existing_signatures)

    existing_product_ids = set(build_product_id_series(existing_df).tolist()) if not existing_df.empty else set()
    new_product_ids = build_product_id_series(new_df)
    conflicting_mask = new_product_ids.isin(existing_product_ids) & ~exact_match_mask

    overwrite_df = new_df.loc[conflicting_mask].copy().reset_index(drop=True)
    append_df = new_df.loc[~exact_match_mask & ~conflicting_mask].copy().reset_index(drop=True)

    append_preview_df = existing_df.copy()
    if not append_df.empty:
        append_preview_df = pd.concat([append_preview_df, append_df], ignore_index=True, sort=False)

    overwrite_preview_df = append_preview_df.copy()
    if not overwrite_df.empty:
        overwrite_product_ids = set(build_product_id_series(overwrite_df).tolist())
        if overwrite_product_ids and "productId" in overwrite_preview_df.columns:
            keep_mask = ~build_product_id_series(overwrite_preview_df).isin(overwrite_product_ids)
            overwrite_preview_df = overwrite_preview_df.loc[keep_mask].copy()
        overwrite_preview_df = pd.concat([overwrite_preview_df, overwrite_df], ignore_index=True, sort=False)

    return (
        append_preview_df.reset_index(drop=True),
        append_df.reset_index(drop=True),
        overwrite_df.reset_index(drop=True),
        overwrite_preview_df.reset_index(drop=True),
        None,
    )

def scroll_excel_a_to_selected_row(selected_index):
    if selected_index is None or selected_index < 0:
        return

    # Restore the Excel A table scroll position after Streamlit reruns.
    components.html(
        f"""
        <script>
        const targetIndex = {int(selected_index)};
        const rowOffset = 2;

        function findScrollableElement(root) {{
          if (!root) return null;
          const queue = [root];
          let best = null;
          while (queue.length) {{
            const node = queue.shift();
            if (!(node instanceof window.parent.HTMLElement)) continue;
            const style = window.parent.getComputedStyle(node);
            const canScroll = node.scrollHeight > node.clientHeight + 20;
            const overflowY = style.overflowY === "auto" || style.overflowY === "scroll";
            if (canScroll && overflowY) {{
              if (!best || node.scrollHeight > best.scrollHeight) {{
                best = node;
              }}
            }}
            queue.push(...node.children);
          }}
          return best;
        }}

        function scrollToSelectedRow(attemptsLeft = 20) {{
          const doc = window.parent.document;
          const tables = doc.querySelectorAll('[data-testid="stDataFrame"]');
          const table = tables.length ? tables[0] : null;
          const scrollable = findScrollableElement(table);

          if (!scrollable) {{
            if (attemptsLeft > 0) {{
              window.setTimeout(() => scrollToSelectedRow(attemptsLeft - 1), 150);
            }}
            return;
          }}

          const firstRow = table.querySelector('[role="row"]');
          const rowHeight = firstRow && firstRow.getBoundingClientRect().height
            ? firstRow.getBoundingClientRect().height
            : 35;
          scrollable.scrollTop = Math.max((targetIndex - rowOffset) * rowHeight, 0);
        }}

        window.setTimeout(() => scrollToSelectedRow(), 120);
        </script>
        """,
        height=0,
    )

def push_undo(df):
    st.session_state.history_stack.append(df.copy())
    st.session_state.redo_stack.clear()

def do_undo():
    if st.session_state.history_stack:
        st.session_state.redo_stack.append(st.session_state.df_a.copy())
        st.session_state.df_a = st.session_state.history_stack.pop()
        st.session_state.selected_row = None

def do_redo():
    if st.session_state.redo_stack:
        st.session_state.history_stack.append(st.session_state.df_a.copy())
        st.session_state.df_a = st.session_state.redo_stack.pop()
        st.session_state.selected_row = None

def filter_b(brand, subcategory):
    df = st.session_state.df_b
    brand_col = next((c for c in df.columns if "brand" in c.lower()), None)
    group_col = next((c for c in df.columns if "group" in c.lower() or "itemgroup" in c.lower().replace("_","")), None)
    if not brand_col or not group_col:
        return df

    def norm_text(s: str) -> str:
        s = str(s).lower()
        # Keep alphanumerics, turn everything else into whitespace, then collapse.
        s = re.sub(r"[^a-z0-9]+", " ", s)
        return " ".join(s.split())

    brand_n = norm_text(brand)
    subcat_n = norm_text(subcategory)

    brand_series = df[brand_col].astype(str).str.lower()
    group_series = df[group_col].astype(str).str.lower()

    # Brand: keep the previous behavior (full-phrase contains or exact match).
    brand_mask = (
        brand_series.str.contains(brand_n, na=False) |
        (brand_series == brand_n)
    ) if brand_n else pd.Series([True] * len(df), index=df.index)

    # Subcategory: Excel A often uses longer phrases (e.g. "Bath soap")
    # while Excel B may store a shorter value (e.g. "soap").
    # So we match if ANY token from Excel A's subcategory appears in Excel B.
    tokens = [t for t in subcat_n.split(" ") if t]
    if tokens:
        token_mask = pd.Series([False] * len(df), index=df.index)
        for t in tokens:
            token_mask = token_mask | group_series.str.contains(t, na=False)
        phrase_mask = (
            group_series.str.contains(subcat_n, na=False) |
            (group_series == subcat_n)
        )
        group_mask = token_mask | phrase_mask
    else:
        group_mask = (
            group_series.str.contains(subcat_n, na=False) |
            (group_series == subcat_n)
        )

    mask = brand_mask & group_mask
    return df[mask].reset_index(drop=True)

def apply_replacement(prod_id, new_item_no, new_item_name, brand, subcategory, a_row, b_row):
    df = st.session_state.df_a
    idx = df.index[df["productId"] == prod_id][0]
    old_item_raw = df.at[idx, "item_No"]
    push_undo(df)

    def parse_item_no(v):
        v_str = "" if v is None else str(v).strip()
        if v_str == "" or v_str.lower() == "nan":
            return ""
        if re.fullmatch(r"-?\d+", v_str):
            return int(v_str)
        return v_str

    new_item_parsed = parse_item_no(new_item_no)
    old_item_parsed = parse_item_no(old_item_raw)

    st.session_state.df_a.at[idx, "item_No"] = new_item_parsed

    # Track which Excel A rows were modified (for UI highlighting).
    if prod_id not in st.session_state.changed_product_ids:
        st.session_state.changed_product_ids.append(prod_id)

    def json_safe(v):
        # Convert numpy/pandas scalars to native python for json serialization.
        if pd.isna(v):
            return ""
        if isinstance(v, (np.generic,)):
            return v.item()
        if isinstance(v, (pd.Timestamp, datetime)):
            return str(v)
        return v

    a_row_safe = {k: json_safe(v) for k, v in (a_row or {}).items()}
    b_row_safe = {k: json_safe(v) for k, v in (b_row or {}).items()}

    entry = {
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "productId": prod_id,
        "name": df.at[idx, "name"],
        "brand": brand,
        "subcategory": subcategory,
        "a_row": a_row_safe,
        "b_row": b_row_safe,
        "old_item": old_item_parsed,
        "new_item": new_item_parsed,
        "item_name": new_item_name,
    }
    st.session_state.change_log.insert(0, entry)
    save_history_file(st.session_state.change_log)
    st.session_state.selected_row = None

# ── layout ───────────────────────────────────────────────────────────────────
st.title("🔄 Item No Replacer")

# Upload row
col_a, col_b = st.columns(2)
with col_a:
    file_a = st.file_uploader("Upload Excel A (products)", type=["xlsx", "xls"], key="up_a")
    if file_a:
        with st.spinner("Loading Excel A..."):
            raw, file_a_error = read_excel_file(file_a, "Excel A")

        st.session_state.file_a_error = file_a_error
        st.session_state.file_a_empty_message = None

        if file_a_error:
            st.session_state.df_a = None
            st.session_state.df_a_original = None
            st.error(file_a_error)
        elif "status" not in raw.columns:
            st.session_state.df_a = None
            st.session_state.df_a_original = None
            st.session_state.file_a_error = "Excel A is missing the required `status` column."
            st.error(st.session_state.file_a_error)
        else:
            inactive = raw[raw["status"].astype(str).str.lower() == "inactive"].reset_index(drop=True)
            if st.session_state.df_a is None or file_a.name != st.session_state.file_a_name:
                st.session_state.df_a = inactive.copy()
                st.session_state.df_a_original = inactive.copy()
                st.session_state.file_a_name = file_a.name
                st.session_state.history_stack.clear()
                st.session_state.redo_stack.clear()
                st.session_state.changed_product_ids = []

            if raw.empty:
                st.session_state.file_a_empty_message = "Excel A is empty."
                st.warning(st.session_state.file_a_empty_message)
            elif inactive.empty:
                st.session_state.file_a_empty_message = "Excel A loaded, but no inactive records were found."
                st.warning(st.session_state.file_a_empty_message)
            else:
                st.success(f"{len(inactive)} inactive records loaded")

with col_b:
    file_b = st.file_uploader("Upload Excel B (item master)", type=["xlsx", "xls"], key="up_b")
    if file_b:
        with st.spinner("Loading Excel B..."):
            df_b, file_b_error = read_excel_file(file_b, "Excel B")

        st.session_state.file_b_error = file_b_error
        st.session_state.file_b_empty_message = None

        if file_b_error:
            st.session_state.df_b = None
            st.error(file_b_error)
        else:
            st.session_state.df_b = df_b
            if df_b.empty:
                st.session_state.file_b_empty_message = "Excel B is empty."
                st.warning(st.session_state.file_b_empty_message)
            else:
                st.success(f"{len(df_b)} records loaded")

if st.session_state.df_a is None or st.session_state.df_b is None:
    st.info("Please upload both Excel A and Excel B to get started.")
    st.stop()

if st.session_state.file_a_error or st.session_state.file_b_error:
    st.stop()

if st.session_state.df_a.empty:
    st.info(st.session_state.file_a_empty_message or "Excel A has no rows to display.")
    st.stop()

if st.session_state.df_b.empty:
    st.info(st.session_state.file_b_empty_message or "Excel B has no rows to match against.")
    st.stop()

st.divider()

# Toolbar
tb1, tb2, tb3, tb4, tb5 = st.columns([1,1,1,1,4])
with tb1:
    if st.button("↩ Undo", disabled=not st.session_state.history_stack, use_container_width=True):
        do_undo()
        st.rerun()
with tb2:
    if st.button("↪ Redo", disabled=not st.session_state.redo_stack, use_container_width=True):
        do_redo()
        st.rerun()
with tb3:
    modified_df = get_modified_product_list()
    project_root = os.path.dirname(os.path.abspath(__file__))
    records_dir = os.path.join(project_root, "files")
    out_path = os.path.join(records_dir, "modified_product_list.xlsx")
    save_df, pending_save_df, overwrite_df, overwrite_save_df, save_df_error = merge_with_existing_modified_list(modified_df, out_path)
    excel_bytes = df_to_excel_bytes(save_df) if save_df_error is None else None
    # Append only brand-new productIds. Existing productIds require explicit overwrite.
    if save_df_error:
        st.error(save_df_error)
    elif modified_df.empty:
        st.caption("No modified rows to save yet.")
    elif pending_save_df.empty and overwrite_df.empty:
        st.caption("All current modified rows are already appended in the saved file.")
    elif not overwrite_df.empty:
        overwrite_ids = ", ".join(overwrite_df["productId"].astype(str).head(5).tolist())
        if len(overwrite_df) > 5:
            overwrite_ids += ", ..."
        st.warning(
            f"{len(overwrite_df)} row(s) already exist in the saved file by productId. "
            f"Use Overwrite to replace them. ProductIds: {overwrite_ids}"
        )

    if st.button("💾 Save", use_container_width=True, disabled=save_df_error is not None or pending_save_df.empty):
        ok, save_error = save_excel_file(save_df, out_path)
        if ok:
            st.success(
                f"Appended {len(pending_save_df)} row(s) to "
                f"`{os.path.join('files', 'modified_product_list.xlsx')}`"
            )
        else:
            st.error(save_error)
    if st.button("Overwrite Existing ProductIds", use_container_width=True, disabled=save_df_error is not None or overwrite_df.empty):
        ok, save_error = save_excel_file(overwrite_save_df, out_path)
        if ok:
            st.success(
                f"Overwrote {len(overwrite_df)} existing productId row(s) in "
                f"`{os.path.join('files', 'modified_product_list.xlsx')}`"
            )
        else:
            st.error(save_error)
    st.download_button(
        "Download modified_product_list.xlsx",
        data=excel_bytes,
        file_name="modified_product_list.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=save_df_error is not None or save_df is None or save_df.empty,
        use_container_width=True,
    )
with tb4:
    if st.button("🗑 Clear Log", use_container_width=True):
        st.session_state.change_log = []
        save_history_file([])
        st.rerun()

st.divider()

# Main panels
left, right = st.columns([3, 2])

with left:
    st.subheader("Excel A — Inactive Records")

    df_display = st.session_state.df_a.copy()
    # Enforce single selection: only the currently-selected row has Select=True.
    current_pid = None
    if st.session_state.selected_row and isinstance(st.session_state.selected_row, dict):
        current_pid = st.session_state.selected_row.get("productId")
    if current_pid is not None and "productId" in df_display.columns:
        df_display.insert(
            0,
            "Select",
            df_display["productId"].astype(str) == str(current_pid),
        )
    else:
        df_display.insert(0, "Select", False)
    # Mark rows where item_No differs from the original Excel A snapshot.
    if (
        st.session_state.df_a_original is not None
        and "productId" in df_display.columns
        and "item_No" in df_display.columns
        and "productId" in st.session_state.df_a_original.columns
        and "item_No" in st.session_state.df_a_original.columns
    ):
        orig_map = (
            st.session_state.df_a_original[["productId", "item_No"]]
            .assign(_pid=lambda d: d["productId"].astype(str))
            .set_index("_pid")["item_No"]
        )
        pid_str = df_display["productId"].astype(str)
        orig_item = pid_str.map(orig_map)
        changed = (
            orig_item.fillna("").astype(str).str.strip()
            != df_display["item_No"].fillna("").astype(str).str.strip()
        )
        df_display.insert(1, "Changed", changed)
    else:
        df_display.insert(1, "Changed", False)

    selected_row_index = None
    if current_pid is not None and "productId" in df_display.columns:
        matching_indexes = df_display.index[df_display["productId"].astype(str) == str(current_pid)]
        if len(matching_indexes) > 0:
            selected_row_index = int(matching_indexes[0])

    edited = st.data_editor(
        df_display,
        column_config={
            "Select": st.column_config.CheckboxColumn("Select", width="small"),
            "Changed": st.column_config.CheckboxColumn("Changed", width="small"),
        },
        hide_index=True,
        use_container_width=True,
        key="main_table",
        disabled=[c for c in df_display.columns if c != "Select"],
    )

    scroll_excel_a_to_selected_row(selected_row_index)

    selected_rows = edited[edited["Select"] == True]
    selected_pid_before = None
    if st.session_state.selected_row and isinstance(st.session_state.selected_row, dict):
        selected_pid_before = st.session_state.selected_row.get("productId")

    if len(selected_rows) == 0:
        if st.session_state.selected_row is not None:
            st.session_state.selected_row = None
            st.rerun()
    else:
        # If multiple checkboxes are True, prefer the one that differs from the previous selection.
        # This prevents "clicking a row above" from keeping the older (below) selection as the winner.
        if selected_pid_before is not None and "productId" in selected_rows.columns:
            other_rows = selected_rows[
                selected_rows["productId"].astype(str) != str(selected_pid_before)
            ]
            if len(other_rows) >= 1:
                new_row = other_rows.iloc[-1]
            else:
                new_row = selected_rows.iloc[-1]
        else:
            new_row = selected_rows.iloc[-1]

        new_pid = new_row.get("productId")
        if str(new_pid) != str(selected_pid_before):
            st.session_state.selected_row = new_row.drop("Select").to_dict()
            st.rerun()

with right:
    sel = st.session_state.selected_row

    if sel:
        brand = str(sel.get("brand", ""))
        subcat = str(sel.get("subcategory", ""))

        st.subheader("Selected Record")
        st.markdown(f"**{sel.get('name', '')}**")
        st.caption(f"ProductId: `{sel.get('productId','')}` · item_No: `{sel.get('item_No','')}` · Brand: `{brand}` · Subcategory: `{subcat}`")

        st.divider()
        st.subheader("Matching Items from Excel B")

        matches = filter_b(brand, subcat)

        if matches.empty:
            st.warning("No matches found in Excel B for this brand + subcategory.")
        else:
            item_no_col = next((c for c in matches.columns if "item_no" in c.lower().replace("_","")), matches.columns[0])
            item_name_col = next((c for c in matches.columns if "item_name" in c.lower().replace("_","")), matches.columns[1])
            status_col = next((c for c in matches.columns if "status" in c.lower()), None)

            st.caption(f"{len(matches)} matches found")

            # Show the full matching Excel B rows (all columns) for transparency.
            st.dataframe(matches, hide_index=True, use_container_width=True)

            # One shared textbox for all Apply buttons in this match list.
            filter_key = f"{brand}__{subcat}"
            if (
                "custom_item_no_filter" not in st.session_state
                or st.session_state.custom_item_no_filter != filter_key
            ):
                st.session_state.custom_item_no_filter = filter_key
                # Default: first matched item's item_no
                st.session_state.custom_item_no = str(matches.iloc[0][item_no_col]) if not matches.empty else ""

            custom_item_no = st.text_input(
                "New item_No",
                value=str(st.session_state.get("custom_item_no", "")),
                key="custom_item_no",
                label_visibility="collapsed",
            )

            for i, (_, row) in enumerate(matches.iterrows()):
                c1, c2 = st.columns([5, 2])
                status_badge = ""
                if status_col:
                    s = str(row.get(status_col, ""))
                    color = "green" if "active" in s.lower() else "red"
                    status_badge = f" :{color}[{s}]"
                label = f"**{row[item_name_col]}**  \n`Item No: {row[item_no_col]}`{status_badge}"
                with c1:
                    st.markdown(label)
                with c2:
                    if st.button("Apply", key=f"apply_{i}"):
                        if str(custom_item_no).strip() == "":
                            st.error("Please enter a valid item_No in the textbox above.")
                        else:
                            apply_replacement(
                                sel["productId"],
                                custom_item_no,
                                str(row[item_name_col]),
                                brand,
                                subcat,
                                a_row={k: v for k, v in sel.items() if k != "Changed"},
                                b_row=row.to_dict(),
                            )
                            st.rerun()
    else:
        st.info("Select a row from Excel A to see matching items.")

    st.divider()
    st.subheader("Change Log")
    log = st.session_state.change_log
    if not log:
        st.caption("No changes yet.")
    else:
        tab_recent, tab_full = st.tabs(["Recent (50)", "Full Details"])

        with tab_recent:
            for entry in log[:50]:
                brand_txt = entry.get("brand", "")
                subcat_txt = entry.get("subcategory", "")
                st.markdown(
                    f"<div class='change-row'>"
                    f"<span style='color:#888;font-size:0.78rem'>{entry.get('ts','')}</span><br>"
                    f"<b>{entry.get('name','')}</b><br>"
                    f"<span style='color:#888;font-size:0.78rem'>{brand_txt} · {subcat_txt}</span><br>"
                    f"<span class='tag-inactive'>was {entry.get('old_item','')}</span> → "
                    f"<span class='tag-replaced'>{entry.get('new_item','')} · {entry.get('item_name','')}</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )

        with tab_full:
            # Summary table (old-style) + selector to view full A/B rows.
            summary_rows = []
            for entry in log:
                summary_rows.append(
                    {
                        "ts": entry.get("ts", ""),
                        "productId": entry.get("productId", ""),
                        "name": entry.get("name", ""),
                        "brand": entry.get("brand", ""),
                        "subcategory": entry.get("subcategory", ""),
                        "old_item_no": entry.get("old_item", ""),
                        "updated_item_no": entry.get("new_item", ""),
                        "item_name": entry.get("item_name", ""),
                    }
                )

            df_summary = pd.DataFrame(summary_rows)
            st.dataframe(df_summary, hide_index=True, use_container_width=True)

            # Pick one log entry to view full Excel A and Excel B rows.
            option_labels = []
            for i, entry in enumerate(log):
                option_labels.append(
                    f"{i+1}) {entry.get('ts','')} | {entry.get('productId','')} | "
                    f"{entry.get('old_item','')} -> {entry.get('new_item','')} | {entry.get('item_name','')}"
                )

            selected_i = st.selectbox(
                "View full rows for selected log entry",
                options=list(range(len(log))),
                format_func=lambda i: option_labels[i],
                index=0,
            )

            entry = log[selected_i]
            a_row = entry.get("a_row") if isinstance(entry.get("a_row"), dict) else {}
            b_row = entry.get("b_row") if isinstance(entry.get("b_row"), dict) else {}

            # Backward-compatible fallback for older history entries.
            if not a_row:
                a_row = {
                    "productId": entry.get("productId", ""),
                    "name": entry.get("name", ""),
                    "brand": entry.get("brand", ""),
                    "subcategory": entry.get("subcategory", ""),
                    "item_No": entry.get("old_item", ""),
                }
            if not b_row:
                b_row = {
                    "item_name": entry.get("item_name", ""),
                    "item_no": entry.get("new_item", ""),
                }

            ca, cb = st.columns(2)
            with ca:
                st.subheader("Excel A row (full)")
                st.dataframe(pd.DataFrame([a_row]), hide_index=True, use_container_width=True)
            with cb:
                st.subheader("Excel B row (full)")
                st.dataframe(pd.DataFrame([b_row]), hide_index=True, use_container_width=True)
