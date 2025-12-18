import streamlit as st
import pandas as pd
from io import BytesIO

# ============================
# Page
# ============================
st.set_page_config(page_title="Question Builder", layout="wide")
st.title("Question Builder ➜ Excel (SPSS Style)")

# ============================
# Global session init
# ============================
if "questions" not in st.session_state:
    st.session_state.questions = []  # module 1 saved variables

if "choices" not in st.session_state:
    st.session_state.choices = [{"label": "", "code": ""}]  # module 1 working choices list

if "form_version" not in st.session_state:
    st.session_state.form_version = 0

if "selected_q_index" not in st.session_state:
    st.session_state.selected_q_index = None  # module 1 selection

if "mode" not in st.session_state:
    st.session_state.mode = "new"  # module 1: "new" or "edit"

if "_defaults" not in st.session_state:
    st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}

# Module 2 state
if "recodes" not in st.session_state:
    # list of recode variables, each:
    # {
    #   "source_qname": "...",
    #   "qname": "...",
    #   "prompt": "...",
    #   "label": "...",
    #   "groups": [
    #       {"new_label": "...", "new_code": "...", "source_pairs": [{"label": "...","code":"..."}...]}
    #   ]
    # }
    st.session_state.recodes = []

if "recode_selected_source" not in st.session_state:
    st.session_state.recode_selected_source = None

if "recode_work_groups" not in st.session_state:
    st.session_state.recode_work_groups = []  # current groups being built

if "recode_form_version" not in st.session_state:
    st.session_state.recode_form_version = 0


# ============================
# Shared helpers
# ============================
def to_excel_bytes_from_export_df(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)
    return output.getvalue()


# ============================
# Module selector (top)
# ============================
MODULES = {
    "Module 0: Project": "project",
    "Module 1: Scripting": "scripting",
    "Module 2: Recodes": "recodes",
    "Module 3: Import + Match Data": "import_match"
}

module_label = st.selectbox("Select module", list(MODULES.keys()))
active_module = MODULES[module_label]

st.divider()


# =========================================================
# MODULE 1: SCRIPTING
# =========================================================
def module1_add_choice():
    st.session_state.choices.append({"label": "", "code": ""})

def module1_remove_last_choice():
    if len(st.session_state.choices) > 1:
        st.session_state.choices.pop()

def module1_reset_form():
    st.session_state.choices = [{"label": "", "code": ""}]
    st.session_state.form_version += 1

def module1_load_question_into_form(q: dict):
    st.session_state.form_version += 1
    st.session_state.choices = [{"label": c["label"], "code": c["code"]} for c in q["choices"]]
    return q["qname"], q["prompt"], q.get("label", "")

def module1_validate_and_clean(qname: str, prompt: str, label: str):
    cleaned = [
        {"label": c["label"].strip(), "code": c["code"].strip()}
        for c in st.session_state.choices
        if c["label"].strip() and c["code"].strip()
    ]

    if not qname.strip():
        return None, "Question Name is required."
    if not prompt.strip():
        return None, "Question prompt is required."
    if not label.strip():
        return None, "Question label is required."
    if len(cleaned) == 0:
        return None, "At least one choice with a value code is required."
    return cleaned, None

def module1_build_export_df():
    rows = []
    for idx, q in enumerate(st.session_state.questions, start=1):
        header = f"{idx}.    {q['qname']}: {q['prompt']}"
        rows.append({"Text": header, "Value": ""})
        for ch in q["choices"]:
            rows.append({"Text": ch["label"], "Value": ch["code"]})
        rows.append({"Text": "", "Value": ""})
    return pd.DataFrame(rows, columns=["Text", "Value"])


def render_module_1():
    # Sidebar: variable list + actions
    st.sidebar.header("Variables coded")

    if len(st.session_state.questions) == 0:
        st.sidebar.info("No variables yet.")
    else:
        options = [f"{i+1}. {q['qname']} — {q.get('label','')}" for i, q in enumerate(st.session_state.questions)]
        sel_index = 0 if st.session_state.selected_q_index is None else st.session_state.selected_q_index
        sel = st.sidebar.radio("Select a variable", options=options, index=sel_index)
        st.session_state.selected_q_index = options.index(sel)

        q_selected = st.session_state.questions[st.session_state.selected_q_index]

        st.sidebar.divider()
        cA, cB, cC = st.sidebar.columns(3)

        if cA.button("Duplicate", use_container_width=True, key="m1_dup"):
            dup = {
                "qname": q_selected["qname"] + "_COPY",
                "prompt": q_selected["prompt"],
                "label": q_selected.get("label", ""),
                "choices": [dict(x) for x in q_selected["choices"]],
            }
            st.session_state.questions.append(dup)
            st.session_state.selected_q_index = len(st.session_state.questions) - 1

            st.session_state.mode = "edit"
            qname_default, prompt_default, label_default = module1_load_question_into_form(dup)
            st.session_state._defaults = {"qname": qname_default, "prompt": prompt_default, "label": label_default}
            st.rerun()

        if cB.button("Edit", use_container_width=True, key="m1_edit"):
            st.session_state.mode = "edit"
            qname_default, prompt_default, label_default = module1_load_question_into_form(q_selected)
            st.session_state._defaults = {"qname": qname_default, "prompt": prompt_default, "label": label_default}
            st.rerun()

        if cC.button("Delete", use_container_width=True, key="m1_del"):
            st.session_state.questions.pop(st.session_state.selected_q_index)
            st.session_state.selected_q_index = None if len(st.session_state.questions) == 0 else 0
            st.session_state.mode = "new"
            module1_reset_form()
            st.sidebar.success("Deleted.")
            st.rerun()

    st.sidebar.divider()
    if st.sidebar.button("➕ New variable", use_container_width=True, key="m1_new"):
        st.session_state.mode = "new"
        st.session_state.selected_q_index = None
        st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}
        module1_reset_form()
        st.rerun()

    defaults = st.session_state._defaults

    # Main form
    form_key = f"script_form_{st.session_state.form_version}"
    with st.form(form_key):
        st.subheader("Build / Edit a variable")

        qname = st.text_input("Question Name", value=defaults.get("qname", ""), placeholder="QAGE")
        prompt = st.text_area("Question prompt", value=defaults.get("prompt", ""), placeholder="What is your age?")
        label = st.text_input(
            "Question label",
            value=defaults.get("label", ""),
            placeholder="Age",
            help="Variable name as it will appear on toplines: ex. QAGE = Age",
        )

        st.subheader("Question choices")
        for i, choice in enumerate(st.session_state.choices):
            col1, col2 = st.columns([3, 1])
            with col1:
                st.session_state.choices[i]["label"] = st.text_input(
                    f"Choice {i+1} label",
                    value=choice["label"],
                    key=f"{form_key}_label_{i}",
                    placeholder="18–24",
                )
            with col2:
                st.session_state.choices[i]["code"] = st.text_input(
                    "Code",
                    value=choice["code"],
                    key=f"{form_key}_code_{i}",
                    placeholder="2",
                )

        a1, a2, a3, a4 = st.columns([1, 1, 1, 2])
        with a1:
            add_clicked = st.form_submit_button("➕ Add choice")
        with a2:
            remove_clicked = st.form_submit_button("➖ Remove last choice")
        with a3:
            clear_clicked = st.form_submit_button("🧹 Clear")
        with a4:
            save_clicked = st.form_submit_button("💾 Save (new / update)")

    if add_clicked:
        module1_add_choice()
        st.rerun()

    if remove_clicked:
        module1_remove_last_choice()
        st.rerun()

    if clear_clicked:
        st.session_state.mode = "new"
        st.session_state.selected_q_index = None
        st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}
        module1_reset_form()
        st.rerun()

    if save_clicked:
        cleaned_choices, err = module1_validate_and_clean(qname, prompt, label)
        if err:
            st.error(err)
        else:
            payload = {
                "qname": qname.strip(),
                "prompt": prompt.strip(),
                "label": label.strip(),
                "choices": cleaned_choices,
            }

            if (
                st.session_state.mode == "edit"
                and st.session_state.selected_q_index is not None
                and st.session_state.selected_q_index < len(st.session_state.questions)
            ):
                st.session_state.questions[st.session_state.selected_q_index] = payload
                st.success("Updated.")
            else:
                st.session_state.questions.append(payload)
                st.session_state.selected_q_index = len(st.session_state.questions) - 1
                st.success("Saved.")

            st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}
            st.session_state.mode = "new"
            module1_reset_form()
            st.rerun()

    # Export
    st.divider()
    st.subheader("Saved questions (export preview)")

    if len(st.session_state.questions) == 0:
        st.info("No questions saved yet.")
    else:
        export_df = module1_build_export_df()
        st.dataframe(export_df, use_container_width=True, hide_index=True)

        st.download_button(
            label="⬇️ Download Excel (SPSS-style)",
            data=to_excel_bytes_from_export_df(export_df, sheet_name="Questions"),
            file_name="questions_spss_style.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if st.button("🗑️ Delete all saved questions", key="m1_delete_all"):
            st.session_state.questions = []
            st.session_state.selected_q_index = None
            st.session_state.mode = "new"
            st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}
            module1_reset_form()
            st.success("Cleared.")
            st.rerun()


# ============================
# Module 2: Recodes  (EDITED)
# ============================

def _safe_int(x):
    try:
        return int(str(x).strip())
    except Exception:
        return None


# ---- Module 2 session init (safe if already defined) ----
if "recodes" not in st.session_state:
    st.session_state.recodes = []  # list of saved recodes

if "m2_selected_recode_index" not in st.session_state:
    st.session_state.m2_selected_recode_index = None  # selected saved recode index

if "m2_mode" not in st.session_state:
    # "new" | "edit"
    st.session_state.m2_mode = "new"

if "m2_ui_version" not in st.session_state:
    st.session_state.m2_ui_version = 0  # bump to force widget refresh / new keys

if "m2_pick" not in st.session_state:
    st.session_state.m2_pick = set()  # selected source codes for the group being built

if "m2_defaults" not in st.session_state:
    st.session_state.m2_defaults = {
        "source_qname": "",
        "new_qname": "",
        "new_label": "",
        "group_text": "",
        "group_code": "",
    }

if "m2_work_groups" not in st.session_state:
    # edit buffer for groups currently being edited
    st.session_state.m2_work_groups = []

if "m2_last_sidebar_sel" not in st.session_state:
    st.session_state.m2_last_sidebar_sel = None

# NEW: track the last source in the UI, so we can detect user source changes
if "m2_last_source_qname" not in st.session_state:
    st.session_state.m2_last_source_qname = None


def _module2_find_source_question(qname: str):
    return next((q for q in st.session_state.questions if q.get("qname") == qname), None)


def build_module2_export_df():
    """
    Output format:

    cQAGE
    18-34    2
    35-44    3
    ...
    (blank row)

    Includes non-recoded original choices automatically (by code order).
    """
    rows = []
    recodes = st.session_state.get("recodes", [])

    for rec in recodes:
        source = str(rec.get("source_qname", "")).strip()
        new_q = str(rec.get("new_qname", "")).strip()
        if not source or not new_q:
            continue

        src_q = _module2_find_source_question(source)
        src_choices = src_q.get("choices", []) if src_q else []

        used_source_codes = set()
        groups = rec.get("groups", [])
        for g in groups:
            for f in g.get("from", []):
                c = str(f.get("code", "")).strip()
                if c:
                    used_source_codes.add(c)

        out_items = []
        for g in groups:
            txt = str(g.get("new_text", "")).strip()
            code = str(g.get("new_code", "")).strip()
            if txt and code:
                out_items.append({"Text": txt, "Value": code})

        # Add any source choices not recoded
        for ch in src_choices:
            ch_label = str(ch.get("label", "")).strip()
            ch_code = str(ch.get("code", "")).strip()
            if not ch_label or not ch_code:
                continue
            if ch_code in used_source_codes:
                continue
            out_items.append({"Text": ch_label, "Value": ch_code})

        out_items.sort(
            key=lambda x: (
                _safe_int(x["Value"]) if _safe_int(x["Value"]) is not None else 10**9,
                x["Text"],
            )
        )

        rows.append({"Text": new_q, "Value": ""})
        rows.extend(out_items)
        rows.append({"Text": "", "Value": ""})

    return pd.DataFrame(rows, columns=["Text", "Value"])


def _module2_load_into_ui(rec: dict, as_new: bool):
    """
    Load a saved recode into the UI.

    - as_new=False: editing an existing saved recode (Update will overwrite)
    - as_new=True:  duplicate workflow (Save as new creates a new one)
    """
    st.session_state.m2_pick = set()

    # defaults
    st.session_state.m2_defaults = {
        "source_qname": rec.get("source_qname", ""),
        "new_qname": rec.get("new_qname", ""),
        "new_label": rec.get("new_label", ""),
        "group_text": "",
        "group_code": "",
    }

    # edit buffer groups (copy)
    st.session_state.m2_work_groups = [
        {
            "new_text": g.get("new_text", ""),
            "new_code": g.get("new_code", ""),
            "from": [dict(x) for x in g.get("from", [])],
        }
        for g in rec.get("groups", [])
    ]

    # mode + selection
    if as_new:
        st.session_state.m2_mode = "new"
        st.session_state.m2_selected_recode_index = None
    else:
        st.session_state.m2_mode = "edit"
        st.session_state.m2_selected_recode_index = st.session_state.recodes.index(rec)

    # IMPORTANT: keep last_source in sync so we can avoid auto-overwrites
    st.session_state.m2_last_source_qname = rec.get("source_qname", "")

    # force all module-2 widgets to rebuild with fresh keys
    st.session_state.m2_ui_version += 1


def render_module_2():
    st.header("Module 2: Recodes")

    # Must have module 1 variables
    if len(st.session_state.questions) == 0:
        st.info("No variables from Module 1 yet. Add variables in Scripting first.")
        return

    # ----------------------------
    # Sidebar: recode list + actions
    # ----------------------------
    st.sidebar.header("Recodes")

    if len(st.session_state.recodes) == 0:
        st.sidebar.info("No recodes yet.")
    else:
        rec_opts = [
            f"{i+1}. {r.get('new_qname','')}  (from {r.get('source_qname','')})"
            for i, r in enumerate(st.session_state.recodes)
        ]
        sel_idx = (
            0
            if st.session_state.m2_selected_recode_index is None
            else st.session_state.m2_selected_recode_index
        )

        # Radio selection
        sel = st.sidebar.radio(
            "Select a recode",
            rec_opts,
            index=sel_idx,
            key="m2_recode_radio",
        )
        new_index = rec_opts.index(sel)

        # Auto-load when sidebar selection changes
        if st.session_state.m2_last_sidebar_sel != new_index:
            st.session_state.m2_last_sidebar_sel = new_index
            rec_selected = st.session_state.recodes[new_index]
            _module2_load_into_ui(rec_selected, as_new=False)
            st.rerun()

        rec_selected = st.session_state.recodes[st.session_state.m2_last_sidebar_sel]

        st.sidebar.divider()
        cA, cB, cC = st.sidebar.columns(3)

        # Duplicate: load into UI as NEW (do not auto-create a saved record)
        if cA.button("Duplicate", use_container_width=True, key="m2_dup"):
            dup = {
                "source_qname": rec_selected.get("source_qname", ""),
                # keep same defaults so user can *rename/label* quickly
                # (you can optionally set _COPY here if you want)
                "new_qname": rec_selected.get("new_qname", ""),
                "new_label": rec_selected.get("new_label", ""),
                "groups": [dict(g) for g in rec_selected.get("groups", [])],
            }
            _module2_load_into_ui(dup, as_new=True)
            st.rerun()

        # Edit: just ensure loaded (already is), but keep for parity
        if cB.button("Edit", use_container_width=True, key="m2_edit"):
            _module2_load_into_ui(rec_selected, as_new=False)
            st.rerun()

        # Delete
        if cC.button("Delete", use_container_width=True, key="m2_del"):
            del_idx = st.session_state.m2_last_sidebar_sel
            st.session_state.recodes.pop(del_idx)
            st.session_state.m2_last_sidebar_sel = None
            st.session_state.m2_selected_recode_index = None
            st.session_state.m2_mode = "new"
            st.session_state.m2_defaults = {
                "source_qname": "",
                "new_qname": "",
                "new_label": "",
                "group_text": "",
                "group_code": "",
            }
            st.session_state.m2_work_groups = []
            st.session_state.m2_pick = set()
            st.session_state.m2_last_source_qname = None
            st.session_state.m2_ui_version += 1
            st.sidebar.success("Deleted.")
            st.rerun()

    st.sidebar.divider()
    if st.sidebar.button("➕ New recode", use_container_width=True, key="m2_new"):
        st.session_state.m2_mode = "new"
        st.session_state.m2_selected_recode_index = None
        st.session_state.m2_last_sidebar_sel = None
        st.session_state.m2_pick = set()
        st.session_state.m2_work_groups = []
        st.session_state.m2_defaults = {
            "source_qname": "",
            "new_qname": "",
            "new_label": "",
            "group_text": "",
            "group_code": "",
        }
        st.session_state.m2_last_source_qname = None
        st.session_state.m2_ui_version += 1
        st.rerun()

    # ----------------------------
    # Main: choose source variable
    # ----------------------------
    qnames = [q["qname"] for q in st.session_state.questions]

    default_source = st.session_state.m2_defaults.get("source_qname") or qnames[0]
    if default_source not in qnames:
        default_source = qnames[0]

    # versioned key so we can refresh widgets safely
    source_qname = st.selectbox(
        "Select source variable to recode",
        qnames,
        index=qnames.index(default_source),
        key=f"m2_source_select_{st.session_state.m2_ui_version}",
    )

    src_q = _module2_find_source_question(source_qname)
    src_choices = src_q.get("choices", []) if src_q else []

    # ----------------------------
    # FIX: "autofill" should NOT stomp on a duplicated recode rename/label.
    #
    # Behavior:
    # - always update defaults["source_qname"] when the source changes
    # - only auto-fill new_qname/new_label if user has NOT customized them
    #   (i.e., they are blank OR still equal to the old auto-filled values)
    # ----------------------------
    prev_source = st.session_state.m2_last_source_qname
    if prev_source is None:
        prev_source = default_source

    if source_qname != prev_source:
        # update source in defaults
        st.session_state.m2_defaults["source_qname"] = source_qname

        # determine "old" auto values
        old_auto_qname = f"c{prev_source}" if prev_source else ""
        old_auto_label = ""
        prev_q = _module2_find_source_question(prev_source) if prev_source else None
        if prev_q:
            old_auto_label = str(prev_q.get("label", "")).strip()

        # current values in defaults (what UI is holding)
        cur_new_qname = str(st.session_state.m2_defaults.get("new_qname", "")).strip()
        cur_new_label = str(st.session_state.m2_defaults.get("new_label", "")).strip()

        # only overwrite if user hasn't customized
        allow_qname_autofill = (cur_new_qname == "" or cur_new_qname == old_auto_qname)
        allow_label_autofill = (cur_new_label == "" or cur_new_label == old_auto_label)

        if allow_qname_autofill:
            st.session_state.m2_defaults["new_qname"] = f"c{source_qname}"

        if allow_label_autofill:
            st.session_state.m2_defaults["new_label"] = str(src_q.get("label", "")).strip() if src_q else ""

        # changing source invalidates current selection + groups (safer default)
        st.session_state.m2_pick = set()
        st.session_state.m2_work_groups = []

        # update last source + refresh widgets
        st.session_state.m2_last_source_qname = source_qname
        st.session_state.m2_ui_version += 1
        st.rerun()
    else:
        # keep last source in sync
        st.session_state.m2_last_source_qname = source_qname

    st.subheader("Create / edit a recode")

    new_qname = st.text_input(
        "New variable name (recoded)",
        value=st.session_state.m2_defaults.get("new_qname", f"c{source_qname}"),
        key=f"m2_new_qname_{st.session_state.m2_ui_version}",
    )
    new_label = st.text_input(
        "New variable label",
        value=st.session_state.m2_defaults.get(
            "new_label",
            str(src_q.get("label", "")).strip() if src_q else "",
        ),
        key=f"m2_new_label_{st.session_state.m2_ui_version}",
    )

    # keep defaults synced to what user types (so we can make good overwrite decisions later)
    st.session_state.m2_defaults["new_qname"] = new_qname
    st.session_state.m2_defaults["new_label"] = new_label

    st.caption("Select source choices to combine into a new recoded choice. Repeat to create multiple groups.")

    # ----------------------------
    # Source choices checkboxes
    # ----------------------------
    st.markdown("**Source choices**")
    cols = st.columns(3)

    for i, ch in enumerate(src_choices):
        lab = str(ch.get("label", "")).strip()
        code = str(ch.get("code", "")).strip()
        if not lab or not code:
            continue

        cb_key = f"m2_cb_{source_qname}_{st.session_state.m2_ui_version}_{i}_{code}"
        with cols[i % 3]:
            checked = st.checkbox(f"{lab} ({code})", key=cb_key)

        if checked:
            st.session_state.m2_pick.add(code)
        else:
            st.session_state.m2_pick.discard(code)

    st.markdown("---")
    g1, g2, g3 = st.columns([3, 1, 1])

    with g1:
        group_text = st.text_input(
            "New recoded choice text",
            value=st.session_state.m2_defaults.get("group_text", ""),
            placeholder="18-34",
            key=f"m2_group_text_{st.session_state.m2_ui_version}",
        )
    with g2:
        group_code = st.text_input(
            "New recoded choice code",
            value=st.session_state.m2_defaults.get("group_code", ""),
            placeholder="2",
            key=f"m2_group_code_{st.session_state.m2_ui_version}",
        )
    with g3:
        if st.button("🧽 Clear selection", key=f"m2_clear_pick_{st.session_state.m2_ui_version}"):
            st.session_state.m2_pick = set()
            st.session_state.m2_ui_version += 1
            st.rerun()

    # ----------------------------
    # Add group (to buffer only)
    # ----------------------------
    if st.button("➕ Add recode group", key=f"m2_add_group_{st.session_state.m2_ui_version}"):
        if not st.session_state.m2_pick:
            st.error("Select at least one source choice to recode.")
        elif not group_text.strip() or not group_code.strip():
            st.error("Enter both a recoded choice text and code.")
        else:
            from_list = []
            for ch in src_choices:
                c = str(ch.get("code", "")).strip()
                if c in st.session_state.m2_pick:
                    from_list.append({"label": str(ch.get("label", "")).strip(), "code": c})

            st.session_state.m2_work_groups.append({
                "new_text": group_text.strip(),
                "new_code": group_code.strip(),
                "from": from_list,
            })

            st.session_state.m2_pick = set()
            st.session_state.m2_defaults["group_text"] = ""
            st.session_state.m2_defaults["group_code"] = ""
            st.session_state.m2_ui_version += 1
            st.success("Added recode group.")
            st.rerun()

    # ----------------------------
    # Show current buffer groups + delete last
    # ----------------------------
    if st.session_state.m2_work_groups:
        st.subheader("Current recode groups")
        for gi, g in enumerate(st.session_state.m2_work_groups, start=1):
            frm = ", ".join([f"{x['label']}({x['code']})" for x in g.get("from", [])])
            st.write(f"{gi}. **{g.get('new_text','')}** → {g.get('new_code','')}  _(from: {frm})_")

        if st.button("🗑️ Delete last group", key=f"m2_del_last_group_{st.session_state.m2_ui_version}"):
            st.session_state.m2_work_groups.pop()
            st.session_state.m2_ui_version += 1
            st.rerun()

    # ----------------------------
    # Save / Update recode
    # ----------------------------
    s1, s2 = st.columns([1, 1])

    with s1:
        if st.button("💾 Save as new recode", key=f"m2_save_new_{st.session_state.m2_ui_version}"):
            if not new_qname.strip():
                st.error("New variable name is required.")
            else:
                payload = {
                    "source_qname": source_qname,
                    "new_qname": new_qname.strip(),
                    "new_label": new_label.strip(),
                    "groups": [dict(g) for g in st.session_state.m2_work_groups],
                }
                st.session_state.recodes.append(payload)

                # After save: select the new saved recode and load it (as edit)
                st.session_state.m2_last_sidebar_sel = len(st.session_state.recodes) - 1
                _module2_load_into_ui(payload, as_new=False)
                st.success("Saved as new recode.")
                st.rerun()

    with s2:
        if st.button("✅ Update selected recode", key=f"m2_update_{st.session_state.m2_ui_version}"):
            if st.session_state.m2_last_sidebar_sel is None:
                st.error("No recode selected to update.")
            else:
                idx = st.session_state.m2_last_sidebar_sel
                if idx < 0 or idx >= len(st.session_state.recodes):
                    st.error("Selected recode index is invalid.")
                elif not new_qname.strip():
                    st.error("New variable name is required.")
                else:
                    st.session_state.recodes[idx] = {
                        "source_qname": source_qname,
                        "new_qname": new_qname.strip(),
                        "new_label": new_label.strip(),
                        "groups": [dict(g) for g in st.session_state.m2_work_groups],
                    }
                    st.success("Updated recode.")
                    st.session_state.m2_ui_version += 1
                    st.rerun()

    # ----------------------------
    # Export preview
    # ----------------------------
    st.subheader("Export preview (recodes)")
    export_df = build_module2_export_df()

    if export_df.empty:
        st.info("No recodes created yet.")
    else:
        st.dataframe(export_df, use_container_width=True, hide_index=True)

        st.download_button(
            label="⬇️ Download Excel (recodes)",
            data=to_excel_bytes_from_export_df(export_df, sheet_name="Recodes"),
            file_name="recodes_spss_style.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )












# =========================================================
# MODULE 0: PROJECT (Save / Load)
# =========================================================
import json
from datetime import datetime

# --- Project state init ---
if "project" not in st.session_state:
    st.session_state.project = {
        "name": "Untitled project",
        "last_saved": None,
    }

if "project_slots" not in st.session_state:
    # quick in-session saves (not persistent across refresh)
    st.session_state.project_slots = {}

def _project_payload() -> dict:
    """
    Everything needed to restore the app state later.
    Keep this conservative: only stable data, not widget keys.
    """
    return {
        "meta": {
            "app": "Question Builder ➜ Excel (SPSS Style)",
            "version": 1,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "project_name": st.session_state.project.get("name", "Untitled project"),
        },
        "module1": {
            "questions": st.session_state.get("questions", []),
        },
        "module2": {
            "recodes": st.session_state.get("recodes", []),
        },
    }

def _load_project_payload(payload: dict):
    """
    Restore project state into session_state.
    Also resets working forms so UI is clean after load.
    """
    # Basic validation
    if not isinstance(payload, dict):
        raise ValueError("Project file is not valid JSON.")
    if "module1" not in payload or "module2" not in payload:
        raise ValueError("Project file missing expected keys (module1/module2).")

    # Restore
    st.session_state.questions = payload.get("module1", {}).get("questions", [])
    st.session_state.recodes = payload.get("module2", {}).get("recodes", [])

    # Project meta
    meta = payload.get("meta", {})
    st.session_state.project["name"] = meta.get("project_name", "Untitled project")
    st.session_state.project["last_saved"] = meta.get("saved_at", None)

    # Reset module 1 working state
    st.session_state.choices = [{"label": "", "code": ""}]
    st.session_state.form_version = st.session_state.get("form_version", 0) + 1
    st.session_state.selected_q_index = None
    st.session_state.mode = "new"
    st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}

    # Reset module 2 working state (safe even if you later rename these)
    st.session_state.recode_pick = set()
    st.session_state.recode_form_version = st.session_state.get("recode_form_version", 0) + 1
    st.session_state.recode_selected_source = None
    st.session_state.recode_work_groups = []

def _bytes_for_download(payload: dict) -> bytes:
    return json.dumps(payload, indent=2).encode("utf-8")

def render_module_0():
    st.header("Module 0: Project")

    # Project name
    st.session_state.project["name"] = st.text_input(
        "Project name",
        value=st.session_state.project.get("name", "Untitled project"),
        key="m0_project_name",
    )

    st.caption(
        "Save = download a project file (.json). Open = upload that file later to continue editing."
    )

    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.6, 1.4])

    # --- New project ---
    with c1:
        if st.button("🆕 New project", use_container_width=True, key="m0_new_project"):
            # Clear core data
            st.session_state.questions = []
            st.session_state.recodes = []

            # Reset forms
            st.session_state.choices = [{"label": "", "code": ""}]
            st.session_state.form_version += 1
            st.session_state.selected_q_index = None
            st.session_state.mode = "new"
            st.session_state._defaults = {"qname": "", "prompt": "", "label": ""}

            st.session_state.recode_pick = set()
            st.session_state.recode_form_version += 1
            st.session_state.recode_selected_source = None
            st.session_state.recode_work_groups = []

            st.session_state.project["last_saved"] = None
            st.success("Started a new project.")
            st.rerun()

    # --- Save project (download JSON) ---
    with c2:
        payload = _project_payload()
        default_name = st.session_state.project.get("name", "project").strip().replace(" ", "_") or "project"
        filename = f"{default_name}.json"

        st.download_button(
            "💾 Save project",
            data=_bytes_for_download(payload),
            file_name=filename,
            mime="application/json",
            use_container_width=True,
            key="m0_save_download",
        )

    # --- Open project (upload JSON) ---
    with c3:
        uploaded = st.file_uploader(
            "Open project (.json)",
            type=["json"],
            label_visibility="collapsed",
            key="m0_open_uploader",
        )
        if uploaded is not None:
            try:
                payload_in = json.loads(uploaded.getvalue().decode("utf-8"))
                _load_project_payload(payload_in)
                st.success("Project loaded.")
                st.rerun()
            except Exception as e:
                st.error(f"Could not load project: {e}")

    # --- Quick in-session slots ---
    with c4:
        st.markdown("**Quick slots (this session only)**")
        slot_name = st.text_input("Slot name", value="Slot 1", key="m0_slot_name")
        s1, s2 = st.columns(2)

        if s1.button("Save slot", use_container_width=True, key="m0_save_slot"):
            st.session_state.project_slots[slot_name] = _project_payload()
            st.success(f"Saved to slot: {slot_name}")

        if s2.button("Load slot", use_container_width=True, key="m0_load_slot"):
            if slot_name not in st.session_state.project_slots:
                st.warning("No saved data in that slot yet.")
            else:
                try:
                    _load_project_payload(st.session_state.project_slots[slot_name])
                    st.success(f"Loaded slot: {slot_name}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not load slot: {e}")

    st.divider()
    st.subheader("Project summary")
    st.write(
        {
            "Project name": st.session_state.project.get("name"),
            "Last saved (from file)": st.session_state.project.get("last_saved"),
            "Module 1 variables": len(st.session_state.get("questions", [])),
            "Module 2 recodes": len(st.session_state.get("recodes", [])),
        }
    )





# ============================
# Module 3: Import + Match Data
# ============================

import difflib
import numpy as np

def _normalize_name(x: str) -> str:
    return "".join(ch for ch in str(x).strip().lower() if ch.isalnum() or ch == "_")

def _similarity(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, _normalize_name(a), _normalize_name(b)).ratio()

def _best_match(col: str, targets: list[str], threshold: float = 0.80):
    best = None
    best_score = 0.0
    for t in targets:
        s = _similarity(col, t)
        if s > best_score:
            best_score = s
            best = t
    if best is not None and best_score >= threshold:
        return best, best_score
    return None, best_score

def _get_series_safe(df: pd.DataFrame, colname: str) -> pd.Series:
    """
    Always return a Series even if df has duplicate colnames.
    If duplicates exist, take the FIRST one.
    """
    if colname not in df.columns:
        return pd.Series([np.nan] * len(df), index=df.index)

    # df.loc[:, colname] returns Series if unique, DataFrame if duplicates
    obj = df.loc[:, colname]
    if isinstance(obj, pd.DataFrame):
        return obj.iloc[:, 0]
    return obj

def _build_choice_code_to_text(qdef: dict) -> dict:
    """
    From Module 1 question definition, build mapping: code(str)->label(str)
    """
    m = {}
    for ch in qdef.get("choices", []):
        code = str(ch.get("code", "")).strip()
        lab = str(ch.get("label", "")).strip()
        if code:
            m[code] = lab
    return m

def _apply_label_rollups(df_out: pd.DataFrame, questions: list[dict]):
    """
    For each Module 1 variable QXXX, create a label column like:
      QAGE__text  (or whatever you want)
    So 3 -> "35-44", etc.
    """
    for q in questions:
        qname = q.get("qname", "").strip()
        if not qname:
            continue
        if qname not in df_out.columns:
            continue

        code_to_text = _build_choice_code_to_text(q)
        s = _get_series_safe(df_out, qname).astype(str).str.strip()

        df_out[f"{qname}__text"] = s.map(lambda v: code_to_text.get(v, ""))

def _apply_recode_definitions(df_out: pd.DataFrame, questions: list[dict], recodes: list[dict]):
    """
    Creates recode columns (new_qname) from source_qname.

    Logic:
    - if value code belongs to any group.from codes -> output group.new_code
    - else keep original code (pass-through)
    """
    # quick lookup for source question defs
    q_lookup = {q.get("qname"): q for q in questions}

    for rec in recodes:
        source_q = str(rec.get("source_qname", "")).strip()
        new_q = str(rec.get("new_qname", "")).strip()
        if not source_q or not new_q:
            continue
        if source_q not in df_out.columns:
            continue

        groups = rec.get("groups", [])

        # Build map source_code(str) -> new_code(str)
        src_to_new = {}
        for g in groups:
            new_code = str(g.get("new_code", "")).strip()
            for f in g.get("from", []):
                src_code = str(f.get("code", "")).strip()
                if src_code and new_code:
                    src_to_new[src_code] = new_code

        s = _get_series_safe(df_out, source_q).astype(str).str.strip()

        # pass-through if not recoded
        df_out[new_q] = s.map(lambda v: src_to_new.get(v, v))

        # Also create text for recoded variable if possible:
        # Prefer group new_text for recoded codes; otherwise fall back to original label text
        newcode_to_text = {}
        for g in groups:
            nc = str(g.get("new_code", "")).strip()
            nt = str(g.get("new_text", "")).strip()
            if nc and nt:
                newcode_to_text[nc] = nt

        # fallback from source label mapping (if pass-through codes)
        src_qdef = q_lookup.get(source_q, {})
        src_code_to_text = _build_choice_code_to_text(src_qdef)

        df_out[f"{new_q}__text"] = df_out[new_q].astype(str).str.strip().map(
            lambda v: newcode_to_text.get(v, src_code_to_text.get(v, ""))
        )

def render_module_3():
    st.header("Module 3: Import + Match Data")

    if "module3_mapping" not in st.session_state:
        st.session_state.module3_mapping = {}  # canonical_qname -> input_colname OR None

    uploaded = st.file_uploader("Upload a dataset (.csv or Excel)", type=["csv", "xlsx", "xls"])
    if not uploaded:
        st.info("Upload a file to begin.")
        return

    # Read file
    if uploaded.name.lower().endswith(".csv"):
        df_in = pd.read_csv(uploaded)
    else:
        df_in = pd.read_excel(uploaded)

    st.write(f"Loaded **{df_in.shape[0]:,}** rows × **{df_in.shape[1]:,}** columns.")

    # Canonical vars from Module 1
    if len(st.session_state.questions) == 0:
        st.warning("No variables from Module 1. Build variables in Module 1 first.")
        return

    canonical_vars = [q["qname"] for q in st.session_state.questions if q.get("qname")]
    input_cols = list(df_in.columns)

    st.subheader("Auto-match columns (fuzzy 80%)")

    threshold = st.slider("Match threshold", 0.50, 0.95, 0.80, 0.01)

    # Auto-match: for each input col, find best canonical target
    # We will store mapping as canonical -> input col (one input per canonical)
    # If multiple inputs match same canonical, take the best score.
    auto_map = {}
    auto_scores = {}

    for col in input_cols:
        best, score = _best_match(col, canonical_vars, threshold=threshold)
        if best:
            if best not in auto_map or score > auto_scores.get(best, 0):
                auto_map[best] = col
                auto_scores[best] = score

    # Initialize session mapping if empty (first run)
    if not st.session_state.module3_mapping:
        for qn in canonical_vars:
            st.session_state.module3_mapping[qn] = auto_map.get(qn, None)

    # ---- Manual override UI
    st.subheader("Review / edit matches")

    # Add an explicit UNMATCH option
    picker_options = ["(unmatched)"] + input_cols

    # Show as a table-like UI
    for qn in canonical_vars:
        current = st.session_state.module3_mapping.get(qn, None)
        current_label = current if current in input_cols else "(unmatched)"

        cols = st.columns([2, 3, 1])
        with cols[0]:
            st.markdown(f"**{qn}**")
        with cols[1]:
            selected = st.selectbox(
                "Match",
                options=picker_options,
                index=picker_options.index(current_label),
                key=f"m3_match_{qn}",
                label_visibility="collapsed",
            )
            st.session_state.module3_mapping[qn] = None if selected == "(unmatched)" else selected
        with cols[2]:
            # quick unmatch button
            if st.button("Unmatch", key=f"m3_unmatch_{qn}"):
                st.session_state.module3_mapping[qn] = None
                st.rerun()

    # ---- Warning: columns starting with Q that didn't match to anything
    matched_input_cols = {v for v in st.session_state.module3_mapping.values() if v}
    q_like_unmatched = [c for c in input_cols if str(c).strip().upper().startswith("Q") and c not in matched_input_cols]

    if q_like_unmatched:
        st.warning(
            "These input columns start with **Q** but are currently **not matched** to any scripted variable:\n\n"
            + ", ".join(q_like_unmatched)
        )

    # ---- Build output df: keep all original columns + add canonical copies
    st.subheader("Build output (original columns + canonical + recodes)")

    df_out = df_in.copy()

    # Add canonical columns (do NOT delete originals)
    for qn, in_col in st.session_state.module3_mapping.items():
        if in_col is None:
            continue
        if in_col not in df_out.columns:
            continue

        # Create/overwrite canonical column with values from input column
        df_out[qn] = df_out[in_col]

    # Apply label rollups for canonical columns
    _apply_label_rollups(df_out, st.session_state.questions)

    # Apply recodes (creates cQ vars + cQ__text)
    _apply_recode_definitions(df_out, st.session_state.questions, st.session_state.get("recodes", []))

    # Preview a small slice
    st.subheader("Preview (first 25 rows)")
    st.dataframe(df_out.head(25), use_container_width=True)

    # Download
    st.subheader("Download output")
    fmt = st.radio("Output format", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)

    if fmt.startswith("Excel"):
        # Write Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name="Data")
        st.download_button(
            label="⬇️ Download mapped + recoded dataset (Excel)",
            data=output.getvalue(),
            file_name="mapped_recoded_dataset.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        csv_bytes = df_out.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Download mapped + recoded dataset (CSV)",
            data=csv_bytes,
            file_name="mapped_recoded_dataset.csv",
            mime="text/csv",
        )







# ============================
# Render selected module ----------------------------------------------------never delete. keep here
# ============================
if active_module == "project":
    render_module_0()
elif active_module == "scripting":
    render_module_1()
elif active_module == "recodes":
    render_module_2() 
elif active_module == "import_match":
    render_module_3() 
