import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Question Builder", layout="centered")
st.title("Question Builder ➜ Excel")

# ----------------------------
# Session state init
# ----------------------------
if "choices" not in st.session_state:
    st.session_state.choices = [""]  # start with 1 choice box

if "rows" not in st.session_state:
    st.session_state.rows = []  # saved questions

if "form_version" not in st.session_state:
    st.session_state.form_version = 0  # used to reset widgets safely

# ----------------------------
# Helpers
# ----------------------------
def add_choice_box():
    st.session_state.choices.append("")

def reset_form_safely():
    # Reset non-widget state, and bump the form key so Streamlit rebuilds widgets fresh
    st.session_state.choices = [""]
    st.session_state.form_version += 1

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Questions")
    return output.getvalue()

# ----------------------------
# Form (key changes when we reset)
# ----------------------------
form_key = f"question_form_{st.session_state.form_version}"

with st.form(form_key):
    qname = st.text_input("Question Name", placeholder="e.g., QAGE")
    qprompt = st.text_area("Question prompt", placeholder="e.g., What is your age?")

    st.subheader("Question choices")
    for i in range(len(st.session_state.choices)):
        st.session_state.choices[i] = st.text_input(
            f"Choice {i+1}",
            value=st.session_state.choices[i],
            key=f"{form_key}_choice_{i}",  # important: unique per form version
            placeholder="e.g., 18-24"
        )

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        add_choice_clicked = st.form_submit_button("➕ Add choice")
    with c2:
        save_clicked = st.form_submit_button("💾 Save question")
    with c3:
        clear_clicked = st.form_submit_button("🧹 Clear form")

# Handle button actions
if add_choice_clicked:
    add_choice_box()
    st.rerun()

if clear_clicked:
    reset_form_safely()
    st.rerun()

if save_clicked:
    cleaned_choices = [c.strip() for c in st.session_state.choices if c.strip()]

    if not qname.strip():
        st.error("Please fill in **Question Name**.")
    elif not qprompt.strip():
        st.error("Please fill in **Question prompt**.")
    elif len(cleaned_choices) == 0:
        st.error("Please add at least **one** question choice.")
    else:
        st.session_state.rows.append(
            {"Question Name": qname.strip(), "Question prompt": qprompt.strip(), "choices": cleaned_choices}
        )
        st.success("Saved!")
        reset_form_safely()
        st.rerun()

# ----------------------------
# Preview + Download
# ----------------------------
st.divider()
st.subheader("Saved questions")

if len(st.session_state.rows) == 0:
    st.info("No questions saved yet.")
else:
    max_choices = max(len(r["choices"]) for r in st.session_state.rows)

    records = []
    for r in st.session_state.rows:
        row = {
            "Question Name": r["Question Name"],
            "Question prompt": r["Question prompt"],
        }
        for i in range(max_choices):
            row[f"Choice {i+1}"] = r["choices"][i] if i < len(r["choices"]) else ""
        records.append(row)

    df = pd.DataFrame(records)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.download_button(
        label="⬇️ Download Excel",
        data=to_excel_bytes(df),
        file_name="questions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("🗑️ Delete all saved questions"):
        st.session_state.rows = []
        st.success("Cleared.")
        st.rerun()
