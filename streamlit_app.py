import io
import json
import os
import tempfile
import zipfile
from contextlib import redirect_stderr, redirect_stdout

import pandas as pd
import streamlit as st
from google.cloud import storage

import SingleStudy_workflow
from SingleStudy_workflow import _DEDUCE_SEX, _detect_bones_from_measurement

# ---------------------------------------------------------------------------
# Config — set GCS_BUCKET env var in Cloud Run, or rename this string
# ---------------------------------------------------------------------------
GCS_BUCKET = os.environ.get("GCS_BUCKET", "pb-workflow-store")


# ---------------------------------------------------------------------------
# GCS helpers
# ---------------------------------------------------------------------------

def _bucket():
    return storage.Client().bucket(GCS_BUCKET)


def load_studies_config():
    if "studies_config" not in st.session_state:
        try:
            blob = _bucket().blob("studies_config.json")
            st.session_state.studies_config = (
                json.loads(blob.download_as_text()) if blob.exists() else {}
            )
        except Exception:
            st.session_state.studies_config = {}
    return st.session_state.studies_config


def save_studies_config(config):
    st.session_state.studies_config = config
    try:
        _bucket().blob("studies_config.json").upload_from_string(
            json.dumps(config, indent=4), content_type="application/json"
        )
    except Exception as e:
        st.sidebar.warning(f"Cloud save failed: {e}")


def download_master(study_name, sex, age, bone_type, dest_dir):
    filename = f"{study_name}_{sex}_{age}_{bone_type}.xlsx"
    try:
        blob = _bucket().blob(f"masters/{filename}")
        if blob.exists():
            blob.download_to_filename(os.path.join(dest_dir, filename))
            return True
    except Exception:
        pass
    return False


def upload_masters(src_dir):
    try:
        bucket = _bucket()
        for fname in os.listdir(src_dir):
            if fname.endswith(".xlsx"):
                bucket.blob(f"masters/{fname}").upload_from_filename(
                    os.path.join(src_dir, fname)
                )
    except Exception as e:
        st.warning(f"Master upload to cloud failed: {e}")


# ---------------------------------------------------------------------------
# General helpers
# ---------------------------------------------------------------------------

def save_upload(upload, folder):
    if not upload:
        return ""
    path = os.path.join(folder, upload.name)
    upload.seek(0)
    with open(path, "wb") as f:
        f.write(upload.read())
    return path


def collect_outputs(base_dir):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(base_dir):
            for file in files:
                if file.lower().endswith(".txt"):
                    continue
                full = os.path.join(root, file)
                zf.write(full, os.path.relpath(full, base_dir))
    buf.seek(0)
    return buf


def run_captured(func, inputs):
    buf = io.StringIO()
    try:
        with redirect_stdout(buf), redirect_stderr(buf):
            success = func(inputs)
    except Exception as e:
        success = False
        buf.write(f"\nUnhandled error: {e}")
    return success, buf.getvalue()


def build_group_map(df, is_mixed):
    group_map = {}
    for _, row in df.iterrows():
        prefix = str(row.get("Prefix", "")).strip().upper()
        group_name = str(row.get("Group Name", "")).strip()
        if not prefix:
            continue
        if is_mixed:
            group_map[prefix] = {"group": group_name, "sex": str(row.get("Sex", "Male"))}
        else:
            group_map[prefix] = group_name
    return group_map


def group_map_to_df(group_map, is_mixed):
    if is_mixed:
        rows = [
            {
                "Prefix": p,
                "Sex": v.get("sex", "Male") if isinstance(v, dict) else "Male",
                "Group Name": v.get("group", "") if isinstance(v, dict) else v,
            }
            for p, v in group_map.items()
        ]
        return pd.DataFrame(rows, columns=["Prefix", "Sex", "Group Name"])
    rows = [
        {"Prefix": p, "Group Name": v.get("group", "") if isinstance(v, dict) else v}
        for p, v in group_map.items()
    ]
    return pd.DataFrame(rows, columns=["Prefix", "Group Name"])


# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Data Analysis Automated Workflows", layout="wide")

# ---------------------------------------------------------------------------
# Password gate
# ---------------------------------------------------------------------------
if not st.session_state.get("authenticated"):
    st.title("Data Analysis Automated Workflows")
    app_pwd = os.environ.get("APP_PASSWORD", "").strip()
    if not app_pwd:
        st.error("APP_PASSWORD is not configured. Contact your administrator.")
        st.stop()
    pwd = st.text_input("Access Password", type="password")
    if st.button("Submit"):
        if pwd.strip() == app_pwd:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

st.title("Data Analysis Automated Workflows")

# ---------------------------------------------------------------------------
# Top-level tool selector
# ---------------------------------------------------------------------------
tool = st.sidebar.radio(
    "Select Analysis Tool",
    ["3-Point Bending", "BioDent (Coming Soon)"],
)

if tool == "BioDent (Coming Soon)":
    st.header("BioDent Analysis")
    st.info("BioDent workflow is not yet available. Check back soon.")
    st.stop()

# ===========================================================================
# 3-Point BENDING — SINGLE STUDY
# ===========================================================================

    # -----------------------------------------------------------------------
    # Sidebar — Study Manager
    # -----------------------------------------------------------------------
    config = load_studies_config()
    st.sidebar.subheader("Study Manager")
    study_names = list(config.keys())
    selected = st.sidebar.selectbox("Saved Studies", ["(New Study)"] + study_names)

    load_col, del_col = st.sidebar.columns(2)

    if load_col.button("Load", disabled=(selected == "(New Study)")):
        d = config[selected]
        is_m = d.get("sex", "Male") == "Mixed"
        st.session_state["f_name"] = selected
        st.session_state["f_sex"] = d.get("sex", "Male")
        st.session_state["f_age"] = d.get("age", "16 Weeks")
        st.session_state["f_bone"] = d.get("bone", "Femur")
        st.session_state["f_auto_csv"] = d.get("auto_csv_subfolder", True)
        st.session_state["f_anat"] = d.get("anatomical_diameters", False)
        st.session_state["ss_group_df"] = group_map_to_df(d.get("group_map", {}), is_m)
        st.session_state["ss_prev_mixed"] = is_m
        st.rerun()

    if del_col.button("Delete", disabled=(selected == "(New Study)")):
        del config[selected]
        save_studies_config(config)
        st.rerun()

    # -----------------------------------------------------------------------
    # Main form
    # -----------------------------------------------------------------------
    st.header("Single Study")

    # Initialise session state defaults
    for key, default in [("f_name", ""), ("f_age", "16 Weeks")]:
        st.session_state.setdefault(key, default)

    c1, c2, c3 = st.columns(3)
    study_name = c1.text_input("Study Name", key="f_name", placeholder="e.g. IFS+SHP099 2026")
    sex = c2.selectbox("Cohort Sex", ["Male", "Female", "Mixed"], key="f_sex")
    age = c3.text_input("Cohort Age", key="f_age")

    c4, c5 = st.columns(2)
    bone = c4.selectbox("Target Bone", ["Femur", "Tibia", "Humerus"], key="f_bone")
    auto_csv = c5.checkbox(
        "Auto-create StudyName_CSVFiles subfolder",
        key="f_auto_csv",
        value=st.session_state.get("f_auto_csv", True),
    )
    anat_diam = st.checkbox(
        "Generate Anatomical Diameter Folders",
        key="f_anat",
        value=st.session_state.get("f_anat", False),
    )

    if st.sidebar.button("Save Study", disabled=not study_name):
        is_m = sex == "Mixed"
        gm_df = st.session_state.get("ss_group_df", pd.DataFrame())
        config[study_name] = {
            "sex": sex,
            "age": age,
            "bone": bone,
            "auto_csv_subfolder": auto_csv,
            "anatomical_diameters": anat_diam,
            "group_map": build_group_map(gm_df, is_m),
        }
        save_studies_config(config)
        st.sidebar.success("Saved!")

    # -----------------------------------------------------------------------
    # Group map editor
    # -----------------------------------------------------------------------
    st.subheader("Group Map")
    is_mixed = sex == "Mixed"

    if st.session_state.get("ss_prev_mixed") != is_mixed:
        st.session_state.ss_group_df = (
            pd.DataFrame(columns=["Prefix", "Sex", "Group Name"])
            if is_mixed
            else pd.DataFrame(columns=["Prefix", "Group Name"])
        )
        st.session_state.ss_prev_mixed = is_mixed

    st.session_state.setdefault(
        "ss_group_df",
        pd.DataFrame(columns=["Prefix", "Sex", "Group Name"] if is_mixed else ["Prefix", "Group Name"]),
    )

    col_cfg = {"Prefix": st.column_config.TextColumn("Prefix", width="small")}
    if is_mixed:
        col_cfg["Sex"] = st.column_config.SelectboxColumn(
            "Sex", options=["Male", "Female", _DEDUCE_SEX], default="Male", width="medium"
        )
    col_cfg["Group Name"] = st.column_config.TextColumn("Group Name")

    edited_gm = st.data_editor(
        st.session_state.ss_group_df,
        column_config=col_cfg,
        num_rows="dynamic",
        use_container_width=True,
        key=f"gm_editor_{is_mixed}",
    )

    # -----------------------------------------------------------------------
    # File uploaders
    # -----------------------------------------------------------------------
    st.subheader("Data Files")
    st.caption("Tip: zip your raw data folder (right-click → Send to → Compressed folder) and upload the zip.")
    raw_zip = st.file_uploader(
        "Raw Data Folder (.zip)", type="zip", key="ss_raw_zip"
    )
    meas_upload = st.file_uploader(
        "Measurement Excel File", type=["xlsx", "xls"], key="ss_meas"
    )

    measurement_sheet = None
    if meas_upload:
        try:
            meas_upload.seek(0)
            sheets = pd.ExcelFile(meas_upload).sheet_names
            meas_upload.seek(0)
            if len(sheets) > 1:
                measurement_sheet = st.selectbox("Measurement Sheet", sheets)
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # Run
    # -----------------------------------------------------------------------
    if st.button("Execute Workflow", type="primary", disabled=not (study_name and raw_zip)):
        group_map = build_group_map(edited_gm, is_mixed)
        if not group_map:
            st.error("Group Map is empty — add at least one prefix row before running.")
            st.stop()

        with tempfile.TemporaryDirectory() as tmpdir:
            data_folder = os.path.join(tmpdir, "raw_data")
            os.makedirs(data_folder)
            raw_zip.seek(0)
            with zipfile.ZipFile(io.BytesIO(raw_zip.read())) as zf:
                for member in zf.namelist():
                    if member.lower().endswith(".txt"):
                        # Flatten into one folder — Single Study scans one directory only
                        dest = os.path.join(data_folder, os.path.basename(member))
                        with zf.open(member) as src, open(dest, "wb") as out:
                            out.write(src.read())

            meas_path = save_upload(meas_upload, tmpdir) if meas_upload else ""
            master_dir = os.path.dirname(meas_path) if meas_path else tmpdir
            csv_out = os.path.join(tmpdir, "output")
            os.makedirs(csv_out)

            # Pull existing masters from GCS so data accumulates across sessions
            with st.spinner("Checking cloud storage for existing master files…"):
                detected = _detect_bones_from_measurement(meas_path, measurement_sheet)
                bones_to_fetch = detected if detected else [bone]
                fetched = [
                    bt for bt in bones_to_fetch
                    if download_master(study_name, sex, age, bt, master_dir)
                ]
            if fetched:
                st.info(f"Loaded existing master(s) from cloud: {', '.join(fetched)}")

            inputs = {
                "study_name": study_name,
                "data_folder": data_folder,
                "measurement_path": meas_path,
                "csv_out_dir": csv_out,
                "group_map": group_map,
                "sex": sex,
                "age": age,
                "bone": bone,
                "anatomical_diameters": anat_diam,
                "measurement_sheet": measurement_sheet,
                "auto_csv_subfolder": auto_csv,
            }

            with st.spinner("Running workflow…"):
                success, log = run_captured(SingleStudy_workflow.run_workflow, inputs)

            with st.spinner("Saving master files to cloud storage…"):
                upload_masters(master_dir)

            if success:
                st.success("Workflow completed successfully.")
            else:
                st.error("Workflow encountered issues — see log below.")

            with st.expander("Workflow Log", expanded=not success):
                st.text(log)

            st.download_button(
                "⬇ Download All Results (.zip)",
                data=collect_outputs(tmpdir),
                file_name=f"{study_name.replace(' ', '_')}_results.zip",
                mime="application/zip",
            )
