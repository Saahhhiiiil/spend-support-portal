from pathlib import Path
import io
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile

import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
PRODUCTION_SCRIPT = BASE_DIR / "spend_support_26(1).py"
QC_SCRIPT = BASE_DIR / "spend_support_qc_26.py"
GZIP_SCRIPT = BASE_DIR / "csv_gz.py"

st.set_page_config(
    page_title="Spend/Support Portal",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
.block-container {max-width: 1180px; padding-top: 1.8rem; padding-bottom: 3rem;}
.hero {padding: 1.6rem 1.8rem; border-radius: 16px; background: linear-gradient(115deg,#14213d,#294b75); color:#fff; margin-bottom:1.2rem;}
.hero h1 {color:#fff; margin:0 0 .35rem 0; font-size:2rem;}
.hero p {color:#e1eafa; margin:0;}
.step {border:1px solid rgba(120,130,150,.3); padding:1rem; border-radius:14px; min-height:130px;}
.stepno {font-size:.76rem; font-weight:800; letter-spacing:.08em; color:#628bd3;}
.muted {color:#788397; font-size:.9rem;}
div[data-testid="stFileUploader"] {border-radius:12px;}
</style>
<div class="hero">
  <h1>Spend/Support Processing Portal</h1>
  <p>Production → QC → GZIP</p>
</div>
""", unsafe_allow_html=True)


st.markdown("### Workflow")
c1, c2, c3 = st.columns(3)
for col, no, title, desc in [
    (c1, "01", "Spend/Support Production", "Generate the flat-file CSV using the existing production script."),
    (c2, "02", "Quality Control", "Run the existing QC script against the ADS, mapping and generated CSV."),
    (c3, "03", "GZIP Compression", "Compress the generated production CSV into a .csv.gz file."),
]:
    with col:
        st.markdown(
            f'<div class="step"><div class="stepno">STEP {no}</div>'
            f'<h3>{title}</h3><div class="muted">{desc}</div></div>',
            unsafe_allow_html=True,
        )

st.markdown("### 1. Upload input files")
left, right = st.columns(2)
with left:
    ads_file = st.file_uploader("ADS file (CSV)", type=["csv"], key="ads_upload")
with right:
    mapping_file = st.file_uploader(
        "Variable Mapping file (Excel)", type=["xlsx", "xlsm", "xls"], key="mapping_upload"
    )

if ads_file:
    st.caption(f"ADS selected: {ads_file.name} · {ads_file.size / (1024**2):,.2f} MB")
if mapping_file:
    st.caption(f"Mapping selected: {mapping_file.name} · {mapping_file.size / (1024**2):,.2f} MB")

st.markdown("### 2. Run the complete workflow")
st.write(
    "Select Run to execute Production, QC and GZIP in sequence. The portal passes "
    "the generated production CSV directly to QC and GZIP; you do not need to upload it again."
)
st.caption("Your output files will be available to download when processing is complete.")

if "workflow_result" not in st.session_state:
    st.session_state.workflow_result = None
if "workflow_logs" not in st.session_state:
    st.session_state.workflow_logs = []

run_clicked = st.button(
    "Run Production → QC → GZIP",
    type="primary",
    use_container_width=True,
    disabled=not (ads_file and mapping_file),
)


def run_script_with_progress(
    script_path: Path,
    input_text: str,
    cwd: Path,
    progress_bar,
    status_box,
    stage_name: str,
    start_percent: int,
    end_percent: int,
    expected_seconds: float,
):
    """Run an existing interactive script and update an estimated workflow percentage.

    The scripts do not expose internal progress callbacks, so the percentage advances
    smoothly while each stage is running and reaches the stage boundary on completion.
    """
    log_path = cwd / f"{script_path.stem}_execution.log"
    started = time.monotonic()
    with log_path.open("w", encoding="utf-8", errors="replace") as log_file:
        process = subprocess.Popen(
            [sys.executable, str(script_path)],
            stdin=subprocess.PIPE,
            stdout=log_file,
            stderr=subprocess.STDOUT,
            text=True,
            cwd=str(cwd),
        )
        try:
            process.stdin.write(input_text)
            process.stdin.close()
        except Exception:
            process.kill()
            process.wait()
            raise

        while process.poll() is None:
            elapsed = time.monotonic() - started
            fraction = min(elapsed / max(expected_seconds, 1), 0.96)
            current = min(end_percent - 1, start_percent + int((end_percent - start_percent) * fraction))
            progress_bar.progress(current, text=f"{current}% · {stage_name}")
            status_box.info(f"{stage_name} is running… {current}% complete")
            time.sleep(0.4)

    output = log_path.read_text(encoding="utf-8", errors="replace") if log_path.exists() else ""
    return subprocess.CompletedProcess(
        args=[sys.executable, str(script_path)],
        returncode=process.returncode,
        stdout=output,
        stderr="",
    )


if run_clicked:
    st.session_state.workflow_result = None
    st.session_state.workflow_logs = []

    missing = [p.name for p in (PRODUCTION_SCRIPT, QC_SCRIPT, GZIP_SCRIPT) if not p.exists()]
    if missing:
        st.error("Required script(s) missing from the deployment package: " + ", ".join(missing))
    else:
        work_dir = Path(tempfile.mkdtemp(prefix="spend_support_run_"))
        ads_path = work_dir / Path(ads_file.name).name
        mapping_path = work_dir / Path(mapping_file.name).name
        ads_path.write_bytes(ads_file.getvalue())
        mapping_path.write_bytes(mapping_file.getvalue())
        output_dir = work_dir / "outputs"
        output_dir.mkdir()

        logs = []
        progress = st.progress(0, text="0% · Preparing workflow")
        status = st.empty()
        try:
            # STEP 1 — Production. Existing script and prompts are preserved.
            status.info("Step 1 of 3 — Generating Spend/Support flat file…")
            prod_inputs = f"{output_dir}\n{ads_path}\n{mapping_path}\n"
            prod = run_script_with_progress(
                PRODUCTION_SCRIPT, prod_inputs, work_dir, progress, status,
                "Production", 2, 34, 90,
            )
            logs.append(("Spend/Support Production", prod.stdout, prod.stderr, prod.returncode))
            if prod.returncode != 0:
                raise RuntimeError(
                    f"Production failed (exit code {prod.returncode}). Open Execution logs below for the full output."
                )

            parts = ads_path.stem.split("_")
            if len(parts) < 2 or not parts[1].strip():
                raise RuntimeError(
                    "Could not detect LOB from ADS filename. Expected a name like USAA_P&C_Master_ADS.csv."
                )
            lob = parts[1].strip()
            flat_path = output_dir / f"USAA_{lob}_Spend_Support_Flat_File.csv"
            if not flat_path.exists():
                raise RuntimeError(
                    f"Production returned successfully, but expected output was not found: {flat_path.name}"
                )
            progress.progress(34, text="34% · Production complete")

            # STEP 2 — QC. Existing QC script and prompts are preserved.
            status.info("Step 2 of 3 — Running Spend/Support QC…")
            qc_inputs = f"{ads_path}\n{mapping_path}\n{flat_path}\n"
            qc = run_script_with_progress(
                QC_SCRIPT, qc_inputs, work_dir, progress, status,
                "Quality Control", 34, 68, 120,
            )
            logs.append(("Spend/Support QC", qc.stdout, qc.stderr, qc.returncode))
            qc_path = work_dir / f"USAA_{lob}_Spend_Support_Flat_File_QC.xlsx"
            if qc.returncode != 0:
                raise RuntimeError(
                    f"QC failed (exit code {qc.returncode}). The full traceback/output is available in Execution logs below."
                )
            if not qc_path.exists():
                raise RuntimeError(
                    "QC process ended without an error code, but the expected QC workbook was not created. "
                    "Check Execution logs below."
                )
            progress.progress(68, text="68% · Quality Control complete")

            # STEP 3 — GZIP. Existing compression script and prompt are preserved.
            status.info("Step 3 of 3 — Compressing production CSV…")
            gz = run_script_with_progress(
                GZIP_SCRIPT, f"{flat_path}\n", work_dir, progress, status,
                "GZIP Compression", 68, 99, 30,
            )
            logs.append(("GZIP Compression", gz.stdout, gz.stderr, gz.returncode))
            gz_path = Path(str(flat_path) + ".gz")
            if gz.returncode != 0 or not gz_path.exists():
                raise RuntimeError(
                    f"GZIP failed (exit code {gz.returncode}). Open Execution logs below for details."
                )
            progress.progress(100, text="100% · Workflow complete")
            status.success("All three steps completed successfully.")

            original_size = flat_path.stat().st_size
            compressed_size = gz_path.stat().st_size
            reduction = (1 - compressed_size / original_size) * 100 if original_size else 0
            with flat_path.open("rb") as f:
                rows = max(sum(1 for _ in f) - 1, 0)

            st.session_state.workflow_result = {
                "flat_bytes": flat_path.read_bytes(),
                "qc_bytes": qc_path.read_bytes(),
                "gz_bytes": gz_path.read_bytes(),
                "flat_name": flat_path.name,
                "qc_name": qc_path.name,
                "gz_name": gz_path.name,
                "lob": lob,
                "rows": rows,
                "original_mb": original_size / (1024**2),
                "compressed_mb": compressed_size / (1024**2),
                "reduction": reduction,
            }
        except Exception as exc:
            status.error(str(exc))
        finally:
            st.session_state.workflow_logs = logs
            # Files have been copied into session state on success; on failure,
            # preserve the complete traceback in logs and remove temporary inputs.
            shutil.rmtree(work_dir, ignore_errors=True)

result = st.session_state.workflow_result
if result:
    st.markdown("---")
    st.markdown("### 3. Download your outputs")
    st.success(f"Workflow completed for LOB: {result['lob']}")
    a, b, c, d = st.columns(4)
    a.metric("Production rows", f"{result['rows']:,}")
    b.metric("Original CSV", f"{result['original_mb']:.2f} MB")
    c.metric("GZIP file", f"{result['compressed_mb']:.2f} MB")
    d.metric("Size reduction", f"{result['reduction']:.1f}%")

    st.download_button(
        f"Download production CSV · {result['flat_name']}",
        data=result["flat_bytes"], file_name=result["flat_name"],
        mime="text/csv", use_container_width=True,
    )
    st.download_button(
        f"Download QC workbook · {result['qc_name']}",
        data=result["qc_bytes"], file_name=result["qc_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.download_button(
        f"Download GZIP file · {result['gz_name']}",
        data=result["gz_bytes"], file_name=result["gz_name"],
        mime="application/gzip", use_container_width=True,
    )

    bundle = io.BytesIO()
    with zipfile.ZipFile(bundle, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(result["flat_name"], result["flat_bytes"])
        zf.writestr(result["qc_name"], result["qc_bytes"])
        zf.writestr(result["gz_name"], result["gz_bytes"])
    st.download_button(
        "Download all outputs as one ZIP",
        data=bundle.getvalue(),
        file_name=f"USAA_{result['lob']}_Spend_Support_Outputs.zip",
        mime="application/zip", use_container_width=True,
    )
    st.caption("Use your browser's download options to choose where to save the files.")

logs = st.session_state.workflow_logs
if logs:
    with st.expander("Execution logs (including QC traceback)", expanded=any(code != 0 for _, _, _, code in logs)):
        for name, stdout, stderr, code in logs:
            st.markdown(f"**{name} — {'SUCCESS' if code == 0 else 'FAILED'} (exit code {code})**")
            if stdout:
                st.code(stdout[-30000:], language="text")
            if stderr:
                st.code(stderr[-30000:], language="text")
            if not stdout and not stderr:
                st.caption("No console output was returned by this step.")

st.markdown("---")
st.caption(
    "The existing production, QC and GZIP scripts are used to generate your output files."
)
