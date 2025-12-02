# label_app.py — v2.12.5
import streamlit as st
import subprocess, os, tempfile, shutil, sys

PYTHON_SCRIPT = "generate_labels_all.py"
OUTPUT_DIR = "output_pdfs"

st.set_page_config(page_title="Tem Nhãn Tân Hòa (v2.12.5)", layout="centered")
st.title("🏷️ Trình Xuất Tem Nhãn Tân Hòa (v2.12.5)")

st.markdown("---")
st.header("📂 Chọn file đầu vào")
uploaded_excel = st.file_uploader("Tải file Excel (.xlsx):", type=["xlsx"])
uploaded_pdf = st.file_uploader("Tải file PDF (.pdf):", type=["pdf"])

excel_path = None; pdf_path = None
if uploaded_excel:
    t = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    t.write(uploaded_excel.read()); t.flush(); excel_path = t.name
if uploaded_pdf:
    t = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    t.write(uploaded_pdf.read()); t.flush(); pdf_path = t.name

st.markdown("---")
mode = st.radio("⚙️ Cách lấy khoảng:", [
    "Mặc định (tự tách file 500 tem)",
    "Lấy khoảng theo Excel (chạy nối tiếp)",
    "Tự điền khoảng số"
], index=0)

export_mode = st.radio("Chọn loại tem cần xuất:", ["Xuất ColorLabel (đỏ)", "Xuất Hangtag (xanh)", "Xuất cả 2"], index=0)

mode2 = st.radio("🎯 Chế độ chọn mã:", ["Xuất tất cả mã", "Xuất mã cụ thể"], index=0)
codes = ""
if mode2 == "Xuất mã cụ thể":
    codes = st.text_input("🔢 Nhập mã (VD: C207720 hoặc C207720,C207721):", "")

manual_from = manual_to = None
if mode == "Tự điền khoảng số":
    c1, c2 = st.columns(2)
    with c1:
        manual_from = st.number_input("Từ số (inclusive)", min_value=1, value=1, step=1)
    with c2:
        manual_to = st.number_input("Đến số (inclusive)", min_value=1, value=100, step=1)

run1, run2 = st.columns([1,1])
with run1:
    run_button = st.button("🚀 BẮT ĐẦU XUẤT TEM")
with run2:
    open_button = st.button("📂 Mở thư mục output")

if open_button:
    try:
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR, exist_ok=True)
        if os.name == "nt":
            os.startfile(OUTPUT_DIR)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", OUTPUT_DIR])
        else:
            subprocess.Popen(["xdg-open", OUTPUT_DIR])
    except Exception as e:
        st.error(f"Không mở được thư mục: {e}")

if run_button:
    if not uploaded_excel or not uploaded_pdf:
        st.error("⚠️ Vui lòng tải cả file Excel và file PDF!")
    else:
        cmd = ["python", PYTHON_SCRIPT, os.path.basename(uploaded_excel.name), os.path.basename(uploaded_pdf.name)]
        if mode == "Lấy khoảng theo Excel (chạy nối tiếp)":
            cmd += ["--range-from-excel"]
        elif mode == "Tự điền khoảng số":
            if not (manual_from and manual_to and manual_from <= manual_to):
                st.error("Vui lòng nhập khoảng hợp lệ (Từ <= Đến).")
                st.stop()
            cmd.append(f"--manual-range={int(manual_from)}-{int(manual_to)}")

        if export_mode == "Xuất ColorLabel (đỏ)":
            cmd += ["--export=color"]
        elif export_mode == "Xuất Hangtag (xanh)":
            cmd += ["--export=hangtag"]
        else:
            cmd += ["--export=both"]

        if mode2 == "Xuất mã cụ thể" and codes.strip():
            cmd.append(codes.strip())

        if excel_path: shutil.copy(excel_path, os.path.basename(uploaded_excel.name))
        if pdf_path:   shutil.copy(pdf_path, os.path.basename(uploaded_pdf.name))

        st.info("⏳ Đang xử lý, vui lòng chờ...")
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            st.success("✅ Hoàn tất xuất tem!")
            st.code(result.stdout)
        except subprocess.CalledProcessError as e:
            st.error("❌ Có lỗi khi chạy script.")
            out = (e.stdout or "") + "\n\n" + (e.stderr or "")
            st.code(out)

st.markdown("---")
st.caption("• 'Mặc định': tự tách 500 tem / file theo từng dòng Excel. • 'Theo Excel': dùng From/To (nếu trống sẽ dồn nối tiếp). • 'Tự điền khoảng': 1 file/mã theo khoảng nhập tay.")
