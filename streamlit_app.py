from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from server import process_docx


st.set_page_config(
    page_title="DOCX Reference Merger",
    page_icon="DOCX",
    layout="wide",
)


st.title("Gộp DOCX và đánh số lại tài liệu tham khảo")
st.caption(
    "Upload file DOCX theo đúng thứ tự cần ghép. App sẽ tách phần tài liệu tham khảo, "
    "đánh lại citation trong thân bài, rồi chia danh mục cuối thành tài liệu tiếng Việt và tiếng Anh."
)


with st.sidebar:
    st.header("Thiết lập")
    file_count = st.number_input("Số file cần xử lý", min_value=1, max_value=30, value=1, step=1)
    st.info(
        "Ghép file: dùng toàn bộ file theo thứ tự upload.\n\n"
        "Tách Việt/Anh: dùng tốt nhất với 1 file, nhưng vẫn xử lý được nhiều file."
    )


uploaded_files = []
st.subheader("Upload DOCX")
for index in range(1, int(file_count) + 1):
    uploaded_file = st.file_uploader(
        f"File {index}",
        type=["docx"],
        key=f"docx_file_{index}",
    )
    uploaded_files.append(uploaded_file)


def run_processing(mode: str) -> None:
    selected_files = [file for file in uploaded_files if file is not None]
    if len(selected_files) != int(file_count):
        st.warning(f"Bạn đã chọn {len(selected_files)}/{int(file_count)} file. Hãy chọn đủ file trước khi xử lý.")
        return

    with tempfile.TemporaryDirectory(prefix="streamlit_docx_ref_") as tmp:
        tmp_dir = Path(tmp)
        input_paths = []
        for index, uploaded_file in enumerate(selected_files, start=1):
            safe_name = Path(uploaded_file.name).name
            input_path = tmp_dir / f"{index:03d}_{safe_name}"
            input_path.write_bytes(uploaded_file.getbuffer())
            input_paths.append(input_path)

        output_name = "tach_tai_lieu_viet_anh.docx" if mode == "split" else "ghep_docx_tai_lieu_tham_khao.docx"
        output_path = tmp_dir / output_name
        report = process_docx(input_paths, output_path)
        output_bytes = output_path.read_bytes()

    st.session_state["docx_output_bytes"] = output_bytes
    st.session_state["docx_output_name"] = output_name
    st.session_state["docx_report"] = report


col1, col2 = st.columns([1, 1])
with col1:
    if st.button("Ghép file", type="primary", use_container_width=True):
        run_processing("merge")
with col2:
    if st.button("Tách tài liệu tiếng Việt/Anh", use_container_width=True):
        run_processing("split")


if "docx_output_bytes" in st.session_state:
    st.success("Đã xử lý xong.")
    st.download_button(
        "Tải xuống DOCX",
        data=st.session_state["docx_output_bytes"],
        file_name=st.session_state["docx_output_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    with st.expander("Báo cáo xử lý", expanded=True):
        st.text(st.session_state["docx_report"])


with st.expander("Ghi chú"):
    st.markdown(
        "- Thứ tự upload chính là thứ tự ghép.\n"
        "- Citation dạng `[1]`, `[1, 2]`, `[1-3]` sẽ được đổi theo số mới.\n"
        "- Phần tham khảo được nhận diện từ heading bắt đầu bằng `Tài liệu tham khảo`.\n"
        "- Nếu citation không có mục tham khảo tương ứng, app giữ nguyên citation đó và ghi cảnh báo trong báo cáo."
    )
