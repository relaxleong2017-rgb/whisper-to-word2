import io
import os
import subprocess
import tempfile

import streamlit as st
import whisper
from docx import Document

st.set_page_config(page_title="MP3 語音轉 Word 工具", page_icon="📝", layout="centered")

st.title("📝 MP3 語音轉 Word (.docx) 工具")
st.caption("支援上傳 MP3 或透過 Google Drive 連結下載轉錄，並匯出 Word 檔")


def extract_gdrive_file_id(url_or_id: str) -> str:
    """從 Google Drive 分享連結或檔案 ID 中擷取 file_id"""
    text = url_or_id.strip()
    if not text:
        return ""
    # 常見格式: https://drive.google.com/file/d/<FILE_ID>/view?usp=sharing
    if "/file/d/" in text:
        return text.split("/file/d/")[1].split("/")[0]
    # 常見格式: https://drive.google.com/open?id=<FILE_ID> 或 uc?id=<FILE_ID>
    if "id=" in text:
        return text.split("id=")[1].split("&")[0]
    return text


def download_from_gdrive(file_id: str, output_mp3_path: str) -> None:
    """使用 gdown 下載公開或已授權分享的 Google Drive 檔案"""
    try:
        import gdown
    except ImportError:
        subprocess.check_call(["pip", "install", "-q", "gdown"])
        import gdown

    url = f"https://drive.google.com/uc?id={file_id}"
    downloaded = gdown.download(url=url, output=output_mp3_path, quiet=False)
    if not downloaded or not os.path.exists(output_mp3_path) or os.path.getsize(output_mp3_path) == 0:
        raise RuntimeError(
            "無法從 Google Drive 下載檔案。請確認檔案分享權限已設為「知道連結的任何人皆可檢視」，且連結正確。"
        )


@st.cache_resource(show_spinner="正在載入 Whisper 語音模型（首次載入需較長時間）...")
def load_whisper_model(model_size: str):
    """快取 Whisper 模型，避免每次互動重新載入"""
    return whisper.load_model(model_size)


def build_docx_bytes(source_name: str, transcript_text: str) -> bytes:
    """將逐字稿建立為 .docx 並回傳位元組資料"""
    doc = Document()
    doc.add_heading("語音轉文字逐字稿", level=1)
    doc.add_paragraph(f"來源檔案：{source_name}")
    doc.add_paragraph("")
    doc.add_paragraph(transcript_text.strip())

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# --- 側邊欄設定 ---
with st.sidebar:
    st.header("辨識設定")
    model_size = st.selectbox(
        "Whisper 模型大小",
        options=["tiny", "base", "small", "medium"],
        index=1,
        help="雲端免費主機建議選 tiny / base / small；medium 較準確但記憶體與時間需求較高。"
    )
    language_option = st.selectbox(
        "語音語言",
        options=["zh（中文）", "en（英文）", "ja（日文）", "auto（自動偵測）"],
        index=0
    )
    language_map = {
        "zh（中文）": "zh",
        "en（英文）": "en",
        "ja（日文）": "ja",
        "auto（自動偵測）": None,
    }
    selected_lang = language_map[language_option]

# --- 音訊來源選擇 ---
source_mode = st.radio(
    "選擇 MP3 來源方式",
    options=["直接上傳 MP3 檔案", "使用 Google Drive 分享連結"],
    horizontal=True
)

uploaded_file = None
gdrive_input = ""

if source_mode == "直接上傳 MP3 檔案":
    uploaded_file = st.file_uploader("請選擇 MP3 語音檔", type=["mp3"])
else:
    gdrive_input = st.text_input(
        "輸入 Google Drive 檔案連結或 File ID",
        placeholder="例如：https://drive.google.com/file/d/xxxx/view?usp=sharing"
    )
    st.info("提醒：雲端伺服器無法直接掛載你的私人 Google Drive，請先將 MP3 權限設為「知道連結的任何人皆可檢視」。")

output_filename = st.text_input("輸出的 Word 檔名", value="轉錄結果.docx")
if not output_filename.lower().endswith(".docx"):
    output_filename += ".docx"

# --- 執行轉錄 ---
if st.button("🚀 開始轉錄並產生 Word", type="primary", use_container_width=True):
    with tempfile.TemporaryDirectory() as tmp_dir:
        mp3_path = os.path.join(tmp_dir, "input_audio.mp3")
        source_label = ""

        try:
            if source_mode == "直接上傳 MP3 檔案":
                if uploaded_file is None:
                    st.warning("請先上傳 MP3 檔案。")
                    st.stop()
                source_label = uploaded_file.name
                with open(mp3_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            else:
                file_id = extract_gdrive_file_id(gdrive_input)
                if not file_id:
                    st.warning("請先輸入有效的 Google Drive 連結或 File ID。")
                    st.stop()
                source_label = f"Google Drive ({file_id}).mp3"
                with st.spinner("正在從 Google Drive 下載語音檔..."):
                    download_from_gdrive(file_id, mp3_path)

            model = load_whisper_model(model_size)

            with st.spinner("正在進行語音辨識，請稍候..."):
                initial_prompt = "以下是繁體中文逐字稿：" if selected_lang == "zh" else None
                result = model.transcribe(
                    mp3_path,
                    language=selected_lang,
                    initial_prompt=initial_prompt
                )

            transcript_text = result.get("text", "").strip()
            if not transcript_text:
                st.error("未能辨識出文字內容，請確認語音檔是否有清晰人聲。")
                st.stop()

            docx_bytes = build_docx_bytes(source_label, transcript_text)

            st.success("✅ 語音轉文字完成，Word 檔已準備好！")
            st.subheader("逐字稿預覽")
            st.text_area("辨識結果", value=transcript_text, height=240)

            st.download_button(
                label=f"⬇️ 下載 {output_filename}",
                data=docx_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"發生錯誤：{e}")
