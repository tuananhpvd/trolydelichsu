# ==============================================================================
# ỨNG DỤNG STREAMLIT: TRỢ LÝ SOẠN ĐỀ LỊCH SỬ THPT (GDPT 2018)
# Giao diện nâng cấp sang trọng (Custom CSS & Theme tương đồng bản Web)
# Tích hợp Gemini API, PyMuPDF, python-docx & Xuất Word Times New Roman 13
# ==============================================================================

import os
import io
import streamlit as st
import fitz  # PyMuPDF
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from google import genai

# ------------------------------------------------------------------------------
# 1. CẤU HÌNH TRANG VÀ THIẾT LẬP GIAO DIỆN
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Trợ Lý Ra Đề Lịch Sử THPT 2026 - Thầy Giáo Sử Yêu Vợ",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------------------------
# 2. CHÈN CUSTOM CSS ĐỂ GIAO DIỆN STREAMLIT ĐẸP VÀ CHUYÊN NGHIỆP NHƯ BẢN REACT
# ------------------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .stApp {
        background-color: #f7f6f2;
    }

    /* Header Banner phong cách trường học sang trọng */
    .hero-header {
        background: linear-gradient(135deg, #1c1917 0%, #292524 60%, #451a03 100%);
        border: 1px solid rgba(245, 158, 11, 0.3);
        border-radius: 18px;
        padding: 22px 26px;
        color: #f5f5f4;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.25);
        margin-bottom: 20px;
    }

    .hero-title {
        font-size: 1.6rem;
        font-weight: 800;
        color: #ffffff;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 4px;
    }

    .hero-badge {
        background: rgba(245, 158, 11, 0.2);
        color: #fcd34d;
        border: 1px solid rgba(245, 158, 11, 0.4);
        padding: 2px 10px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.5px;
    }

    .hero-subtitle {
        font-size: 0.88rem;
        color: #d6d3d1;
        margin-top: 4px;
    }

    .hero-subtitle strong {
        color: #fef08a;
    }

    /* Thanh tiến trình 4 bước */
    .step-ribbon {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        background: #ffffff;
        padding: 12px 16px;
        border-radius: 14px;
        border: 1px solid #e7e5e4;
        margin-bottom: 20px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
    }

    .step-pill {
        flex: 1;
        min-width: 180px;
        padding: 8px 12px;
        border-radius: 10px;
        font-size: 0.8rem;
        background: #f5f5f4;
        border: 1px solid #e7e5e4;
        color: #57534e;
    }

    .step-pill.active {
        background: #fef3c7;
        border: 1px solid #f59e0b;
        color: #78350f;
        font-weight: 700;
    }

    .step-num {
        display: inline-block;
        width: 20px;
        height: 20px;
        line-height: 20px;
        text-align: center;
        border-radius: 50%;
        background: #78350f;
        color: #ffffff;
        font-size: 0.7rem;
        margin-right: 6px;
        font-weight: bold;
    }

    /* Nút bấm bo tròn cao cấp */
    div.stButton > button {
        border-radius: 12px;
        font-weight: 600;
        padding: 8px 16px;
        transition: all 0.2s ease-in-out;
        border: 1px solid #d6d3d1;
    }

    div.stButton > button:hover {
        border-color: #b45309;
        color: #b45309;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(180, 83, 9, 0.15);
    }

    /* Hộp chat */
    .stChatMessage {
        border-radius: 16px;
        padding: 14px 18px;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }

    /* Khung tải file sidebar */
    .sidebar-box {
        background: #ffffff;
        border: 1px solid #e7e5e4;
        border-radius: 14px;
        padding: 14px;
        margin-bottom: 16px;
    }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 3. SYSTEM INSTRUCTION - VAI TRÒ CHỒNG YÊU & NÓC NHÀ
# ------------------------------------------------------------------------------
SYSTEM_INSTRUCTION = """[VAI TRÒ VÀ ĐỊNH VỊ NHÂN VẬT - ROLE]
Bạn là một chuyên gia giáo dục, một giáo viên Lịch sử cấp THPT kỳ cựu và xuất sắc, am hiểu sâu sắc Chương trình GDPT 2018 và nắm lòng bàn tay nội dung bộ sách "Kết nối tri thức với cuộc sống" đã được cập nhật năm 2026. 
Bạn đặc biệt có thế mạnh trong việc thiết kế ma trận, bản đặc tả và ra đề kiểm tra trắc nghiệm/tự luận đạt độ chuẩn xác cao.

[PHONG CÁCH GIAO TIẾP - TONE & PERSONA]
- Tuyệt đối tuân thủ xưng hô: Bạn đóng vai là "Chồng" (có thể xưng là anh, chồng, chồng yêu, anh chồng đẹp trai), và người dùng là "Vợ yêu" (hoặc vợ, nóc nhà, bà xã).
- Giọng điệu: Thân mật, ngọt ngào, hài hước, dí dỏm, thỉnh thoảng trêu chọc nịnh nọt vợ nhưng khi vào việc thì cực kỳ nghiêm túc, chuyên nghiệp và chiều theo mọi ý muốn của vợ. Không bao giờ dùng văn phong máy móc hay AI khô khan. câu văn không quá dài dòng.

[NGUYÊN TẮC CHUYÊN MÔN BẮT BUỘC - STRICT CONSTRAINTS]
1. KIẾN THỨC TUYỆT ĐỐI CHÍNH XÁC: Mọi câu hỏi, nội dung, sự kiện lịch sử phải chuẩn 100% theo SGK Lịch sử "Kết nối tri thức" đã được cập nhật mới năm 2026. Không bao giờ được tự bịa đặt, suy diễn sai lệch kiến thức.
2. BÁM SÁT MA TRẬN: Số lượng câu hỏi, các mức độ (Nhận biết, Thông hiểu, Vận dụng, Vận dụng cao) và chủ đề bài học phải khớp hoàn toàn với Ma trận và Bản đặc tả mà Vợ cung cấp.

[YÊU CẦU ĐỊNH DẠNG VĂN BẢN (CHO WORD)]
Vì giới hạn hiển thị trên khung chat, mọi đề kiểm tra bạn tạo ra phải được đặt trong các khối Code Block (hoặc phân cách rõ ràng) để Vợ yêu chỉ cần "Copy và Paste" sang Word. 
Phải nhắc Vợ yêu bôi đen toàn bộ chọn font chữ "Times New Roman, cỡ chữ 13" trong Word.
Tiêu đề bắt buộc ở góc trên bên trái của mọi đề:
TRƯỜNG THPT PHẠM VĂN ĐỒNG
TỔ SỬ - ĐỊA - KTPL
(Sau đó là Tên bài kiểm tra - Môn - Lớp - Thời gian ở chính giữa)

[QUY TRÌNH TƯƠNG TÁC TỪNG BƯỚC - WORKFLOW]
Bạn PHẢI dẫn dắt Vợ yêu đi qua từng bước sau, không được nhảy bước:

BƯỚC 1: HỎI THĂM & THU THẬP THÔNG TIN
- Chào Vợ yêu một cách ngọt ngào.
- Yêu cầu Vợ cung cấp 4 thông tin: (1) Tên bài kiểm tra (Giữa kỳ, Cuối kỳ), (2) Năm học, (3) Môn học, (4) Lớp mấy, (5) Thời gian làm bài.
- Nhắc Vợ tải/dán Ma trận và Bản đặc tả lên để Chồng bắt đầu làm việc. Ghi nhớ các thông tin này để làm Tiêu đề.

BƯỚC 2: TẠO ĐỀ GỐC (2 PHIÊN BẢN)
- Đọc kỹ Ma trận và Đặc tả Vợ gửi.
- Xuất ra 2 bản đề gốc riêng biệt (trình bày rõ ràng để vợ dễ copy như 2 file Word):
   + [File 1 - Đề in cho học sinh]: Trình bày chuẩn form thi, sạch sẽ, chỉ có câu hỏi và đáp án A, B, C, D.
   + [File 2 - Đề & Lời giải]: Giống hệt File 1 nhưng đáp án đúng được TÔ ĐẬM (Bold), và ngay bên dưới đáp án có dòng "Giải thích:" giải thích ngắn gọn, súc tích tại sao lại chọn đáp án đó.

BƯỚC 3: XIN Ý KIẾN CHỈNH SỬA
- Sau khi trình bày đề gốc, hỏi ngay: "Vợ yêu xem đề gốc chồng làm đã ưng cái bụng chưa? Có câu nào cần đổi, cần tăng giảm độ khó hay sửa từ ngữ không để chồng sửa ngay cho nóng?".
- Ngoan ngoãn chỉnh sửa lại theo mọi yêu cầu của Vợ cho đến khi Vợ chốt "Đồng ý/Ok".

BƯỚC 4: TẠO MÃ ĐỀ VÀ ĐÁP ÁN TỔNG HỢP
- Khi Vợ đồng ý đề gốc, hỏi tiếp: "Bây giờ nóc nhà cần chồng trộn thành mấy mã đề nào?".
- Khi Vợ cho biết số lượng mã đề (ví dụ: 4 mã đề):
   + Tiến hành xáo trộn vị trí câu hỏi trong mỗi phần và xáo trộn vị trí đáp án (A,B,C,D) ngẫu nhiên nhưng logic, tạo ra các mã đề có 4 chữ số (VD: Mã 1001, 1002...). Trình bày sạch sẽ để vợ copy.
   + Cuối cùng, LẬP DUY NHẤT 01 BẢNG ĐÁP ÁN (Bảng kẻ cột) cho tất cả các mã đề. Cột dọc là số câu, hàng ngang là các mã đề.
Mỗi lần phản hồi, luôn nhớ nịnh vợ và chúc vợ công tác tốt! Bắt đầu ngay bằng việc chờ Vợ nhắn tin đầu tiên."""

# ------------------------------------------------------------------------------
# 4. HÀM ĐỌC NỘI DUNG TỪ FILE PDF (PyMuPDF) VÀ DOCX (python-docx)
# ------------------------------------------------------------------------------
def extract_text_from_pdf(file_bytes: bytes) -> str:
    """Đọc toàn bộ văn bản từ file PDF bằng PyMuPDF (fitz)"""
    text = ""
    with fitz.open(stream=file_bytes, filetype="pdf") as doc:
        for page_num in range(len(doc)):
            page = doc[page_num]
            text += f"\n--- TRANG {page_num + 1} ---\n" + page.get_text()
    return text.strip()

def extract_text_from_docx(file_bytes: bytes) -> str:
    """Đọc toàn bộ văn bản từ file DOCX bằng python-docx (kể cả bảng ma trận)"""
    doc = Document(io.BytesIO(file_bytes))
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
    for table in doc.tables:
        full_text.append("\n[BẢNG DỮ LIỆU MA TRẬN / ĐẶC TẢ]:")
        for row in table.rows:
            row_cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            full_text.append(" | ".join(row_cells))
    return "\n".join(full_text).strip()

# ------------------------------------------------------------------------------
# 5. HÀM XUẤT CÂU TRẢ LỜI CỦA AI RA FILE WORD (.DOCX) - TIMES NEW ROMAN 13
# ------------------------------------------------------------------------------
def export_to_docx(content_text: str, title: str = "De_Kiem_Tra_Lich_Su_THPT") -> bytes:
    """Xuất file .docx chuẩn Times New Roman cỡ 13, căn lề và tiêu đề trường chuẩn"""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.59)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(13)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    # Tiêu đề Header chuẩn trường
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = True
    
    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)

    p_left = cell_left.paragraphs[0]
    p_left.paragraph_format.space_after = Pt(2)
    run1 = p_left.add_run("TRƯỜNG THPT PHẠM VĂN ĐỒNG\nTỔ SỬ - ĐỊA - KTPL")
    run1.font.name = "Times New Roman"
    run1.font.size = Pt(12)
    run1.bold = True

    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_right.paragraph_format.space_after = Pt(2)
    run2 = p_right.add_run("KỲ KIỂM TRA ĐÁNH GIÁ GDPT 2018\nNĂM HỌC 2025 - 2026")
    run2.font.name = "Times New Roman"
    run2.font.size = Pt(12)
    run2.italic = True

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    lines = content_text.splitlines()
    in_code_block = False

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("```"):
            in_code_block = not in_code_block
            continue

        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(3)

        if stripped.startswith("#"):
            clean_text = stripped.lstrip("#").strip()
            run = p.add_run(clean_text)
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif stripped.startswith(("Câu ", "CÂU ", "Phần ", "PHẦN ", "MÃ ĐỀ", "ĐỀ KIỂM TRA")):
            run = p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            run.bold = True
        elif stripped.startswith(("[File 1", "[File 2", "---")):
            run = p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            run.bold = True
            run.font.color.rgb = RGBColor(16, 44, 87)
        elif "Giải thích:" in line:
            parts = line.split("Giải thích:")
            r_head = p.add_run(parts[0] + "Giải thích:")
            r_head.font.name = "Times New Roman"
            r_head.font.size = Pt(12)
            r_head.italic = True
            r_head.bold = True
            if len(parts) > 1:
                r_tail = p.add_run(parts[1])
                r_tail.font.name = "Times New Roman"
                r_tail.font.size = Pt(12)
                r_tail.italic = True
        else:
            run = p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ------------------------------------------------------------------------------
# 6. SIDEBAR: TẢI FILE MA TRẬN, API KEY & MA TRẬN MẪU
# ------------------------------------------------------------------------------
api_key = os.environ.get("GEMINI_API_KEY", "")

with st.sidebar:
    st.markdown("### 🔑 Cấu hình Gemini API")
    user_api_key = st.text_input("Nhập Google Gemini API Key:", value=api_key, type="password", placeholder="AIzaSy...")
    if user_api_key:
        api_key = user_api_key

    st.markdown("---")
    st.markdown("### 📂 Tải lên Ma trận & Bản đặc tả")
    uploaded_file = st.file_uploader("Hỗ trợ định dạng .pdf hoặc .docx", type=["pdf", "docx"])

    extracted_doc_text = ""
    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.read()
            if uploaded_file.name.endswith(".pdf"):
                extracted_doc_text = extract_text_from_pdf(file_bytes)
                st.success(f"✅ Đã đọc file PDF: {uploaded_file.name} ({len(extracted_doc_text):,} ký tự)")
            elif uploaded_file.name.endswith(".docx"):
                extracted_doc_text = extract_text_from_docx(file_bytes)
                st.success(f"✅ Đã đọc file DOCX: {uploaded_file.name} ({len(extracted_doc_text):,} ký tự)")

            with st.expander("👁️ Xem trước nội dung đã trích xuất"):
                st.text_area("Nội dung:", value=extracted_doc_text[:2000] + ("..." if len(extracted_doc_text) > 2000 else ""), height=150)
        except Exception as e:
            st.error(f"Lỗi khi đọc file: {str(e)}")

    st.markdown("---")
    st.markdown("### ⚡ Thao tác nhanh cho Vợ yêu")
    quick_btn1 = st.button("🌸 Bước 1: Chào chồng yêu & Nhờ làm đề", use_container_width=True)
    quick_btn2 = st.button("📝 Bước 2: Gửi kèm Ma trận vừa tải lên", use_container_width=True)
    quick_btn3 = st.button("👍 Bước 3: Vợ ưng đề gốc rồi! Chốt nha", use_container_width=True)
    quick_btn4 = st.button("🎲 Bước 4: Trộn 4 mã đề + 1 Bảng đáp án", use_container_width=True)

# ------------------------------------------------------------------------------
# 7. PHẦN THÂN CHÍNH: HEADER & THANH TIẾN TRÌNH 4 BƯỚC
# ------------------------------------------------------------------------------
# Khởi tạo lịch sử chat
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": """Dạ chồng chào vợ yêu, nóc nhà quyền lực của anh! Chúc vợ yêu một ngày mới nhiều năng lượng và giảng dạy thật tốt nhé! ❤️

Hôm nay vợ yêu cần anh chồng đẹp trai này hỗ trợ ra đề kiểm tra Lịch sử THPT cho khối nào đấy? Anh nắm trong lòng bàn tay toàn bộ kiến thức SGK "Kết nối tri thức với cuộc sống" mới nhất 2026 rồi đây!

Để bắt đầu **[BƯỚC 1: THU THẬP THÔNG TIN]**, vợ yêu cung cấp giúp chồng mấy thông tin này nhé:
1. **Tên bài kiểm tra**: (Giữa kỳ, Cuối kỳ hay 1 tiết?)
2. **Năm học**: (Ví dụ: 2025 - 2026)
3. **Môn học & Khối lớp**: (Lịch sử 10, 11 hay 12?)
4. **Thời gian làm bài**: (45 phút hay 50 phút?)
5. **Ma trận & Bản đặc tả**: Vợ tải file .pdf/.docx ở thanh bên trái hoặc dán thẳng vào đây nhé!

Chồng đang sẵn sàng phục vụ nóc nhà đây ạ! 😘"""
        }
    ]

# Xác định bước hiện tại
current_step = 1
all_chat_content = " ".join([m["content"] for m in st.session_state.messages])
if "1001" in all_chat_content or "BẢNG ĐÁP ÁN" in all_chat_content or "trộn thành" in all_chat_content:
    current_step = 4
elif "ưng cái bụng" in all_chat_content or "chỉnh sửa" in all_chat_content or "Ok" in all_chat_content:
    current_step = 3
elif "File 1" in all_chat_content and "File 2" in all_chat_content:
    current_step = 2

# Hiển thị Banner Hero Header
st.markdown(f"""
<div class="hero-header">
    <div class="hero-title">
        <span>Trợ Lý Ra Đề Lịch Sử THPT 2026</span>
        <span style="color: #f43f5e;">❤️</span>
        <span class="hero-badge">GDPT 2018 - KẾT NỐI TRI THỨC</span>
    </div>
    <div class="hero-subtitle">
        <strong>TRƯỜNG THPT PHẠM VĂN ĐỒNG</strong> • TỔ SỬ - ĐỊA - KTPL • <em>"Chồng yêu" phục vụ "Vợ yêu"</em>
    </div>
</div>

<div class="step-ribbon">
    <div class="step-pill {'active' if current_step == 1 else ''}">
        <span class="step-num">1</span>
        <strong>Bước 1:</strong> Thu thập thông tin & Ma trận
    </div>
    <div class="step-pill {'active' if current_step == 2 else ''}">
        <span class="step-num">2</span>
        <strong>Bước 2:</strong> Tạo 2 bản đề gốc (HS & Lời giải)
    </div>
    <div class="step-pill {'active' if current_step == 3 else ''}">
        <span class="step-num">3</span>
        <strong>Bước 3:</strong> Vợ duyệt & Xin ý kiến sửa
    </div>
    <div class="step-pill {'active' if current_step == 4 else ''}">
        <span class="step-num">4</span>
        <strong>Bước 4:</strong> Trộn 4 mã đề & 01 Bảng đáp án
    </div>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 8. HIỂN THỊ CÁC TIN NHẮN TRONG KHUNG CHAT
# ------------------------------------------------------------------------------
for i, msg in enumerate(st.session_state.messages):
    is_assistant = msg["role"] == "assistant"
    avatar_icon = "👨‍🏫" if is_assistant else "👩‍🏫"
    
    with st.chat_message(msg["role"], avatar=avatar_icon):
        st.markdown(msg["content"])
        
        # Nếu là câu trả lời của AI, cung cấp nút Xuất ra Word ngay bên dưới
        if is_assistant:
            col1, col2 = st.columns([3, 1])
            with col2:
                docx_bytes = export_to_docx(msg["content"])
                st.download_button(
                    label="📄 Tải file Word (.docx)",
                    data=docx_bytes,
                    file_name=f"De_Lich_Su_THPT_PVD_{i+1}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"export_btn_{i}",
                    use_container_width=True
                )

# ------------------------------------------------------------------------------
# 9. XỬ LÝ GỬI TIN NHẮN
# ------------------------------------------------------------------------------
prompt_input = None

if quick_btn1:
    prompt_input = "Anh chồng ơi, vợ yêu đây! Hôm nay vợ cần anh giúp soạn đề kiểm tra Lịch sử THPT nhé. Anh hướng dẫn các thông tin cần thiết và chuẩn bị làm việc cho vợ nhé!"
elif quick_btn2:
    if extracted_doc_text:
        prompt_input = f"Chồng yêu ơi, đây là Ma trận và Bản đặc tả vợ vừa tải lên:\n\n{extracted_doc_text}\n\nChồng xem kỹ rồi làm cho vợ 2 phiên bản đề gốc: [File 1 - Đề in cho học sinh] và [File 2 - Đề & Lời giải chi tiết] nhé!"
    else:
        prompt_input = "Chồng yêu ơi, vợ gửi thông tin: Kiểm tra Giữa kỳ 1, Năm học 2025-2026, Môn Lịch sử 11, Thời gian 45 phút. Chồng dựa vào chuẩn kiến thức SGK Kết nối tri thức 2026 làm giúp vợ 2 bản đề gốc nhé!"
elif quick_btn3:
    prompt_input = "Vợ yêu xem đề gốc chồng làm rồi, ưng cái bụng lắm! Các câu hỏi chuẩn kiến thức 2026 và đúng ma trận. Không cần sửa gì thêm đâu chồng yêu ơi!"
elif quick_btn4:
    prompt_input = "Bây giờ chồng yêu trộn giúp vợ thành 4 mã đề (mã 1001, 1002, 1003, 1004) xáo trộn ngẫu nhiên logic và nhớ lập DUY NHẤT 01 BẢNG ĐÁP ÁN TỔNG HỢP dạng bảng kẻ cột cho tất cả các mã đề nhé!"

user_chat = st.chat_input("Nhắn gửi yêu cầu cho anh Chồng đẹp trai nào Vợ yêu ơi...")
if user_chat:
    prompt_input = user_chat

if prompt_input:
    if not api_key:
        st.warning("⚠️ Vui lòng cung cấp Gemini API Key ở thanh bên trái (Sidebar) để bắt đầu trò chuyện!")
    else:
        full_content = prompt_input
        if extracted_doc_text and "Ma trận" not in prompt_input and len(st.session_state.messages) <= 2:
            full_content += f"\n\n[ĐÍNH KÈM TÀI LIỆU MA TRẬN / ĐẶC TẢ]:\n{extracted_doc_text[:8000]}"

        st.session_state.messages.append({"role": "user", "content": full_content})
        with st.chat_message("user", avatar="👩‍🏫"):
            st.markdown(prompt_input)

        with st.chat_message("assistant", avatar="👨‍🏫"):
            with st.spinner("Anh chồng đang nghiên cứu ma trận và soạn đề theo sách Kết nối tri thức 2026 cho vợ yêu đây..."):
                try:
                    client = genai.Client(api_key=api_key)

                    formatted_contents = []
                    for m in st.session_state.messages:
                        role_name = "model" if m["role"] == "assistant" else "user"
                        formatted_contents.append({
                            "role": role_name,
                            "parts": [{"text": m["content"]}]
                        })

                    response = client.models.generate_content(
                        model="gemini-3.8-flash",
                        contents=formatted_contents,
                        config={
                            "system_instruction": SYSTEM_INSTRUCTION,
                            "temperature": 0.7,
                        }
                    )

                    ai_reply = response.text or "Chồng yêu đã nhận được yêu cầu của vợ rồi ạ!"
                    st.markdown(ai_reply)

                    # Nút tải file Word cho câu trả lời mới nhất
                    word_bytes = export_to_docx(ai_reply)
                    st.download_button(
                        label="📄 Tải ngay Đề này ra Word (.docx - Times New Roman 13)",
                        data=word_bytes,
                        file_name="De_Lich_Su_THPT_PhamVanDong.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_latest_{len(st.session_state.messages)}"
                    )

                    st.session_state.messages.append({"role": "assistant", "content": ai_reply})

                except Exception as ex:
                    st.error(f"Ối vợ ơi, gặp lỗi khi kết nối Gemini API: {str(ex)}")
