# ==============================================================================
# ỨNG DỤNG STREAMLIT: TRỢ LÝ SOẠN ĐỀ LỊCH SỬ THPT (GDPT 2018)
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
# 1. CẤU HÌNH TRANG VÀ GIAO DIỆN STREAMLIT
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Trợ Lý Ra Đề Lịch Sử THPT 2026 - Thầy Giáo Sử Yêu Vợ",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------------------------
# 2. SYSTEM INSTRUCTION - VAI TRÒ CHỒNG YÊU & NÓC NHÀ
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
# 3. HÀM ĐỌC NỘI DUNG TỪ FILE PDF (PyMuPDF) VÀ DOCX (python-docx)
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
    """Đọc toàn bộ văn bản từ file DOCX bằng python-docx"""
    doc = Document(io.BytesIO(file_bytes))
    full_text = []
    # Đọc paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
    # Đọc tables (ma trận đặc tả thường nằm trong bảng)
    for table in doc.tables:
        full_text.append("\n[BẢNG DỮ LIỆU MA TRẬN / ĐẶC TẢ]:")
        for row in table.rows:
            row_cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            full_text.append(" | ".join(row_cells))
    return "\n".join(full_text).strip()

# ------------------------------------------------------------------------------
# 4. HÀM XUẤT CÂU TRẢ LỜI CỦA AI RA FILE WORD (.DOCX) - TIMES NEW ROMAN 13
# ------------------------------------------------------------------------------
def export_to_docx(content_text: str, title: str = "De_Kiem_Tra_Lich_Su_THPT") -> bytes:
    """
    Chuyển nội dung AI thành file .docx chuẩn:
    - Font: Times New Roman
    - Size: 13 pt
    - Header trường, môn, tổ bộ môn chuẩn
    """
    doc = Document()

    # Cấu hình lề trang A4 chuẩn (Trên: 2cm, Dưới: 2cm, Trái: 2.5cm, Phải: 1.5cm)
    for section in doc.sections:
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.59)

    # Đặt style mặc định cho toàn bộ tài liệu là Times New Roman 13pt
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(13)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    # Thêm Tiêu đề Header chuẩn trường theo quy định
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

    # Đường phân cách nhẹ
    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Xử lý nội dung AI: duyệt từng dòng
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

        # Định dạng dòng tiêu đề hoặc câu hỏi
        if stripped.startswith("#"):
            clean_text = stripped.lstrip("#").strip()
            run = p.add_run(clean_text)
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
            run.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif stripped.startswith(("Câu ", "CÂU ", "Phần ", "PHẦN ", "MÃ ĐỀ")):
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
            # Xử lý các đáp án A. B. C. D. có bôi đậm
            run = p.add_run(line)
            run.font.name = "Times New Roman"
            run.font.size = Pt(13)
            if in_code_block:
                # Giữ nguyên font Times New Roman 13 chuẩn Word
                run.font.name = "Times New Roman"
                run.font.size = Pt(13)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ------------------------------------------------------------------------------
# 5. KHỞI TẠO CLIENT GEMINI TỪ API KEY
# ------------------------------------------------------------------------------
api_key = os.environ.get("GEMINI_API_KEY", "")

with st.sidebar:
    st.image("https://api.iconify.design/heroicons:academic-cap-20-solid.svg?color=%231e40af", width=50)
    st.title("Chồng Yêu Soạn Sử THPT")
    st.markdown("**Trợ lý chuẩn GDPT 2018 - Kết nối tri thức 2026**")
    st.markdown("---")

    user_api_key = st.text_input("Nhập Google Gemini API Key:", value=api_key, type="password")
    if user_api_key:
        api_key = user_api_key

    st.markdown("### 📂 Tải lên Ma trận / Bản đặc tả")
    uploaded_file = st.file_uploader("Chọn file .pdf hoặc .docx", type=["pdf", "docx"])

    extracted_doc_text = ""
    if uploaded_file is not None:
        try:
            file_bytes = uploaded_file.read()
            if uploaded_file.name.endswith(".pdf"):
                extracted_doc_text = extract_text_from_pdf(file_bytes)
                st.success(f"Đã đọc file PDF thành công ({len(extracted_doc_text)} ký tự)!")
            elif uploaded_file.name.endswith(".docx"):
                extracted_doc_text = extract_text_from_docx(file_bytes)
                st.success(f"Đã đọc file DOCX thành công ({len(extracted_doc_text)} ký tự)!")

            with st.expander("👁️ Xem trước nội dung file đã đọc"):
                st.text_area("Nội dung file trích xuất:", value=extracted_doc_text[:2000] + ("..." if len(extracted_doc_text) > 2000 else ""), height=200)
        except Exception as e:
            st.error(f"Lỗi khi đọc file: {str(e)}")

    st.markdown("---")
    st.markdown("### 💡 Gợi ý lệnh nhanh cho Vợ yêu:")
    quick_step1 = st.button("Bước 1: Chào chồng & Nhờ làm đề")
    quick_step2 = st.button("Bước 2: Gửi kèm Ma trận vừa tải lên")
    quick_step4 = st.button("Bước 4: Đồng ý đề gốc, nhờ trộn 4 mã đề")

# ------------------------------------------------------------------------------
# 6. QUẢN LÝ LỊCH SỬ CHAT STREAMLIT
# ------------------------------------------------------------------------------
if "messages" not in st.session_state:
    st.session_state.messages = []

# Header chính
st.subheader("❤️ Góc làm việc cùng Anh Chồng Lịch Sử THPT Phạm Văn Đồng")
st.caption("Chuyên gia GDPT 2018 bộ Kết nối tri thức 2026 - Ra đề thi trắc nghiệm chuẩn ma trận")

# Hiển thị lịch sử tin nhắn
for i, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"], avatar="👨‍🏫" if msg["role"] == "assistant" else "👩‍🏫"):
        st.markdown(msg["content"])
        
        # Nếu là câu trả lời của AI, cung cấp nút Xuất Word ngay dưới câu trả lời!
        if msg["role"] == "assistant":
            docx_data = export_to_docx(msg["content"], title=f"De_Lich_Su_Cau_{i}")
            st.download_button(
                label=f"📄 Xuất câu trả lời này ra Word (.docx - Times New Roman 13)",
                data=docx_data,
                file_name=f"De_Lich_Su_THPT_PVD_{i+1}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"dl_btn_{i}"
            )

# Xử lý các nút bấm nhanh từ Sidebar
prompt_input = None
if quick_step1:
    prompt_input = "Anh chồng ơi, vợ yêu đây! Hôm nay vợ cần anh giúp soạn đề kiểm tra Lịch sử THPT nhé."
elif quick_step2:
    if extracted_doc_text:
        prompt_input = f"Chồng yêu ơi, đây là nội dung Ma trận và Bản đặc tả từ file vợ vừa gửi lên:\n\n{extracted_doc_text}\n\nChồng xem kỹ rồi làm cho vợ 2 phiên bản đề gốc (File 1 in cho học sinh, File 2 đề & giải chi tiết) nhé!"
    else:
        prompt_input = "Chồng yêu ơi, hãy hướng dẫn vợ gửi Ma trận và Bản đặc tả nhé!"
elif quick_step4:
    prompt_input = "Vợ duyệt đề gốc rồi, ưng cái bụng lắm! Giờ chồng yêu trộn giúp vợ thành 4 mã đề kèm theo duy nhất 01 bảng đáp án tổng hợp nhé!"

# Nhận tin nhắn chat từ người dùng
user_text = st.chat_input("Nhắn gì cho anh Chồng đẹp trai nào Vợ yêu...")
if user_text:
    prompt_input = user_text

if prompt_input:
    if not api_key:
        st.warning("Vui lòng cung cấp Gemini API Key ở thanh bên trái (Sidebar) để bắt đầu trò chuyện!")
    else:
        # Nếu có file vừa upload chưa được kèm vào, hỏi hoặc tự động kèm
        full_user_content = prompt_input
        if extracted_doc_text and "Ma trận" not in prompt_input and len(st.session_state.messages) <= 2:
            full_user_content += f"\n\n[ĐÍNH KÈM NỘI DUNG TÀI LIỆU MA TRẬN/ĐẶC TẢ]:\n{extracted_doc_text[:8000]}"

        # Lưu tin nhắn người dùng
        st.session_state.messages.append({"role": "user", "content": full_user_content})
        with st.chat_message("user", avatar="👩‍🏫"):
            st.markdown(prompt_input)

        # Gọi Gemini API (@google/genai chuẩn SDK)
        with st.chat_message("assistant", avatar="👨‍🏫"):
            with st.spinner("Anh chồng đang nghiên cứu ma trận và soạn đề theo sách Kết nối tri thức 2026 cho vợ yêu đây..."):
                try:
                    client = genai.Client(api_key=api_key)

                    # Chuẩn bị lịch sử trò chuyện
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

                    ai_reply = response.text or "Chồng yêu đã nhận được yêu cầu nhưng chưa kịp sinh câu trả lời. Vợ yêu thử gửi lại xem sao nhé!"
                    st.markdown(ai_reply)

                    # Nút tải file Word ngay lập tức
                    word_bytes = export_to_docx(ai_reply)
                    st.download_button(
                        label="📄 Tải ngay Đề kiểm tra ra Word (.docx - Times New Roman 13)",
                        data=word_bytes,
                        file_name="De_Lich_Su_THPT_PhamVanDong.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_latest_{len(st.session_state.messages)}"
                    )

                    st.session_state.messages.append({"role": "assistant", "content": ai_reply})

                except Exception as ex:
                    st.error(f"Anh chồng gặp chút trục trặc khi kết nối Gemini API: {str(ex)}")
