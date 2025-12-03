"""
PDF Table Extraction Web App
PDF → 페이지 분할 → 이미지 → GPT 테이블 추출 → Excel 저장

[주요 기능]
1. 2단계 처리 프로세스:
   - 1단계: GPT를 사용하여 테이블이 있는 페이지 자동 탐지
   - 2단계: 사용자가 원하는 페이지만 선택하여 테이블 추출
2. 선택적 테이블 추출: 체크박스로 원하는 페이지만 선택 가능

[성능 개선 사항]
1. 병렬처리: ThreadPoolExecutor를 사용하여 GPT API 호출 병렬화 (최대 5개 동시 처리)
   - detect_table_pages(): 테이블 존재 여부 확인 병렬 처리
   - process_jpgs_to_excel(): 테이블 추출 병렬 처리
2. DPI 최적화: DPI 150으로 설정하여 속도와 품질 균형
3. 세션 상태 최소화: 필수 항목만 세션에 저장하여 메모리 오버헤드 감소
"""

import streamlit as st
import os
import json
import base64
import pandas as pd
import tempfile
import shutil
import platform
import subprocess
from io import BytesIO
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

from pdf2image import convert_from_path
from openai import OpenAI
from PyPDF2 import PdfReader, PdfWriter


# ======================== 유틸리티 함수 ========================

def get_poppler_path():
    """Windows에서 poppler 경로를 찾는 함수"""
    if platform.system() != "Windows":
        return None
    
    # 1. PATH 환경변수에서 pdftoppm.exe 찾기
    pdftoppm_path = shutil.which("pdftoppm")
    if pdftoppm_path:
        bin_dir = os.path.dirname(pdftoppm_path)
        if os.path.exists(bin_dir):
            return bin_dir
    
    # 2. 환경변수에서 poppler 경로 확인
    if "POPPLER_PATH" in os.environ:
        env_path = os.environ["POPPLER_PATH"]
        if os.path.isdir(env_path):
            bin_path = os.path.join(env_path, "bin")
            if os.path.exists(bin_path):
                return bin_path
            if os.path.exists(os.path.join(env_path, "pdftoppm.exe")):
                return env_path
    
    # 3. 일반적인 poppler 설치 경로들
    possible_paths = [
        r"C:\poppler\bin",
        r"C:\poppler-24.08.0\Library\bin",
        r"C:\poppler-24.06.0\Library\bin",
        r"C:\poppler-24.02.0\Library\bin",
        r"C:\poppler-23.11.0\Library\bin",
        r"C:\poppler-23.10.0\Library\bin",
        r"C:\poppler-23.08.0\Library\bin",
        r"C:\Program Files\poppler\bin",
        r"C:\Program Files (x86)\poppler\bin",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "poppler", "bin"),
        os.path.join(os.environ.get("PROGRAMFILES", ""), "poppler", "bin"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)", ""), "poppler", "bin"),
    ]
    
    # 4. C 드라이브에서 poppler 폴더 검색
    if os.path.exists("C:\\"):
        try:
            for item in os.listdir("C:\\"):
                poppler_dir = os.path.join("C:\\", item)
                if os.path.isdir(poppler_dir) and "poppler" in item.lower():
                    bin_path = os.path.join(poppler_dir, "bin")
                    if os.path.exists(bin_path):
                        possible_paths.append(bin_path)
                    lib_bin_path = os.path.join(poppler_dir, "Library", "bin")
                    if os.path.exists(lib_bin_path):
                        possible_paths.append(lib_bin_path)
        except:
            pass
    
    # 5. 가능한 경로들 확인
    for path in possible_paths:
        if os.path.exists(path):
            pdftoppm_exe = os.path.join(path, "pdftoppm.exe")
            if os.path.exists(pdftoppm_exe):
                return path
    
    return None


# ======================== PDF 처리 함수 ========================

def split_pdf(input_pdf_path, output_dir, chunk_size=15):
    """PDF를 청크 단위로 분할"""
    os.makedirs(output_dir, exist_ok=True)
    
    reader = PdfReader(input_pdf_path)
    total_pages = len(reader.pages)
    outputs = []
    
    for i in range(0, total_pages, chunk_size):
        writer = PdfWriter()
        start = i
        end = min(i + chunk_size, total_pages)
        
        for p in range(start, end):
            writer.add_page(reader.pages[p])
        
        out_name = os.path.join(output_dir, f"chunk_{i // chunk_size + 1}.pdf")
        
        with open(out_name, "wb") as f:
            writer.write(f)
        
        outputs.append(out_name)
    
    return outputs


# ======================== GPT 분석 함수 ========================

def analyze_image_for_table(client, img_path):
    """이미지에 테이블이 있는지 확인"""
    with open(img_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    
    resp = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": "Is there a table in this image? Answer yes or no."},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}
                ]
            }
        ]
    )
    
    return resp.choices[0].message.content.strip().lower()


def detect_table_pages(client, pdf_path, poppler_path, status_placeholder):
    """PDF에서 테이블이 있는 페이지 탐지 (병렬처리)"""
    kwargs = {"dpi": 150}
    if poppler_path:
        kwargs["poppler_path"] = poppler_path
    
    pages = convert_from_path(pdf_path, **kwargs)
    detected_pages = []
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        # 모든 페이지를 먼저 저장
        img_paths = []
        for idx, img in enumerate(pages, start=1):
            img_path = os.path.join(temp_dir, f"tmp_page_{idx}.jpg")
            img.save(img_path, "JPEG")
            img_paths.append((idx, img_path))
        
        # 병렬처리로 GPT API 호출 (최대 5개 동시 처리)
        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_page = {
                executor.submit(analyze_image_for_table, client, img_path): idx 
                for idx, img_path in img_paths
            }
            
            for future in as_completed(future_to_page):
                page_num = future_to_page[future]
                try:
                    result = future.result()
                    if "yes" in result.lower():
                        detected_pages.append(page_num)
                    
                    # 진행 상황 업데이트
                    completed = len([f for f in future_to_page if f.done()])
                    status_placeholder.info(f"💬 Asking GPT... ({completed}/{len(pages)} pages analyzed)")
                except Exception as e:
                    print(f"Error analyzing page {page_num}: {e}")
    
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    return sorted(detected_pages)


def save_table_pages_as_jpg(pdf_path, table_pages, output_dir, poppler_path):
    """테이블 페이지를 JPG로 저장"""
    os.makedirs(output_dir, exist_ok=True)
    
    kwargs = {"dpi": 150}
    if poppler_path:
        kwargs["poppler_path"] = poppler_path
    
    pages = convert_from_path(pdf_path, **kwargs)
    base = Path(pdf_path).stem
    saved = []
    
    for page_num in table_pages:
        img = pages[page_num - 1]
        jpg_path = os.path.join(output_dir, f"{base}_page_{page_num}.jpg")
        img.save(jpg_path, "JPEG")
        saved.append(jpg_path)
    
    return saved


def extract_tables_from_image(client, image_path):
    """이미지에서 테이블 데이터 추출"""
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    
    prompt_message = (
        "You MUST extract three things from this image:\n"
        "1) The title text located ABOVE the table (even if it is not inside the table box).\n"
        "2) The header row of the table (the first row that describes columns).\n"
        "3) The table body.\n\n"
        "Return ONLY pure JSON in this structure:\n"
        "{\n"
        "  \"tables\": [\n"
        "    {\n"
        "      \"title\": \"...\",\n"
        "      \"header\": [\"...\", \"...\", ...],\n"
        "      \"data\": [[...], [...]]\n"
        "    }\n"
        "  ]\n"
        "}\n"
        "If the table has no explicit header row, leave 'header' as an empty list.\n"
        "If multiple lines of text exist above the table, combine them into a single title string.\n"
        "The JSON MUST NOT be inside markdown code fences. Return only raw JSON."
    )
    
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt_message},
                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}
                ]
            }
        ]
    )
    
    raw = response.choices[0].message.content.strip()
    
    # 코드블록 제거
    if raw.startswith("```"):
        raw = raw.replace("```json", "").replace("```", "").strip()
    
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {"tables": []}


def process_jpgs_to_excel(client, jpg_folder, status_placeholder):
    """JPG 이미지들을 Excel로 변환 (병렬처리)"""
    jpg_files = sorted(
        [os.path.join(jpg_folder, f) for f in os.listdir(jpg_folder) if f.lower().endswith(".jpg")]
    )
    
    if not jpg_files:
        return None
    
    total_files = len(jpg_files)
    
    # 병렬처리로 GPT API 호출 (최대 5개 동시 처리)
    all_results = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_img = {
            executor.submit(extract_tables_from_image, client, img_path): img_path 
            for img_path in jpg_files
        }
        
        for future in as_completed(future_to_img):
            img_path = future_to_img[future]
            try:
                tables = future.result().get("tables", [])
                if tables:
                    all_results.append((img_path, tables))
                
                # 진행 상황 업데이트
                completed = len([f for f in future_to_img if f.done()])
                status_placeholder.info(f"💬 Extracting tables... ({completed}/{total_files} images processed)")
            except Exception as e:
                print(f"Error extracting from {Path(img_path).name}: {e}")
    
    # BytesIO를 사용하여 메모리에 Excel 파일 생성
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine="openpyxl")
    
    # 결과를 파일명 순서대로 정렬하여 Excel에 작성
    all_results.sort(key=lambda x: x[0])
    
    for img_path, tables in all_results:
        base = Path(img_path).stem
        
        for idx, t in enumerate(tables, start=1):
            title = t.get("title", "")
            header = t.get("header", [])
            data = t.get("data", [])
            
            final_rows = []
            
            if title:
                final_rows.append([title] + [""] * (max(len(header) - 1, 0)))
            
            if header:
                final_rows.append(header)
            
            final_rows.extend(data)
            
            df = pd.DataFrame(final_rows)
            sheet_name = f"{base}_T{idx}"[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    
    writer.close()
    output.seek(0)
    
    return output


# ======================== Streamlit UI ========================

def main():
    st.set_page_config(
        page_title="PDF Table Extractor",
        page_icon="📊",
        layout="wide"
    )
    
    # 세션 상태 초기화 (필수 항목만 유지)
    if "processed" not in st.session_state:
        st.session_state.processed = False
    if "excel_data" not in st.session_state:
        st.session_state.excel_data = None
    if "save_dir" not in st.session_state:
        st.session_state.save_dir = None
    if "detection_complete" not in st.session_state:
        st.session_state.detection_complete = False
    if "selected_images" not in st.session_state:
        st.session_state.selected_images = []
    
    # 타이틀
    st.title("📊 PDF Table Extractor")
    st.markdown("**PDF → 페이지 분할 → 이미지 → GPT 테이블 추출 → Excel 저장**")
    st.markdown("*사용 모델: GPT-4o-Mini*")
    
    st.divider()
    
    # ======================== 사이드바 ========================
    with st.sidebar:
        st.header("⚙️ 설정")
        
        # OpenAI API Key 입력
        api_key = st.text_input(
            "OpenAI API Key",
            type="password",
            help="개인 OpenAI API 키를 입력하세요"
        )
        
        st.divider()
        
        # Poppler 경로 입력 (Windows 전용)
        manual_poppler_path = None
        if platform.system() == "Windows":
            # 자동 감지 시도
            auto_detected = get_poppler_path()
            
            if auto_detected:
                st.success(f"✅ Poppler 자동 감지됨")
                st.code(auto_detected, language=None)
            else:
                st.warning("⚠️ Poppler 자동 감지 실패")
            
            # 수동 입력 옵션
            with st.expander("🔧 Poppler 경로 수동 입력"):
                manual_poppler_path = st.text_input(
                    "Poppler bin 폴더 경로",
                    value=r"C:\poppler\poppler-23.11.0\Library\bin",
                    help="Poppler의 bin 폴더 경로를 입력하세요 (pdftoppm.exe가 있는 폴더)"
                )
                
                if manual_poppler_path and manual_poppler_path.strip():
                    pdftoppm_exe = os.path.join(manual_poppler_path, "pdftoppm.exe")
                    if os.path.exists(pdftoppm_exe):
                        st.success("✅ 올바른 경로입니다!")
                    else:
                        st.error("❌ pdftoppm.exe를 찾을 수 없습니다.")
            
            st.divider()
        
        # PDF 업로드
        uploaded_file = st.file_uploader(
            "📄 PDF 파일 업로드",
            type=["pdf"],
            help="테이블을 추출할 PDF 파일을 업로드하세요"
        )
        
        st.divider()
        
        # 처리 시작 버튼
        start_button = st.button("🚀 Start Processing", type="primary", use_container_width=True)
        
        st.divider()
        
        # Excel 다운로드 (처리 완료 후에만 표시)
        if st.session_state.processed and st.session_state.excel_data:
            st.download_button(
                label="📥 Download Excel",
                data=st.session_state.excel_data,
                file_name="extracted_tables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # ======================== 메인 페이지 ========================
    
    # 입력 검증
    if start_button:
        if not api_key:
            st.error("❌ OpenAI API Key를 입력해주세요!")
            return
        
        if not uploaded_file:
            st.error("❌ PDF 파일을 업로드해주세요!")
            return
        
        # 초기화
        st.session_state.processed = False
        st.session_state.excel_data = None
        st.session_state.save_dir = None
        st.session_state.detection_complete = False
        st.session_state.selected_images = []
        
        # Poppler 경로 확인
        # 수동 입력된 경로가 있으면 우선 사용
        if manual_poppler_path and manual_poppler_path.strip():
            poppler_path = manual_poppler_path.strip()
        else:
            poppler_path = get_poppler_path()
        
        if platform.system() == "Windows" and not poppler_path:
            st.error(
                "❌ Poppler를 찾을 수 없습니다!\n\n"
                "**해결 방법:**\n"
                "1. 사이드바에서 '🔧 Poppler 경로 수동 입력'을 열어 경로를 입력하거나\n"
                "2. https://github.com/oschwartz10612/poppler-windows/releases 에서 다운로드\n"
                "3. 압축 해제 후 C:\\poppler 경로에 저장\n"
                "4. 환경변수 PATH에 C:\\poppler\\bin 추가"
            )
            return
        
        # pdftoppm.exe 존재 확인
        if platform.system() == "Windows" and poppler_path:
            pdftoppm_exe = os.path.join(poppler_path, "pdftoppm.exe")
            if not os.path.exists(pdftoppm_exe):
                st.error(
                    f"❌ 잘못된 Poppler 경로입니다!\n\n"
                    f"입력된 경로: `{poppler_path}`\n\n"
                    f"pdftoppm.exe 파일이 이 경로에 없습니다.\n"
                    f"올바른 bin 폴더 경로를 입력했는지 확인하세요."
                )
                return
        
        # API 키 세션 상태에 저장
        st.session_state.api_key = api_key
        
        # OpenAI 클라이언트 초기화
        try:
            client = OpenAI(api_key=api_key)
        except Exception as e:
            st.error(f"❌ OpenAI 클라이언트 초기화 실패: {e}")
            return
        
        # Progress bar와 status box
        progress_bar = st.progress(0)
        status_box = st.empty()
        
        # save 폴더 생성 (영구 저장)
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_dir = os.path.join("save", f"session_{timestamp}")
        os.makedirs(save_dir, exist_ok=True)
        st.session_state.save_dir = save_dir
        
        # 임시 디렉토리는 PDF와 chunk만 저장
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # Step 1: PDF 저장
                status_box.info("📄 Saving uploaded PDF...")
                progress_bar.progress(0.05)
                
                pdf_path = os.path.join(temp_dir, "uploaded.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                
                # Step 2: PDF 분할
                status_box.info("✂️ Splitting PDF into chunks...")
                progress_bar.progress(0.10)
                
                chunks_dir = os.path.join(save_dir, "chunks")
                chunks = split_pdf(pdf_path, chunks_dir, chunk_size=15)
                
                # Step 3: 테이블 페이지 탐지
                status_box.info("🔍 Detecting table pages with GPT...")
                progress_bar.progress(0.20)
                
                # 로컬 변수로 처리 (세션 상태 최소화)
                all_table_pages = {}
                all_jpg_files = []
                jpg_output_dir = os.path.join(save_dir, "PDF_single_page_jpg")
                
                total_chunks = len(chunks)
                
                for chunk_idx, chunk_path in enumerate(chunks, start=1):
                    # 청크별 진행률 계산 (20% ~ 60%)
                    chunk_progress = 0.20 + (0.40 * chunk_idx / total_chunks)
                    progress_bar.progress(chunk_progress)
                    
                    chunk_name = Path(chunk_path).name
                    status_box.info(f"🔍 Analyzing {chunk_name} ({chunk_idx}/{total_chunks})...")
                    
                    table_pages = detect_table_pages(client, chunk_path, poppler_path, status_box)
                    
                    if table_pages:
                        all_table_pages[chunk_name] = table_pages
                        
                        # JPG 저장 (save 폴더에 영구 저장)
                        status_box.info(f"🖼️ Saving table pages as JPG...")
                        saved_jpgs = save_table_pages_as_jpg(chunk_path, table_pages, jpg_output_dir, poppler_path)
                        all_jpg_files.extend(saved_jpgs)
                
                # 테이블 탐지 완료
                if all_jpg_files:
                    progress_bar.progress(0.70)
                    status_box.success(f"✅ Table detection completed! Found {len(all_jpg_files)} pages with tables.")
                    st.session_state.detection_complete = True
                    st.session_state.selected_images = all_jpg_files  # 기본적으로 모두 선택
                else:
                    progress_bar.progress(1.0)
                    status_box.warning("⚠️ No table pages were detected in the PDF.")
                    st.session_state.detection_complete = True
            
            except Exception as e:
                st.error(f"❌ 처리 중 오류 발생: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
    
    # ======================== 페이지 선택 UI ========================
    
    if st.session_state.detection_complete and st.session_state.save_dir and not st.session_state.processed:
        st.divider()
        save_dir = st.session_state.save_dir
        
        st.subheader("📋 테이블 추출할 페이지 선택")
        st.markdown("추출하고 싶은 페이지를 선택하세요. 선택된 페이지만 테이블 추출 작업이 진행됩니다.")
        
        # JPG 파일 목록 가져오기
        jpg_output_dir = os.path.join(save_dir, "PDF_single_page_jpg")
        if os.path.exists(jpg_output_dir):
            jpg_files = sorted([
                os.path.join(jpg_output_dir, f) 
                for f in os.listdir(jpg_output_dir) 
                if f.lower().endswith('.jpg')
            ])
            
            if jpg_files:
                # 전체 선택/해제 버튼
                col1, col2, col3 = st.columns([1, 1, 4])
                with col1:
                    if st.button("✅ 전체 선택", use_container_width=True):
                        st.session_state.selected_images = jpg_files.copy()
                        st.rerun()
                with col2:
                    if st.button("❌ 전체 해제", use_container_width=True):
                        st.session_state.selected_images = []
                        st.rerun()
                
                st.markdown(f"**선택된 페이지: {len(st.session_state.selected_images)}/{len(jpg_files)}**")
                st.divider()
                
                # 이미지 그리드 표시 (체크박스 포함)
                cols = st.columns(3)
                
                for idx, jpg_path in enumerate(jpg_files):
                    with cols[idx % 3]:
                        # 이미지 표시
                        st.image(jpg_path, caption=Path(jpg_path).name, use_container_width=True)
                        
                        # 체크박스
                        is_selected = jpg_path in st.session_state.selected_images
                        if st.checkbox(
                            f"선택", 
                            value=is_selected, 
                            key=f"checkbox_{idx}_{Path(jpg_path).name}"
                        ):
                            if jpg_path not in st.session_state.selected_images:
                                st.session_state.selected_images.append(jpg_path)
                        else:
                            if jpg_path in st.session_state.selected_images:
                                st.session_state.selected_images.remove(jpg_path)
                
                st.divider()
                
                # 테이블 추출 시작 버튼
                if st.session_state.selected_images:
                    if st.button(
                        f"🚀 선택한 {len(st.session_state.selected_images)}개 페이지에서 테이블 추출", 
                        type="primary", 
                        use_container_width=True
                    ):
                        # Progress bar와 status box
                        progress_bar = st.progress(0)
                        status_box = st.empty()
                        
                        try:
                            # OpenAI 클라이언트 재사용
                            api_key = st.session_state.get("api_key")
                            if not api_key:
                                st.error("❌ API 키가 없습니다. 페이지를 새로고침하고 다시 시도하세요.")
                                st.stop()
                            
                            client = OpenAI(api_key=api_key)
                            
                            # 선택된 이미지만 처리
                            status_box.info("📊 Extracting tables from selected pages...")
                            progress_bar.progress(0.1)
                            
                            # 선택된 이미지를 임시 폴더에 복사
                            temp_selected_dir = os.path.join(save_dir, "selected_pages")
                            os.makedirs(temp_selected_dir, exist_ok=True)
                            
                            for img_path in st.session_state.selected_images:
                                shutil.copy(img_path, temp_selected_dir)
                            
                            progress_bar.progress(0.2)
                            
                            # 선택된 페이지에서 테이블 추출
                            excel_data = process_jpgs_to_excel(client, temp_selected_dir, status_box)
                            
                            if excel_data:
                                st.session_state.excel_data = excel_data.getvalue()
                                
                                # Excel 파일 저장
                                excel_path = os.path.join(save_dir, "extracted_tables.xlsx")
                                with open(excel_path, "wb") as f:
                                    f.write(st.session_state.excel_data)
                                
                                progress_bar.progress(1.0)
                                status_box.success(f"✅ Table extraction completed! {len(st.session_state.selected_images)} pages processed.")
                                st.session_state.processed = True
                                st.rerun()
                            else:
                                progress_bar.progress(1.0)
                                status_box.warning("⚠️ No tables were extracted from the selected images.")
                        
                        except Exception as e:
                            st.error(f"❌ 테이블 추출 중 오류 발생: {str(e)}")
                            import traceback
                            st.code(traceback.format_exc())
                else:
                    st.warning("⚠️ 최소 1개 이상의 페이지를 선택해주세요.")
    
    # ======================== 결과 표시 ========================
    
    if st.session_state.processed and st.session_state.save_dir:
        st.divider()
        save_dir = st.session_state.save_dir
        
        # 선택된 이미지 표시
        st.subheader("🖼️ 추출된 페이지")
        st.markdown(f"**총 {len(st.session_state.selected_images)}개 페이지**")
        
        # 3열 그리드로 표시
        cols = st.columns(3)
        
        for idx, jpg_path in enumerate(st.session_state.selected_images):
            with cols[idx % 3]:
                st.image(jpg_path, caption=Path(jpg_path).name, use_container_width=True)
        
        # Excel 다운로드 버튼 (메인 페이지)
        if st.session_state.excel_data:
            st.divider()
            st.subheader("📥 결과 다운로드")
            
            # 다운로드 버튼 (클라우드 환경 고려)
            is_cloud = os.path.exists("/mount/src")  # Streamlit Cloud 감지
            
            if is_cloud:
                # 클라우드 환경: 다운로드 버튼만 표시
                st.download_button(
                    label="📥 Download Excel File",
                    data=st.session_state.excel_data,
                    file_name="extracted_tables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            else:
                # 로컬 환경: 다운로드 + 폴더 열기 버튼
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="📥 Download Excel File",
                        data=st.session_state.excel_data,
                        file_name="extracted_tables.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                
                with col2:
                    if st.session_state.save_dir and os.path.exists(st.session_state.save_dir):
                        if st.button("📂 Open Folder", type="secondary", use_container_width=True):
                            try:
                                if platform.system() == "Windows":
                                    os.startfile(st.session_state.save_dir)
                                elif platform.system() == "Darwin":  # macOS
                                    subprocess.Popen(["open", st.session_state.save_dir])
                                else:  # Linux
                                    subprocess.Popen(["xdg-open", st.session_state.save_dir])
                                st.success(f"✅ 폴더를 열었습니다!")
                            except Exception as e:
                                st.error(f"❌ 폴더를 열 수 없습니다: {e}")
            
            # 저장 위치 정보
            if st.session_state.save_dir:
                st.info(f"💾 **저장 위치**: `{st.session_state.save_dir}`")
                
                with st.expander("📁 저장된 파일 목록 보기"):
                    # PDF 청크
                    chunks_dir = os.path.join(st.session_state.save_dir, "chunks")
                    if os.path.exists(chunks_dir):
                        chunks = [f for f in os.listdir(chunks_dir) if f.endswith('.pdf')]
                        st.markdown(f"**PDF 청크**: {len(chunks)}개")
                        for chunk in sorted(chunks):
                            st.text(f"  - {chunk}")
                    
                    # JPG 이미지
                    jpg_dir = os.path.join(st.session_state.save_dir, "PDF_single_page_jpg")
                    if os.path.exists(jpg_dir):
                        jpgs = [f for f in os.listdir(jpg_dir) if f.endswith('.jpg')]
                        st.markdown(f"**JPG 이미지**: {len(jpgs)}개")
                        for jpg in sorted(jpgs):
                            st.text(f"  - {jpg}")
                    
                    # Excel 파일
                    excel_path = os.path.join(st.session_state.save_dir, "extracted_tables.xlsx")
                    if os.path.exists(excel_path):
                        st.markdown("**Excel 파일**: extracted_tables.xlsx")


if __name__ == "__main__":
    main()

