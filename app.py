"""
리움동물의료센터 심장 초음파 검진 소견서 자동 생성 웹 애플리케이션
"""
import os
import streamlit as st
from datetime import datetime
from report_generator import (
    HeartReportData,
    generate_heart_report_docx,
    convert_docx_to_pdf,
    evaluate_metrics,
    get_default_owner_message,
)

# 페이지 기본 설정
st.set_page_config(
    page_title="리움동물의료센터 심장 초음파 검진 소견서 생성기",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 스타일링
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 800;
        color: #1a365d;
        margin-bottom: 2px;
    }
    .sub-header {
        font-size: 14px;
        color: #4a5568;
        margin-bottom: 20px;
    }
    .card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
    }
    .card-title {
        font-size: 16px;
        font-weight: 700;
        color: #2b6cb0;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .badge-normal {
        background-color: #e6fffa;
        color: #234e52;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 12px;
        border: 1px solid #b2f5ea;
    }
    .badge-abnormal {
        background-color: #fff5f5;
        color: #9b2c2c;
        padding: 3px 8px;
        border-radius: 4px;
        font-weight: 700;
        font-size: 12px;
        border: 1px solid #feb2b2;
    }
    .badge-empty {
        background-color: #edf2f7;
        color: #718096;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 12px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">🩺 리움동물의료센터 심장 초음파 검진 소견서 자동 생성기</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">환자 정보와 수치를 입력하면 정상/비정상 자동 판정 및 템플릿 기반 Word/PDF 소견서를 자동 생성합니다.</div>', unsafe_allow_html=True)

# 2단 레이아웃 분할
col_left, col_right = st.columns([1.1, 0.9], gap="large")

with col_left:
    # 1. 환자 정보
    with st.container():
        st.markdown('<div class="card"><div class="card-title">📋 1. 환자 기본 정보</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            patient_name = st.text_input("환자명 *", value="별이", placeholder="예: 별이")
            breed = st.text_input("품종 *", value="말티즈", placeholder="예: 말티즈, 포메라니안")
            age = st.text_input("나이 *", value="11년령", placeholder="예: 11년 4개월")
        with c2:
            gender = st.text_input("성별 *", value="중성화 암컷 (SF)", placeholder="예: 중성화 수컷 / SF")
            owner_name = st.text_input("보호자 성함 *", value="홍길동", placeholder="예: 홍길동")
        st.markdown('</div>', unsafe_allow_html=True)

    # 2. 신체검사 및 흉부 방사선
    with st.container():
        st.markdown('<div class="card"><div class="card-title">🩺 2. 신체검사 및 흉부 방사선 (X-ray)</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            murmur_type = st.radio("청진 (심잡음) [정상 기준: 심잡음 없음]", ["심잡음 없음", "심잡음 청진됨", "직접 입력"], horizontal=True)
            if murmur_type == "심잡음 청진됨":
                grade = st.selectbox("심잡음 등급 (Grade)", ["Grade I/VI", "Grade II/VI", "Grade III/VI", "Grade IV/VI", "Grade V/VI", "Grade VI/VI"], index=2)
                murmur = f"심잡음 청진됨 ({grade})"
            elif murmur_type == "직접 입력":
                murmur = st.text_input("청진 상세 소견", value="심잡음 청진됨 (좌측 심첨부 수축기 잡음)")
            else:
                murmur = "심잡음 없음"

            heart_rate = st.text_input("심박수 (회/분)", value="138", placeholder="예: 130")
        with c2:
            vhs = st.text_input("X-ray (VHS)", value="10.8", placeholder="정상 < 10.5 (개)")
            vlas = st.text_input("X-ray (VLAS)", value="2.6", placeholder="정상 < 2.5")
        st.markdown('</div>', unsafe_allow_html=True)

    # 3. 심장 초음파 지표 입력
    with st.container():
        st.markdown('<div class="card"><div class="card-title">💓 3. 심장 초음파 결과 수치 (자동 판정)</div>', unsafe_allow_html=True)
        st.caption("※ 기준값과 같거나 큰 경우 정상 범위 초과(비정상)로 자동 판정됩니다.")
        c1, c2 = st.columns(2)
        with c1:
            la_ao = st.text_input("LA:Ao (좌심방:대동맥 비율) [정상 < 1.6]", value="1.85", placeholder="예: 1.85")
            lviddn = st.text_input("LVIDDn (이완기 좌심실 직경 지수) [정상 < 1.7]", value="1.72", placeholder="예: 1.72")
        with c2:
            e_peak = st.text_input("E peak (좌심실 혈류속도, cm/s) [정상 < 60]", value="75.4", placeholder="예: 75.4")
            ee_ratio = st.text_input("E/E' (혈류속도/이완속도) [정상 < 10]", value="11.5", placeholder="예: 11.5")
        st.markdown('</div>', unsafe_allow_html=True)

    # 4. ACVIM 병기 선택 및 보호자 안내문
    with st.container():
        st.markdown('<div class="card"><div class="card-title">🏷️ 4. ACVIM 병기 및 맞춤 보호자 안내</div>', unsafe_allow_html=True)
        stage = st.selectbox(
            "ACVIM 병기 선택 (수의사 직접 지정) *",
            options=["A", "B1", "B2", "C", "D"],
            index=2, # 기본 B2
            format_func=lambda x: {
                "A": "Stage A (위험군 / 심장 질환 없음)",
                "B1": "Stage B1 (심잡음 O, 심장 비대 X, 모니터링)",
                "B2": "Stage B2 (심잡음 O, 심장 비대 O, 약물 치료 권장)",
                "C": "Stage C (과거/현재 심부전 증상 동반, 적극 치료)",
                "D": "Stage D (난치성 심부전, 집중 치료)"
            }[x]
        )

        diagnosis = st.text_input(
            "2. 진단명 표기",
            value=f"■ 승모판 점액종성 판막 질환 (MMVD) Stage {stage}"
        )

        # 기본 안내문 추천
        default_msg = get_default_owner_message(patient_name, stage)
        owner_message = st.text_area(
            "7. 보호자님께 드리는 말씀 (병기 연동 / 직접 편집 가능)",
            value=default_msg,
            height=120
        )
        st.markdown('</div>', unsafe_allow_html=True)

    # 5. 초음파 사진 업로드
    with st.container():
        st.markdown('<div class="card"><div class="card-title">🖼️ 5. 초음파 사진 업로드 (2장)</div>', unsafe_allow_html=True)
        st.caption("기존 템플릿의 사진 위치에 비율 유지(찌그러짐 없이) 삽입됩니다. (미업로드 시 기본 견본 유지)")
        c1, c2 = st.columns(2)
        with c1:
            img_file1 = st.file_uploader("초음파 사진 1 (좌측)", type=["jpg", "jpeg", "png"])
            if img_file1:
                st.image(img_file1, caption="사진 1 미리보기", use_container_width=True)
        with c2:
            img_file2 = st.file_uploader("초음파 사진 2 (우측)", type=["jpg", "jpeg", "png"])
            if img_file2:
                st.image(img_file2, caption="사진 2 미리보기", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

# 데이터 인스턴스 구성
report_data = HeartReportData(
    patient_name=patient_name,
    breed=breed,
    age=age,
    gender=gender,
    owner_name=owner_name,
    murmur=murmur,
    heart_rate=heart_rate,
    vhs=vhs,
    vlas=vlas,
    la_ao=la_ao,
    lviddn=lviddn,
    e_peak=e_peak,
    ee_ratio=ee_ratio,
    stage=stage,
    diagnosis=diagnosis,
    owner_message=owner_message,
    image1_bytes=img_file1.getvalue() if img_file1 else None,
    image2_bytes=img_file2.getvalue() if img_file2 else None,
)

eval_res = evaluate_metrics(report_data)

# 우측 패널: 실시간 미리보기 및 소견서 생성/다운로드
with col_right:
    st.markdown('<div class="card"><div class="card-title">🔍 실시간 판정 및 소견서 미리보기</div>', unsafe_allow_html=True)

    # 환자 요약
    st.markdown(f"**환자:** `{patient_name or '미입력'}` ({breed or '-'} / {age or '-'} / {gender or '-'}) | **보호자:** `{owner_name or '미입력'}`")
    st.markdown(f"**청진:** `{murmur}` | **심박수:** `{heart_rate or '-'}회/분` | **X-ray:** `VHS {vhs or '-'} / VLAS {vlas or '-'}`")
    st.markdown(f"**진단명:** `{diagnosis}` | **최종 병기:** :red[**Stage {stage}**]")

    st.divider()

    # 초음파 판정 결과 표
    st.markdown("##### 📊 심장초음파 지표 자동 판정표")
    
    def render_badge(status_info):
        st_type = status_info["status"]
        remark = status_info["remark"]
        if st_type == "normal":
            return f'<span class="badge-normal">✅ 정상</span>'
        elif st_type == "abnormal":
            return f'<span class="badge-abnormal">⚠️ {remark}</span>'
        else:
            return f'<span class="badge-empty">- (미입력)</span>'

    table_html = f"""
    <table style="width:100%; border-collapse: collapse; font-size: 13px; text-align: left; margin-bottom: 16px;">
        <thead>
            <tr style="background-color: #f1f5f9; border-bottom: 2px solid #cbd5e1;">
                <th style="padding: 8px;">검사항목</th>
                <th style="padding: 8px;">정상 기준</th>
                <th style="padding: 8px;">입력값</th>
                <th style="padding: 8px;">자동 판정 결과 (비고)</th>
            </tr>
        </thead>
        <tbody>
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px; font-weight:600;">LA:Ao (좌심방:대동맥)</td>
                <td style="padding: 8px;">&lt; 1.6</td>
                <td style="padding: 8px;">{la_ao or '-'}</td>
                <td style="padding: 8px;">{render_badge(eval_res['la_ao'])}</td>
            </tr>
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px; font-weight:600;">LVIDDn (좌심실 직경)</td>
                <td style="padding: 8px;">&lt; 1.7</td>
                <td style="padding: 8px;">{lviddn or '-'}</td>
                <td style="padding: 8px;">{render_badge(eval_res['lviddn'])}</td>
            </tr>
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px; font-weight:600;">E peak (혈류속도)</td>
                <td style="padding: 8px;">&lt; 60 cm/s</td>
                <td style="padding: 8px;">{e_peak or '-'} cm/s</td>
                <td style="padding: 8px;">{render_badge(eval_res['e_peak'])}</td>
            </tr>
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 8px; font-weight:600;">E/E' (충만압)</td>
                <td style="padding: 8px;">&lt; 10</td>
                <td style="padding: 8px;">{ee_ratio or '-'}</td>
                <td style="padding: 8px;">{render_badge(eval_res['ee_ratio'])}</td>
            </tr>
        </tbody>
    </table>
    """
    st.markdown(table_html, unsafe_allow_html=True)

    st.markdown("##### 💬 보호자 전달 문구 미리보기")
    st.info(owner_message)

    st.markdown('</div>', unsafe_allow_html=True)

    # 소견서 문서 생성 액션
    st.markdown("### 🚀 소견서 출력 및 다운로드")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_name = patient_name.strip() or "환자"
    docx_filename = f"26년_{clean_name}_심장건강검진_소견서.docx"
    pdf_filename = f"26년_{clean_name}_심장건강검진_소견서.pdf"
    
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)
    out_docx_path = os.path.join(output_dir, docx_filename)
    out_pdf_path = os.path.join(output_dir, pdf_filename)

    c_btn1, c_btn2 = st.columns(2)

    with c_btn1:
        if st.button("📄 Word (.docx) 생성하기", type="primary", use_container_width=True):
            with st.spinner("Word 소견서 생성 중..."):
                generate_heart_report_docx(report_data, out_docx_path)
                with open(out_docx_path, "rb") as f:
                    docx_bytes = f.read()
                st.session_state["docx_bytes"] = docx_bytes
                st.session_state["docx_ready"] = True
                st.success("✅ Word 소견서 생성 완료!")

        if st.session_state.get("docx_ready", False):
            st.download_button(
                label=f"📥 {docx_filename} 다운로드",
                data=st.session_state["docx_bytes"],
                file_name=docx_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    with c_btn2:
        if st.button("📑 PDF (.pdf) 변환 및 생성", use_container_width=True):
            with st.spinner("Word 생성 및 고품질 PDF 변환 중..."):
                generate_heart_report_docx(report_data, out_docx_path)
                convert_docx_to_pdf(out_docx_path, out_pdf_path)
                with open(out_pdf_path, "rb") as f:
                    pdf_bytes = f.read()
                st.session_state["pdf_bytes"] = pdf_bytes
                st.session_state["pdf_ready"] = True
                st.success("✅ PDF 변환 완료!")

        if st.session_state.get("pdf_ready", False):
            st.download_button(
                label=f"📥 {pdf_filename} 다운로드",
                data=st.session_state["pdf_bytes"],
                file_name=pdf_filename,
                mime="application/pdf",
                use_container_width=True
            )

