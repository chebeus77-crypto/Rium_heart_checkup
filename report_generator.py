"""
리움동물의료센터 심장 초음파 검진 소견서 자동 생성 모듈
"""
import os
import io
import re
from typing import Optional, Dict, Any
from dataclasses import dataclass
from PIL import Image
import docx

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.docx")


@dataclass
class HeartReportData:
    # 1. 환자 정보
    patient_name: str = ""
    breed: str = ""
    age: str = ""
    gender: str = ""
    owner_name: str = ""

    # 2. 검사 결과
    murmur: str = "심잡음 없음"   # 청진 (예: "심잡음 없음", "심잡음 청진됨 (Grade 3/6)")
    heart_rate: str = ""        # 심박수 (예: "130")
    vhs: str = ""               # VHS (예: "10.8")
    vlas: str = ""              # VLAS (예: "2.6")
    la_ao: str = ""             # LA:Ao (예: "1.85")
    lviddn: str = ""            # LVIDDn (예: "1.74")
    e_peak: str = ""            # E peak (예: "72.5")
    ee_ratio: str = ""          # E/E' (예: "11.2")

    # 3. ACVIM 병기 및 진단
    stage: str = "B1"           # A, B1, B2, C, D
    diagnosis: str = ""         # 진단명
    owner_message: str = ""     # 보호자 안내문

    # 4. 초음파 사진 (바이너리)
    image1_bytes: Optional[bytes] = None
    image2_bytes: Optional[bytes] = None


def evaluate_metrics(data: HeartReportData) -> Dict[str, Dict[str, Any]]:
    """
    검사 결과 수치 자동 판정 로직
    - LA:Ao < 1.6: 정상, >= 1.6: 좌심방 확장
    - LVIDDn < 1.7: 정상, >= 1.7: 좌심실 확장
    - E peak < 60: 정상, >= 60: 좌심방 압력 증가
    - E/E' < 10: 정상, >= 10: 좌심실 충만압 증가/폐수종 위험성 증가
    - 기준값과 같거나 큰 경우 정상 범위 초과로 처리
    - 미입력 항목은 정상으로 간주하지 않고 '-' 처리
    """
    results = {}

    # LA:Ao
    la_ao_str = str(data.la_ao).strip()
    if la_ao_str:
        try:
            val = float(la_ao_str)
            if val < 1.6:
                results["la_ao"] = {"status": "normal", "remark": "정상", "value": val}
            else:
                results["la_ao"] = {"status": "abnormal", "remark": "좌심방 확장", "value": val}
        except ValueError:
            results["la_ao"] = {"status": "unknown", "remark": "-", "value": la_ao_str}
    else:
        results["la_ao"] = {"status": "empty", "remark": "-", "value": ""}

    # LVIDDn
    lviddn_str = str(data.lviddn).strip()
    if lviddn_str:
        try:
            val = float(lviddn_str)
            if val < 1.7:
                results["lviddn"] = {"status": "normal", "remark": "정상", "value": val}
            else:
                results["lviddn"] = {"status": "abnormal", "remark": "좌심실 확장", "value": val}
        except ValueError:
            results["lviddn"] = {"status": "unknown", "remark": "-", "value": lviddn_str}
    else:
        results["lviddn"] = {"status": "empty", "remark": "-", "value": ""}

    # E peak
    e_peak_str = str(data.e_peak).strip()
    if e_peak_str:
        try:
            val = float(e_peak_str)
            if val < 60.0:
                results["e_peak"] = {"status": "normal", "remark": "정상", "value": val}
            else:
                results["e_peak"] = {"status": "abnormal", "remark": "좌심방 압력 증가", "value": val}
        except ValueError:
            results["e_peak"] = {"status": "unknown", "remark": "-", "value": e_peak_str}
    else:
        results["e_peak"] = {"status": "empty", "remark": "-", "value": ""}

    # E/E'
    ee_ratio_str = str(data.ee_ratio).strip()
    if ee_ratio_str:
        try:
            val = float(ee_ratio_str)
            if val < 10.0:
                results["ee_ratio"] = {"status": "normal", "remark": "정상", "value": val}
            else:
                results["ee_ratio"] = {"status": "abnormal", "remark": "좌심실 충만압 증가/폐수종 위험성 증가", "value": val}
        except ValueError:
            results["ee_ratio"] = {"status": "unknown", "remark": "-", "value": ee_ratio_str}
    else:
        results["ee_ratio"] = {"status": "empty", "remark": "-", "value": ""}

    return results


def get_default_owner_message(patient_name: str, stage: str) -> str:
    """ACVIM 병기별 표준 보호자 안내문 생성"""
    name = patient_name.strip() or "환자"
    messages = {
        "A": f"-{name}는 심장병 위험군(Stage A) 단계로 현재 판막 질환이나 심장 비대 등 구조적 이상은 확인되지 않았습니다. 현재 특별한 약물 치료는 필요하지 않으나, 품종적 소인(소형견 노령화)을 고려하여 6-12개월 간격의 정기적인 청진 및 심장 검진을 권장합니다.",
        "B1": f"-{name}는 심장병 B1단계로서 심잡음은 청진되나 심장의 구조적 비대나 심부전 징후는 아직 관찰되지 않는 초기 단계입니다. 현재는 약물 치료를 시작할 단계는 아니며, 질환의 진행 여부를 3-6개월 간격의 정기 검진을 통해 지속적으로 모니터링하는 것이 가장 중요합니다.",
        "B2": f"-{name}는 심장병 B2단계로서 심잡음과 함께 심장의 유의미한 비대(확장) 소견이 확인되었습니다. 이 단계부터는 심부전 발생을 지연시키기 위한 약물 치료(강심제 등) 시작이 강력히 권장되며, 정기적인 검진과 가정 내 안정 시 호흡수(SRR) 체크가 중요합니다.",
        "C": f"-{name}는 심장병 C단계로서 심장병이 많이 진행되어 심장초음파 수치가 높은 상태이며 세밀한 관리가 필요합니다. 평소 분당 안정 시 호흡수를 잘 체크해 주셔야 하며, 무리한 운동산책이나 스트레스는 위험하므로 항상 편안하고 환기가 잘 되는 환경을 유지해 주시기 바랍니다.",
        "D": f"-{name}는 심장병 D단계로서 난치성 심부전에 해당하는 단계입니다. 표준 약물 치료 외에 추가적인 약제 조절과 매우 면밀한 집중 모니터링이 필요하며, 호흡곤란이나 기절 등 응급 증상 발견 시 즉시 병원에 내원하셔야 합니다."
    }
    return messages.get(stage.upper(), messages["B1"])


def _replace_paragraph_text_preserving_format(paragraph, new_text: str):
    """단락의 첫 번째 run 서식을 보존하면서 텍스트 교체 (나머지 run 클리어)"""
    if not paragraph.runs:
        paragraph.add_run(new_text)
        return

    paragraph.runs[0].text = new_text
    for r in paragraph.runs[1:]:
        r.text = ""


def _replace_in_paragraph(paragraph, old_pattern: str, new_text: str):
    """단락 전체 텍스트에서 패턴 검색 후 치환 (첫 run 서식 유지)"""
    full_text = "".join(r.text for r in paragraph.runs)
    if old_pattern in full_text:
        updated = full_text.replace(old_pattern, new_text)
        _replace_paragraph_text_preserving_format(paragraph, updated)
        return True
    return False


def _update_cell_image_ratio(cell, new_image_bytes: bytes, orig_cx: int = 5731510):
    """이미지 종횡비를 계산하여 cell XML 내 extent 높이(cy)를 비율에 맞춰 자동 조정"""
    try:
        with Image.open(io.BytesIO(new_image_bytes)) as img:
            w, h = img.size
            if w > 0:
                new_cy = int(orig_cx * (h / w))
                xml = cell._tc.xml
                # Update wp:extent cy
                xml = re.sub(
                    r'(<wp:extent\s+cx=")' + str(orig_cx) + r'("\s+cy=")(\d+)(")',
                    r'\g<1>' + str(orig_cx) + r'\g<2>' + str(new_cy) + r'\g<4>',
                    xml
                )
                # Update a:ext cy
                xml = re.sub(
                    r'(<a:ext\s+cx=")' + str(orig_cx) + r'("\s+cy=")(\d+)(")',
                    r'\g<1>' + str(orig_cx) + r'\g<2>' + str(new_cy) + r'\g<4>',
                    xml
                )
                new_tc = docx.oxml.parse_xml(xml)
                cell._tc.getparent().replace(cell._tc, new_tc)
    except Exception as e:
        print(f"Ratio update skipped: {e}")


def generate_heart_report_docx(data: HeartReportData, output_path: str, template_path: str = TEMPLATE_PATH) -> str:
    """HeartReportData를 바탕으로 완성된 소견서 docx 파일 생성"""
    eval_res = evaluate_metrics(data)
    doc = docx.Document(template_path)

    # 1. 상단 환자 정보 치환
    for p in doc.paragraphs[:10]:
        _replace_in_paragraph(p, "{{환자명}}", data.patient_name.strip())
        _replace_in_paragraph(p, "{{품종}}", data.breed.strip())
        _replace_in_paragraph(p, "{{나이}}", data.age.strip())
        _replace_in_paragraph(p, "{{성별}}", data.gender.strip())
        _replace_in_paragraph(p, "{{보호자}}", data.owner_name.strip())

    # 2. 1번 검사 항목 표 (doc.tables[0])
    if doc.tables:
        t = doc.tables[0]
        # Row 1: 신체검사 (청진)
        # 템플릿에 "심잡음 {{}}" 형식으로 되어 있음
        murmur_input = data.murmur.strip()
        if not murmur_input:
            formatted_murmur = "-"
        elif murmur_input.startswith("심잡음"):
            formatted_murmur = murmur_input
        else:
            formatted_murmur = f"심잡음 {murmur_input}"

        # 결과 열 (Cell 1, Cell 2)
        for cell in t.rows[1].cells[:3]:
            for p in cell.paragraphs:
                full = "".join(r.text for r in p.runs)
                if "{{}}" in full or "심잡음" in full:
                    _replace_paragraph_text_preserving_format(p, formatted_murmur)

        # 정상 범위 열 (Cell 3): "심잡음 없음"
        if len(t.rows[1].cells) > 3:
            normal_cell = t.rows[1].cells[3]
            if normal_cell.paragraphs:
                _replace_paragraph_text_preserving_format(normal_cell.paragraphs[0], "심잡음 없음")

        # Row 2: 심박수
        hr_val = data.heart_rate.strip()
        hr_text = f"{hr_val}회/분" if hr_val else "-"
        for cell in t.rows[2].cells:
            for p in cell.paragraphs:
                full = "".join(r.text for r in p.runs)
                if "{{심박수}}" in full:
                    _replace_in_paragraph(p, "{{심박수}}회/분", hr_text)
                    _replace_in_paragraph(p, "{{심박수}}", hr_val or "-")

        # Row 3: X-ray (VHS / VLAS)
        vhs_val = data.vhs.strip()
        vlas_val = data.vlas.strip()
        if vhs_val and vlas_val:
            xray_text = f"VHS: {vhs_val}   VLAS: {vlas_val}"
        elif vhs_val:
            xray_text = f"VHS: {vhs_val}"
        elif vlas_val:
            xray_text = f"VLAS: {vlas_val}"
        else:
            xray_text = "-"

        for cell in t.rows[3].cells:
            for p in cell.paragraphs:
                full = "".join(r.text for r in p.runs)
                if "{{VHS}}" in full or "{{VLAS}}" in full:
                    _replace_in_paragraph(p, "{{VHS}}  {{VLAS}}", xray_text)
                    _replace_in_paragraph(p, "{{VHS}} {{VLAS}}", xray_text)
                    _replace_in_paragraph(p, "{{VHS}}", vhs_val or "-")
                    _replace_in_paragraph(p, "{{VLAS}}", vlas_val or "-")

        # Row 5: 초음파 사진 2장 교체 (Cell 0: rId10, Cell 3: rId11)
        if data.image1_bytes:
            try:
                doc.part.related_parts['rId10']._blob = data.image1_bytes
                _update_cell_image_ratio(t.rows[5].cells[0], data.image1_bytes, 5731510)
            except Exception as e:
                print(f"Warning: Failed to replace image1: {e}")

        if data.image2_bytes:
            try:
                doc.part.related_parts['rId11']._blob = data.image2_bytes
                _update_cell_image_ratio(t.rows[5].cells[3], data.image2_bytes, 5731510)
            except Exception as e:
                print(f"Warning: Failed to replace image2: {e}")

        # Row 7: LA:Ao
        la_ao_display = data.la_ao.strip() if data.la_ao.strip() else "-"
        for p in t.rows[7].cells[1].paragraphs:
            if "{{LA:Ao}}" in "".join(r.text for r in p.runs):
                _replace_in_paragraph(p, "{{LA:Ao}}", la_ao_display)
        if t.rows[7].cells[3].paragraphs:
            _replace_paragraph_text_preserving_format(t.rows[7].cells[3].paragraphs[0], eval_res["la_ao"]["remark"])

        # Row 8: LVIDDn
        lviddn_display = data.lviddn.strip() if data.lviddn.strip() else "-"
        for p in t.rows[8].cells[1].paragraphs:
            if "{{LVIDDn}}" in "".join(r.text for r in p.runs):
                _replace_in_paragraph(p, "{{LVIDDn}}", lviddn_display)
        if t.rows[8].cells[3].paragraphs:
            _replace_paragraph_text_preserving_format(t.rows[8].cells[3].paragraphs[0], eval_res["lviddn"]["remark"])

        # Row 9: E peak
        e_peak_display = f"{data.e_peak.strip()}cm/s" if data.e_peak.strip() else "-"
        for p in t.rows[9].cells[1].paragraphs:
            full = "".join(r.text for r in p.runs)
            if "{{E peak}}" in full:
                _replace_in_paragraph(p, "{{E peak}}cm/s", e_peak_display)
                _replace_in_paragraph(p, "{{E peak}}", data.e_peak.strip())
        if t.rows[9].cells[3].paragraphs:
            _replace_paragraph_text_preserving_format(t.rows[9].cells[3].paragraphs[0], eval_res["e_peak"]["remark"])

        # Row 10: E/E'
        ee_display = data.ee_ratio.strip() if data.ee_ratio.strip() else "-"
        for p in t.rows[10].cells[1].paragraphs:
            full = "".join(r.text for r in p.runs)
            if "E/E" in full:
                _replace_paragraph_text_preserving_format(p, ee_display)
        if t.rows[10].cells[3].paragraphs:
            _replace_paragraph_text_preserving_format(t.rows[10].cells[3].paragraphs[0], eval_res["ee_ratio"]["remark"])

    # 3. 2. 진단명 치환
    stage_upper = data.stage.strip().upper() or "B1"
    diag_text = data.diagnosis.strip() or f"■ 승모판 점액종성 판막 질환 (MMVD) Stage {stage_upper}"
    for p in doc.paragraphs:
        full = "".join(r.text for r in p.runs)
        if "승모판 점액종성 판막 질환 (MMVD) A/B1/B2/C/D" in full or "승모판 점액종성 판막 질환 (MMVD)" in full:
            _replace_paragraph_text_preserving_format(p, diag_text)
            break

    # 4. 4. ACVIM 현재 병기 치환
    for p in doc.paragraphs:
        full = "".join(r.text for r in p.runs)
        if "현재 병기: Stage" in full:
            _replace_paragraph_text_preserving_format(p, f"현재 병기: Stage {stage_upper}")
            break

    # 5. 7. 보호자 안내문 치환
    owner_msg = data.owner_message.strip()
    if not owner_msg:
        owner_msg = get_default_owner_message(data.patient_name, stage_upper)

    for i, p in enumerate(doc.paragraphs):
        full = "".join(r.text for r in p.runs)
        if "7. 보호자님께 드리는 말씀" in full:
            if i + 1 < len(doc.paragraphs):
                target_p = doc.paragraphs[i + 1]
                _replace_paragraph_text_preserving_format(target_p, owner_msg)
            break

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    return output_path


def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> str:
    """Windows Word COM을 사용하여 docx를 pdf로 고품질 변환"""
    import win32com.client
    import pythoncom

    docx_abs = os.path.abspath(docx_path)
    pdf_abs = os.path.abspath(pdf_path)

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        doc = word.Documents.Open(docx_abs, ReadOnly=True)
        doc.SaveAs2(pdf_abs, FileFormat=17) # 17 = wdFormatPDF
        doc.Close(SaveChanges=False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()

    return pdf_abs
