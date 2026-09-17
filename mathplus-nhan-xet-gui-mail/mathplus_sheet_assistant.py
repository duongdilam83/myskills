#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MathPlus Sheet Assistant - Quản lý Nhận xét BTVN & Gửi Email Phụ huynh
Trung tâm Giáo dục MathPlus Academy

Hỗ trợ gửi email KHÔNG CẦN MẬT KHẨU:
1. macos-mail: Sử dụng trực tiếp tài khoản Google / iCloud đã đăng nhập sẵn trong macOS Mail.app (Zero-Password).
2. service-account: Sử dụng Google Service Account với Domain-Wide Delegation (dành cho Google Workspace).
3. smtp: Gửi qua SMTP truyền thống (với App Password).
"""

import os
import sys

# Đảm bảo in tiếng Việt chuẩn trên Windows console
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

import re
import json
import base64
import smtplib
import argparse
import hashlib
import subprocess
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Dict, Any, Optional, Tuple


import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

POSSIBLE_CREDS_PATHS = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credentials.json'),
    os.path.expanduser('~/.gemini/config/skills/mathplus-nhan-xet-gui-mail/credentials.json'),
    os.path.expanduser('~/.gemini/config/plugins/myskills/skills/mathplus-nhan-xet-gui-mail/credentials.json'),
    os.path.expanduser('~/Documents/credentials.json'),
    os.path.expanduser('~/Documents/service_account.json'),
    '/Volumes/data/code/skills/mcp-google-sheets/credentials.json',
    os.path.expanduser('~/.gemini/config/skills/mcp-google-sheets/credentials.json'),
    os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', '')
]

POSSIBLE_CONFIG_PATHS = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json'),
    os.path.expanduser('~/.gemini/config/skills/mathplus-nhan-xet-gui-mail/config.json'),
    os.path.expanduser('~/.gemini/config/plugins/myskills/skills/mathplus-nhan-xet-gui-mail/config.json'),
    os.path.expanduser('~/Documents/email_config.json'),
    os.path.expanduser('~/Documents/config.json')
]

def find_credentials_file() -> Optional[str]:
    for path in POSSIBLE_CREDS_PATHS:
        if path and os.path.exists(path):
            return path
    return None

def load_smtp_config() -> Dict[str, Any]:
    for p in POSSIBLE_CONFIG_PATHS:
        if p and os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}


def get_gspread_client(custom_creds: Optional[str] = None):
    creds_path = custom_creds or find_credentials_file()
    if not creds_path or not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"Không tìm thấy credentials.json cho Google Sheets API. "
            f"Vui lòng đặt credentials.json tại thư mục skill hoặc chỉ định qua --credentials."
        )
    credentials = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(credentials)

def parse_sheet_url(url: str) -> Tuple[str, Optional[int]]:
    """Trích xuất spreadsheet_id và gid từ URL Google Sheets."""
    sheet_id_match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    if not sheet_id_match:
        spreadsheet_id = url
    else:
        spreadsheet_id = sheet_id_match.group(1)
        
    gid_match = re.search(r'gid=([0-9]+)', url)
    gid = int(gid_match.group(1)) if gid_match else None
    return spreadsheet_id, gid

def get_worksheet(client: gspread.Client, spreadsheet_id: str, gid: Optional[int] = None):
    sheet = client.open_by_key(spreadsheet_id)
    if gid is not None:
        for ws in sheet.worksheets():
            if ws.id == gid:
                return ws
    return sheet.sheet1

# ==================== BỘ SINH NHẬN XÉT SƯ PHẠM MATHPLUS ====================

REMARK_POOLS = {
    "excellent": { # 9.5 - 10.0
        "prep": [
            "- Bài làm chuẩn mực, thể hiện tinh thần tự học và chuẩn bị bài rất chu đáo.",
            "- Con hoàn thành xuất sắc toàn bộ bài tập về nhà được giao, nộp bài đúng hạn.",
            "- Ý thức tự giác học tập rất cao, bài vở chỉn chu và cẩn thận."
        ],
        "comprehension": [
            "- Khả năng tiếp thu bài vượt trội, nắm bắt bản chất các dạng toán rất nhanh và sâu sắc.",
            "- Tư duy logic nhạy bén, phương pháp giải toán linh hoạt và trình bày mạch lạc.",
            "- Con nắm chắc phương pháp giải các bài toán nâng cao và vận dụng rất hiệu quả."
        ],
        "advice": [
            "- Tiếp tục rèn luyện thêm các dạng bài toán mở rộng và tư duy sâu để bứt phá.",
            "- Nên duy trì thói quen đào sâu suy nghĩ tìm nhiều cách giải khác nhau cho một bài toán.",
            "- Giữ vững sự cẩn trọng và kiểm tra kỹ từng bước biến đổi."
        ],
        "encourage": [
            "- Xuất sắc! Thầy cô rất tự hào về tinh thần học tập và năng lực toán học của con.",
            "- Thầy tin con sẽ luôn là tấm gương học tập xuất sắc của lớp!",
            "- Tiếp tục phát huy niềm say mê với môn Toán con nhé!"
        ]
    },
    "very_good": { # 8.0 - 9.0
        "prep": [
            "- Con nộp bài đúng hạn và hoàn thành đầy đủ bài tập được giao.",
            "- Bài tập về nhà được hoàn thành rất tốt, chất lượng bài làm cao.",
            "- Con có ý thức tự giác làm bài tập về nhà rất tốt, trình bày sạch đẹp."
        ],
        "comprehension": [
            "- Con hiểu bài rất nhanh và nắm chắc các dạng bài từ cơ bản đến nâng cao.",
            "- Tư duy toán học sắc sảo, tìm ra hướng giải bài toán một cách nhanh chóng.",
            "- Kỹ năng tính toán tốt, các bước biến đổi rõ ràng, dễ hiểu."
        ],
        "advice": [
            "- Nên kiểm tra lại từng dòng biến đổi trước khi kết luận đáp số để đạt điểm tuyệt đối.",
            "- Tránh tâm lý chủ quan ở các câu hỏi dễ để không mất điểm đáng tiếc.",
            "- Con hiểu bài nhưng cần cẩn thận hơn một chút trong tính toán để tránh nhầm lẫn số học."
        ],
        "encourage": [
            "- Bài làm rất tốt, thầy tin con sẽ sớm đạt điểm tối đa ở các bài tiếp theo!",
            "- Tiếp tục phát huy phong độ và sự tự tin này con nhé!",
            "- Chúc mừng con với kết quả học tập rất ấn tượng ở buổi học này."
        ]
    },
    "good": { # 7.0 - 7.5
        "prep": [
            "- Con có ý thức làm bài tập về nhà và nộp bài đầy đủ.",
            "- Đã cố gắng hoàn thành các câu hỏi trong phiếu bài tập được giao.",
            "- Vở ghi và bài làm tương đối rõ ràng, có trách nhiệm với việc học."
        ],
        "comprehension": [
            "- Con nắm được kiến thức trọng tâm và làm tốt các bài tập cơ bản.",
            "- Có nền tảng tư duy tốt, tuy nhiên ở các bài toán nâng cao còn đôi chút lúng túng.",
            "- Đã vận dụng được công thức và phương pháp thầy cô hướng dẫn trên lớp."
        ],
        "advice": [
            "- Cần đọc kỹ đề bài hơn, chú ý điều kiện xác định và các bước suy luận trung gian.",
            "- Nên dành thêm thời gian kiểm tra lại kết quả tính toán trước khi nộp bài.",
            "- Rà soát lại những câu còn sai để hiểu rõ nguyên nhân và rút kinh nghiệm."
        ],
        "encourage": [
            "- Con có nhiều tiềm năng, chỉ cần cẩn thận hơn chắc chắn điểm số sẽ bứt phá!",
            "- Cố gắng rèn luyện thêm mỗi ngày con nhé, thầy cô luôn sẵn sàng hỗ trợ con.",
            "- Hãy tự tin hơn trong các bài toán mới, con hoàn toàn có thể làm tốt hơn nữa!"
        ]
    },
    "average": { # 5.0 - 6.5
        "prep": [
            "- Con đã nộp bài tập về nhà nhưng số lượng bài hoàn thành chưa trọn vẹn.",
            "- Bài tập về nhà còn thiếu một số bài hoặc làm còn sơ sài.",
            "- Cần rèn luyện thêm tính kỷ luật và dành thời gian nghiêm túc hơn cho BTVN."
        ],
        "comprehension": [
            "- Con nắm bài ở mức cơ bản, cần củng cố thêm các định nghĩa và phương pháp giải.",
            "- Kỹ năng tính toán còn dễ nhầm lẫn số, cần tập trung cao độ hơn khi làm bài.",
            "- Ở những bài toán suy luận logic, con cần suy nghĩ kỹ từng bước."
        ],
        "advice": [
            "- Con cần xem lại thật kỹ bài giảng trên lớp và các ví dụ thầy cô đã chữa mẫu.",
            "- Hãy chủ động hỏi lại thầy cô ngay những phần bài tập con chưa hiểu rõ.",
            "- Đề nghị phụ huynh đôn đốc con hoàn thành lại những bài chưa đạt yêu cầu."
        ],
        "encourage": [
            "- Đừng nản lòng con nhé, kiên trì tự học từng ngày con sẽ thấy môn Toán thú vị và dễ dàng hơn.",
            "- Thầy cô tin rằng nếu con tập trung hơn, con sẽ có sự tiến bộ rõ rệt.",
            "- Cố gắng nhiều hơn ở buổi học tiếp theo con nhé!"
        ]
    },
    "missing": { # < 5.0 hoặc "chưa nộp", "thiếu", "chưa đạt", "0"
        "prep": [
            "- Con chưa hoàn thành đầy đủ bài tập về nhà theo yêu cầu của thầy cô.",
            "- Bài tập về nhà chưa nộp đúng hạn hoặc số lượng bài làm quá ít.",
            "- Ý thức tự giác làm bài tập về nhà tuần này chưa tốt."
        ],
        "comprehension": [
            "- Việc không làm BTVN sẽ ảnh hưởng trực tiếp đến khả năng tiếp thu bài mới trên lớp.",
            "- Con cần nghiêm túc chấn chỉnh lại thái độ học tập và tính tự giác của mình."
        ],
        "advice": [
            "- Yêu cầu con hoàn thành bổ sung đầy đủ toàn bộ bài tập về nhà trước buổi học tới.",
            "- Đề nghị Quý Phụ huynh phối hợp chặt chẽ, nhắc nhở và kiểm tra con làm bài ở nhà.",
            "- Dành thời gian ôn lại lý thuyết và liên hệ thầy cô để được hướng dẫn bổ sung."
        ],
        "encourage": [
            "- Thầy cô mong con nghiêm túc rút kinh nghiệm để không bị hổng kiến thức quan trọng.",
            "- Hãy nỗ lực bắt nhịp lại ngay từ buổi học tới con nhé!"
        ]
    }
}

def parse_score_value(score_str: str) -> Tuple[Optional[float], str]:
    """Chuyển chuỗi điểm số thành số thực và loại phân loại."""
    s = str(score_str).strip()
    if not s:
        return None, "empty"
    
    clean_s = s.replace(',', '.')
    
    try:
        val = float(clean_s)
        if val >= 9.25:
            return val, "excellent"
        elif val >= 8.0:
            return val, "very_good"
        elif val >= 7.0:
            return val, "good"
        elif val >= 5.0:
            return val, "average"
        else:
            return val, "missing"
    except ValueError:
        lower_s = s.lower()
        if any(w in lower_s for w in ["chưa", "thiếu", "không", "k nộp", "k làm", "0"]):
            return 0.0, "missing"
        return None, "unknown"

def generate_remark(score_str: str, student_name: str, date_str: str) -> str:
    """Sinh lời nhận xét chuẩn sư phạm dựa trên điểm BTVN và tên học sinh."""
    val, category = parse_score_value(score_str)
    
    if category == "empty" or category == "unknown":
        return ""
    
    pool = REMARK_POOLS.get(category, REMARK_POOLS["very_good"])
    
    seed_str = f"{student_name}_{date_str}_{score_str}"
    h = int(hashlib.md5(seed_str.encode('utf-8')).hexdigest(), 16)
    
    prep_line = pool["prep"][h % len(pool["prep"])]
    h //= len(pool["prep"])
    
    comp_line = pool["comprehension"][h % len(pool["comprehension"])]
    h //= len(pool["comprehension"])
    
    advice_line = pool["advice"][h % len(pool["advice"])]
    h //= len(pool["advice"])
    
    enc_line = pool["encourage"][h % len(pool["encourage"])]
    
    lines = [prep_line, comp_line, advice_line, enc_line]
    return "\n".join(lines)

# ==================== PHÂN TÍCH VÀ CẬP NHẬT GOOGLE SHEET ====================

def analyze_sheet_data(worksheet_data: List[List[str]]) -> Dict[str, Any]:
    classes = []
    current_class = None
    
    for r_idx, row in enumerate(worksheet_data):
        row_str = " ".join(row).strip()
        if not row_str:
            continue
            
        first_cell = row[0].strip() if len(row) > 0 else ""
        if first_cell.startswith("Lớp ") or first_cell.startswith("Lớp"):
            class_name = first_cell
            date_cols = {}
            for c_idx, cell in enumerate(row):
                c_val = cell.strip()
                if re.match(r'^\d{1,2}/\d{1,2}$', c_val):
                    date_cols[c_val] = c_idx
                    
            current_class = {
                "name": class_name,
                "header_row_idx": r_idx,
                "date_cols": date_cols,
                "sub_headers": {},
                "students": []
            }
            classes.append(current_class)
            continue
            
        if current_class and ("STT" in row or "Tên" in row):
            for date, col_idx in current_class["date_cols"].items():
                current_class["sub_headers"][date] = {
                    "diem_danh": col_idx,
                    "btvn": col_idx + 1,
                    "nhan_xet": col_idx + 2
                }
            continue
            
        if current_class and first_cell.isdigit():
            stt = first_cell
            name = row[1].strip() if len(row) > 1 else ""
            phone = row[2].strip() if len(row) > 2 else ""
            email = row[3].strip() if len(row) > 3 else ""
            
            if not email or "@" not in email:
                for check_c in range(2, min(5, len(row))):
                    if "@" in row[check_c]:
                        email = row[check_c].strip()
                        break
                        
            student_data = {
                "row_number": r_idx + 1,
                "stt": stt,
                "name": name,
                "phone": phone,
                "email": email,
                "class_name": current_class["name"],
                "sessions": {}
            }
            
            for date, cols in current_class.get("sub_headers", {}).items():
                dd_col = cols["diem_danh"]
                btvn_col = cols["btvn"]
                nx_col = cols["nhan_xet"]
                
                dd_val = row[dd_col].strip() if dd_col < len(row) else ""
                btvn_val = row[btvn_col].strip() if btvn_col < len(row) else ""
                nx_val = row[nx_col].strip() if nx_col < len(row) else ""
                
                student_data["sessions"][date] = {
                    "diem_danh": dd_val,
                    "btvn": btvn_val,
                    "nhan_xet": nx_val,
                    "btvn_col_idx": btvn_col + 1,
                    "nx_col_idx": nx_col + 1
                }
                
            current_class["students"].append(student_data)
            
    return {"classes": classes}

def find_active_dates(classes_data: Dict[str, Any]) -> List[str]:
    dates_found = []
    for cls in classes_data["classes"]:
        for d in cls["date_cols"].keys():
            if d not in dates_found:
                dates_found.append(d)
    return dates_found

def get_students_for_session(classes_data: Dict[str, Any], session_date: str) -> List[Dict[str, Any]]:
    results = []
    for cls in classes_data["classes"]:
        for stu in cls["students"]:
            sess = stu["sessions"].get(session_date)
            if sess:
                btvn = sess["btvn"]
                results.append({
                    "row_number": stu["row_number"],
                    "class_name": stu["class_name"],
                    "stt": stu["stt"],
                    "name": stu["name"],
                    "phone": stu["phone"],
                    "email": stu["email"],
                    "session_date": session_date,
                    "diem_danh": sess["diem_danh"],
                    "btvn": btvn,
                    "current_nhan_xet": sess["nhan_xet"],
                    "nx_col_idx": sess["nx_col_idx"]
                })
    return results

# ==================== RENDER EMAIL CONTENT ====================

def render_email_content(student_info: Dict[str, Any], remark_text: str, template_path: Optional[str] = None) -> Tuple[str, str]:
    tpl_file = template_path or os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates", "email_template.html")
    html_tpl = ""
    if os.path.exists(tpl_file):
        with open(tpl_file, "r", encoding="utf-8") as f:
            html_tpl = f.read()
    else:
        html_tpl = """
        <html><body>
        <h2>MathPlus Academy - Kết quả học tập ngày {{session_date}}</h2>
        <p>Kính gửi Quý Phụ huynh học sinh <strong>{{student_name}}</strong> ({{class_name}}),</p>
        <p>Điểm BTVN: <strong>{{btvn_score}}</strong></p>
        <h3>Nhận xét chi tiết:</h3>
        {{remarks_html}}
        </body></html>
        """
        
    remarks_list = [line.strip().lstrip("-").strip() for line in remark_text.split("\n") if line.strip()]
    remarks_html = "<ul>" + "".join([f"<li>{r}</li>" for r in remarks_list]) + "</ul>"
    
    score_display = student_info.get("btvn") or "Chưa có điểm"
    
    rendered_html = html_tpl.replace("{{student_name}}", student_info["name"])
    rendered_html = rendered_html.replace("{{class_name}}", student_info["class_name"])
    rendered_html = rendered_html.replace("{{session_date}}", student_info["session_date"])
    rendered_html = rendered_html.replace("{{btvn_score}}", score_display)
    rendered_html = rendered_html.replace("{{remarks_html}}", remarks_html)
    
    bullet_remarks = "\n".join([f"• {r}" for r in remarks_list])

    plain_text = f"""Kính gửi Quý Phụ huynh học sinh {student_info['name']},

TRUNG TÂM GIÁO DỤC MATHPLUS ACADEMY
Thông báo kết quả học tập và nhận xét buổi học ngày {student_info['session_date']}

- Học sinh: {student_info['name']}
- Lớp học: {student_info['class_name']}
- Ngày học: {student_info['session_date']}
- Điểm bài tập về nhà (BTVN): {score_display}

📝 NHẬN XÉT CHI TIẾT TỪ THẦY/CÔ PHỤ TRÁCH:
{bullet_remarks}

💡 LỜI NHẮN GỬI TỪ MATHPLUS ACADEMY:
Mỗi bước tiến bộ của con dù nhỏ đều là kết quả của sự nỗ lực và tự giác rèn luyện. Thầy cô mong Quý Phụ huynh tiếp tục đồng hành, động viên con giữ vững niềm đam mê và thói quen tự học toán mỗi ngày!

Nếu Quý Phụ huynh có bất kỳ trao đổi nào, xin vui lòng phản hồi email này hoặc liên hệ trung tâm.

---
TRUNG TÂM GIÁO DỤC MATHPLUS ACADEMY
Cơ sở 1: MathPlus Xã Đàn | Cơ sở 2: MathPlus Hồ Đắc Di, Hà Nội
Website: https://mathplus.com.vn | Email: cskh@mathplus.vn
"""
    return plain_text, rendered_html

# ==================== CÁC PHƯƠNG THỨC GỬI EMAIL KHÔNG MẬT KHẨU ====================

def get_macos_mail_accounts() -> List[Dict[str, str]]:
    """Lấy danh sách các tài khoản email đã đăng nhập sẵn trong ứng dụng macOS Mail."""
    if sys.platform != "darwin":
        return []
    script = """
    tell application "Mail"
        set accList to {}
        repeat with acc in accounts
            set end of accList to (name of acc & "|" & (email addresses of acc as string))
        end repeat
        return accList
    end tell
    """
    try:
        res = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=10)
        if res.returncode != 0 or not res.stdout.strip():
            return []
        items = res.stdout.strip().split(", ")
        accounts = []
        for it in items:
            if "|" in it:
                acc_name, emails = it.split("|", 1)
                accounts.append({"name": acc_name.strip(), "emails": [e.strip() for e in emails.split(",") if e.strip()]})
        return accounts
    except Exception:
        return []

def send_via_macos_mail(
    to_email: str,
    subject: str,
    body_text: str,
    sender_email: Optional[str] = None,
    open_draft_only: bool = False
) -> bool:
    """
    Gửi email trực tiếp qua macOS Mail.app KHÔNG CẦN NHẬP MẬT KHẨU!
    Tự động dùng phiên đăng nhập Google / iCloud đã lưu trong macOS Keychain.
    """
    if sys.platform != "darwin":
        raise RuntimeError("Phương thức macos-mail chỉ hỗ trợ trên hệ điều hành macOS.")

    # Escape cho AppleScript
    def escape_applescript(s: str) -> str:
        return s.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n').replace('\r', '')

    safe_subject = escape_applescript(subject)
    safe_body = escape_applescript(body_text)
    
    sender_clause = f'set sender of newMessage to "{sender_email}"' if sender_email else ""
    send_action = "set visible of newMessage to true" if open_draft_only else "send newMessage"

    applescript = f"""
    tell application "Mail"
        set newMessage to make new outgoing message with properties {{subject:"{safe_subject}", content:"{safe_body}", visible:false}}
        tell newMessage
            make new to recipient at end of to recipients with properties {{address:"{to_email}"}}
            {sender_clause}
        end tell
        {send_action}
    end tell
    """

    res = subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True, timeout=15)
    if res.returncode != 0:
        raise RuntimeError(f"Lỗi macOS Mail: {res.stderr.strip()}")
    return True

def send_via_service_account(
    creds_path: str,
    impersonate_user: str,
    to_email: str,
    student_name: str,
    session_date: str,
    plain_text: str,
    html_text: str
) -> bool:
    """
    Gửi email qua Google Workspace Domain-Wide Delegation (Service Account).
    Không cần mật khẩu của user, Service Account tự cấp quyền gửi mail.
    """
    import urllib.request
    from google.oauth2.service_account import Credentials
    import google.auth.transport.requests

    scopes = ['https://www.googleapis.com/auth/gmail.send']
    credentials = Credentials.from_service_account_file(
        creds_path,
        scopes=scopes,
        subject=impersonate_user
    )
    request = google.auth.transport.requests.Request()
    credentials.refresh(request)

    subject = f"[MathPlus Academy] Kết quả học tập & Nhận xét buổi học ngày {session_date} - Học sinh {student_name}"
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Trung tâm Giáo dục MathPlus Academy <{impersonate_user}>"
    msg["To"] = to_email
    msg.attach(MIMEText(plain_text, "plain", "utf-8"))
    msg.attach(MIMEText(html_text, "html", "utf-8"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')
    url = f"https://gmail.googleapis.com/gmail/v1/users/{impersonate_user}/messages/send"
    data = json.dumps({"raw": raw}).encode('utf-8')

    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Authorization": f"Bearer {credentials.token}",
            "Content-Type": "application/json"
        },
        method="POST"
    )

    with urllib.request.urlopen(req) as resp:
        return resp.status in (200, 202)

def send_via_windows_outlook(
    to_email: str,
    subject: str,
    body_text: str,
    html_text: str,
    open_draft_only: bool = False
) -> bool:
    """
    Gửi email qua Microsoft Outlook trên Windows (Zero-Password).
    Tận dụng tài khoản đã đăng nhập sẵn trong Outlook máy tính mà không cần mật khẩu.
    """
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("Để gửi qua Outlook trên Windows, vui lòng cài đặt pywin32: pip install pywin32")
        
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0) # 0 = olMailItem
    mail.To = to_email
    mail.Subject = subject
    mail.HTMLBody = html_text
    
    if open_draft_only:
        mail.Display() # Mở cửa sổ Outlook trên màn hình để người dùng tự xem và gửi
    else:
        mail.Send()
    return True

def send_via_smtp(
    smtp_cfg: Dict[str, Any],
    to_email: str,
    student_name: str,
    session_date: str,
    plain_text: str,
    html_text: str
) -> bool:
    """Gửi email qua SMTP truyền thống với Google App Password."""
    sender_email = smtp_cfg.get("sender_email") or smtp_cfg.get("email")
    sender_pass = smtp_cfg.get("sender_password") or smtp_cfg.get("password") or smtp_cfg.get("app_password")
    
    if not sender_email or not sender_pass:
        raise ValueError("Chưa cấu hình tài khoản SMTP (sender_email và sender_password / app_password).")
        
    sender_name = smtp_cfg.get("sender_name", "MathPlus Academy")
    subject = f"[MathPlus Academy] Kết quả học tập & Nhận xét buổi học ngày {session_date} - Học sinh {student_name}"
    
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = to_email
    msg.attach(MIMEText(plain_text, "plain", "utf-8"))
    msg.attach(MIMEText(html_text, "html", "utf-8"))
    
    server_host = smtp_cfg.get("smtp_server", "smtp.gmail.com")
    server_port = int(smtp_cfg.get("smtp_port", 465))
    clean_pass = sender_pass.replace(" ", "").strip()
    
    if server_port == 465:
        with smtplib.SMTP_SSL(server_host, server_port, timeout=15) as server:
            server.login(sender_email, clean_pass)
            server.sendmail(sender_email, [to_email], msg.as_string())
    else:
        with smtplib.SMTP(server_host, server_port, timeout=15) as server:
            server.starttls()
            server.login(sender_email, clean_pass)
            server.sendmail(sender_email, [to_email], msg.as_string())
        
    return True

# ==================== CLI & MAIN ====================

def main():
    parser = argparse.ArgumentParser(
        description="MathPlus Sheet Assistant: Quản lý nhận xét BTVN & Gửi Email Phụ huynh (Zero Password)"
    )
    parser.add_argument("--sheet-url", required=True, help="URL Google Sheet hoặc Spreadsheet ID")
    parser.add_argument("--gid", type=int, default=None, help="Worksheet GID (nếu có)")
    parser.add_argument("--date", default=None, help="Ngày học cần xử lý (ví dụ: '12/09'). Mặc định tự phát hiện.")
    parser.add_argument("--credentials", default=None, help="Đường dẫn file credentials.json")
    parser.add_argument(
        "--mailer",
        choices=["auto", "macos-mail", "outlook", "service-account", "smtp"],
        default="auto",
        help="Cơ chế gửi: auto (tự chọn), macos-mail (Mac Mail), outlook (Windows Outlook), service-account (Google Workspace), smtp (App Password)"
    )
    parser.add_argument("--sender", default=None, help="Email người gửi (ví dụ: mathplus.edu.vn1987@gmail.com)")
    parser.add_argument("--impersonate-user", default=None, help="Email user Google Workspace cần impersonate khi dùng service-account")
    parser.add_argument(
        "--open-drafts",
        action="store_true",
        help="Chỉ tạo bản nháp (Draft) mở trên màn hình macOS Mail để người dùng tự bấm gửi"
    )
    parser.add_argument(
        "--action",
        choices=["preview", "update-sheet", "send-email", "all"],
        default="preview",
        help="Thao tác: preview (chỉ xem trước), update-sheet (cập nhật Google Sheet), send-email (gửi email), all (cập nhật + gửi email)"
    )
    parser.add_argument("--overwrite", action="store_true", help="Ghi đè nhận xét kể cả khi ô đã có nhận xét cũ")
    parser.add_argument("--confirm", action="store_true", help="Xác nhận gửi email từ người dùng (bắt buộc để gửi email)")
    parser.add_argument("--export-preview", default=None, help="Xuất file HTML xem trước danh sách email và nhận xét")

    args = parser.parse_args()

    # 1. Khởi tạo kết nối Google Sheet
    spreadsheet_id, url_gid = parse_sheet_url(args.sheet_url)
    target_gid = args.gid if args.gid is not None else url_gid

    client = get_gspread_client(args.credentials)
    worksheet = get_worksheet(client, spreadsheet_id, target_gid)
    print(f"[*] Đã kết nối Google Sheet: '{worksheet.spreadsheet.title}' | Tab: '{worksheet.title}' (gid: {worksheet.id})")

    # 2. Phân tích dữ liệu bảng tính
    all_values = worksheet.get_all_values()
    parsed_data = analyze_sheet_data(all_values)

    all_dates = find_active_dates(parsed_data)
    if not all_dates:
        print("[!] Không tìm thấy cột ngày học nào trong định dạng dd/mm (ví dụ 12/09).", file=sys.stderr)
        sys.exit(1)

    target_date = args.date
    if not target_date:
        for d in reversed(all_dates):
            stus = get_students_for_session(parsed_data, d)
            if any(s["btvn"].strip() for s in stus):
                target_date = d
                break
        if not target_date:
            target_date = all_dates[0]

    print(f"[*] Ngày học được chọn: '{target_date}' (Các ngày có trong sheet: {', '.join(all_dates)})")

    students = get_students_for_session(parsed_data, target_date)
    students_with_scores = [s for s in students if s["btvn"].strip()]

    print(f"[*] Tìm thấy {len(students)} học sinh, trong đó {len(students_with_scores)} học sinh đã có Điểm BTVN.")

    if not students_with_scores:
        print("[i] Chưa có học sinh nào được nhập điểm BTVN cho ngày này.")
        return

    tasks = []
    for s in students_with_scores:
        current_nx = s["current_nhan_xet"]
        needs_update = args.overwrite or (not current_nx.strip())
        generated_nx = generate_remark(s["btvn"], s["name"], target_date)
        
        final_nx = generated_nx if needs_update else current_nx
        tasks.append({
            "student": s,
            "score": s["btvn"],
            "needs_sheet_update": needs_update,
            "remark": final_nx,
            "has_email": bool(s["email"].strip() and "@" in s["email"]),
            "email": s["email"].strip()
        })

    # ==================== HIỂN THỊ PREVIEW ====================

    print("\n" + "="*80)
    print(f" BẢNG XEM TRƯỚC NHẬN XÉT BTVN - BUỔI HỌC NGÀY {target_date}")
    print("="*80)
    for i, t in enumerate(tasks, 1):
        st = t["student"]
        email_status = f"✅ {t['email']}" if t['has_email'] else "❌ Chưa có email"
        update_status = "Sẽ cập nhật" if t['needs_sheet_update'] else "Giữ nguyên nhận xét cũ"
        print(f"\n{i}. [{st['class_name']}] {st['name']} (Hàng {st['row_number']})")
        print(f"   - Điểm BTVN: {t['score']}")
        print(f"   - Email PH:  {email_status}")
        print(f"   - Trạng thái Sheet: {update_status}")
        print(f"   - Nội dung nhận xét:")
        for line in t['remark'].split("\n"):
            print(f"     {line}")
    print("\n" + "="*80)

    if args.export_preview:
        html_cards = []
        for t in tasks:
            st = t["student"]
            plain, html_body = render_email_content(st, t["remark"])
            html_cards.append(f"""
            <div style="border: 1px solid #ccc; margin-bottom: 20px; padding: 15px; border-radius: 8px; background: #fff;">
              <h3>[{st['class_name']}] {st['name']} - Điểm: {t['score']} - Email: {t['email']}</h3>
              {html_body}
            </div>
            """)
        with open(args.export_preview, "w", encoding="utf-8") as f:
            f.write(f"<html><body style='background:#eef2ee; padding: 20px; font-family:sans-serif;'>{''.join(html_cards)}</body></html>")
        print(f"[+] Đã xuất file HTML xem trước tại: {args.export_preview}")

    # ==================== CẬP NHẬT SHEET ====================

    if args.action in ["update-sheet", "all"]:
        print(f"\n[*] Đang tiến hành cập nhật nhận xét vào Google Sheet...")
        updates_to_make = [t for t in tasks if t["needs_sheet_update"]]
        if not updates_to_make:
            print("[i] Tất cả các ô nhận xét đã có dữ liệu (dùng --overwrite nếu muốn ghi đè).")
        else:
            cells_to_update = []
            for t in updates_to_make:
                row_idx = t["student"]["row_number"]
                col_idx = t["student"]["nx_col_idx"]
                cells_to_update.append(gspread.Cell(row=row_idx, col=col_idx, value=t["remark"]))
                
            worksheet.update_cells(cells_to_update)
            print(f"[✅] Đã cập nhật thành công {len(cells_to_update)} nhận xét vào cột 'Điểm trên lớp + Nhận xét' trên Google Sheet!")

    # ==================== GỬI EMAIL PHỤ HUYNH ====================

    if args.action in ["send-email", "all"]:
        email_tasks = [t for t in tasks if t["has_email"]]
        print(f"\n[*] Chuẩn bị gửi email cho {len(email_tasks)} phụ huynh có email.")

        # CHỐT CHẶN AN TOÀN
        if not args.confirm:
            print("\n" + "!"*80, file=sys.stderr)
            print(" [DỪNG LẠI] CHỐT CHẶN AN TOÀN - YÊU CẦU HUMAN CAN THIỆP PHÊ DUYỆT!", file=sys.stderr)
            print(" Thao tác gửi email đến phụ huynh KHÔNG ĐƯỢC PHÉP chạy tự động.", file=sys.stderr)
            print(" Vui lòng kiểm tra kỹ nội dung bảng preview ở trên.", file=sys.stderr)
            print(" Nếu đồng ý gửi, hãy bổ sung cờ --confirm để tiến hành gửi email.", file=sys.stderr)
            print("!"*80 + "\n", file=sys.stderr)
            sys.exit(2)

        # Quyết định Mailer backend
        mailer = args.mailer
        sender_email = args.sender
        smtp_cfg = load_smtp_config()



        if mailer == "auto":
            if smtp_cfg.get("sender_email") and (smtp_cfg.get("sender_password") or smtp_cfg.get("password")):
                mailer = "smtp"
                sender_email = sender_email or smtp_cfg.get("sender_email")
                print(f"[*] Tự động phát hiện cấu hình SMTP trong config.json (Tài khoản: {sender_email})")
            elif sys.platform == "darwin":
                macos_accounts = get_macos_mail_accounts()
                if macos_accounts:
                    mailer = "macos-mail"
                    google_accs = [a for a in macos_accounts if "google" in a["name"].lower() or any("@gmail.com" in e for e in a["emails"])]
                    if google_accs and not sender_email:
                        sender_email = google_accs[0]["emails"][0]
                    elif not sender_email and macos_accounts[0]["emails"]:
                        sender_email = macos_accounts[0]["emails"][0]
                    print(f"[*] Tự động phát hiện phương thức: 'macos-mail' (Tài khoản: {sender_email}) - KHÔNG CẦN MẬT KHẨU!")
                else:
                    mailer = "smtp"
            elif sys.platform == "win32":
                try:
                    import win32com.client
                    outlook_test = win32com.client.Dispatch("Outlook.Application")
                    mailer = "outlook"
                    print(f"[*] Tự động phát hiện phương thức: 'outlook' trên Windows - KHÔNG CẦN MẬT KHẨU!")
                except Exception:
                    mailer = "smtp"
            else:
                mailer = "smtp"

        success_count = 0
        failed_count = 0

        print(f"[*] Đang thực hiện gửi email bằng cơ chế: '{mailer}'...")

        for t in email_tasks:
            st = t["student"]
            plain_body, html_body = render_email_content(st, t["remark"])
            subject = f"[MathPlus Academy] Kết quả học tập & Nhận xét buổi học ngày {target_date} - Học sinh {st['name']}"

            try:
                if mailer == "macos-mail":
                    send_via_macos_mail(
                        to_email=t["email"],
                        subject=subject,
                        body_text=plain_body,
                        sender_email=sender_email,
                        open_draft_only=args.open_drafts
                    )
                    action_verb = "Đã tạo bản nháp trong Mail.app cho" if args.open_drafts else "Đã gửi email thành công tới"
                    print(f" [✅ {action_verb}] {st['name']} ({t['email']})")
                elif mailer == "outlook":
                    send_via_windows_outlook(
                        to_email=t["email"],
                        subject=subject,
                        body_text=plain_body,
                        html_text=html_body,
                        open_draft_only=args.open_drafts
                    )
                    action_verb = "Đã tạo bản nháp trong Outlook cho" if args.open_drafts else "Đã gửi email thành công qua Outlook tới"
                    print(f" [✅ {action_verb}] {st['name']} ({t['email']})")
                elif mailer == "service-account":
                    impersonate = args.impersonate_user or "cskh@mathplus.vn"
                    send_via_service_account(
                        creds_path=args.credentials or find_credentials_file(),
                        impersonate_user=impersonate,
                        to_email=t["email"],
                        student_name=st["name"],
                        session_date=target_date,
                        plain_text=plain_body,
                        html_text=html_body
                    )
                    print(f" [✅ ĐÃ GỬI qua Service Account] {st['name']} ({t['email']})")
                elif mailer == "smtp":
                    send_via_smtp(
                        smtp_cfg=smtp_cfg,

                        to_email=t["email"],
                        student_name=st["name"],
                        session_date=target_date,
                        plain_text=plain_body,
                        html_text=html_body
                    )
                    print(f" [✅ ĐÃ GỬI qua SMTP] {st['name']} ({t['email']})")

                success_count += 1
            except Exception as e:
                print(f" [❌ THẤT BẠI] Gửi email tới {st['name']} ({t['email']}) gặp lỗi: {e}", file=sys.stderr)
                failed_count += 1

        print(f"\n[+] Tổng kết: Thành công {success_count} / Thất bại {failed_count}")

if __name__ == "__main__":
    main()
