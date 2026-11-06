"""
Interactive Web Application Demo (FastAPI + Modern Dashboard UI)
Project: Student Burnout & Dropout Risk Prediction (Nhóm 03: Đỗ Quang Minh, Vũ Quang Huy)
Run with: python app.py
Then open: http://localhost:8000
"""

import os
import sys
import joblib
import pandas as pd
import uvicorn
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
from src.train_and_compare_models import DomainFeatureEngineer  # noqa: E402
setattr(sys.modules["__main__"], "DomainFeatureEngineer", DomainFeatureEngineer)

MODEL_FULL_PATH = os.path.join(BASE_DIR, "models", "best_dropout_model_full.pkl")
MODEL_EARLY_PATH = os.path.join(BASE_DIR, "models", "best_dropout_model_early_warning.pkl")
MODEL_BURNOUT_PATH = os.path.join(BASE_DIR, "models", "best_burnout_model.pkl")

app = FastAPI(
    title="Student Burnout & Dropout Risk Prediction - Nhóm 03 (SIC AI Capstone)",
    version="1.0.0",
)

model_full = joblib.load(MODEL_FULL_PATH) if os.path.exists(MODEL_FULL_PATH) else None
model_early = joblib.load(MODEL_EARLY_PATH) if os.path.exists(MODEL_EARLY_PATH) else None
model_burnout = joblib.load(MODEL_BURNOUT_PATH) if os.path.exists(MODEL_BURNOUT_PATH) else None


class StudentInput(BaseModel):
    Age: int = 21
    Gender: str = "Female"
    Year_of_Study: int = 2
    Department: str = "Engineering"
    Residence_Type: str = "Hostel"
    Attendance_Percent: float = 75.0
    Study_Hours_Per_Day: float = 3.0
    Previous_GPA: float = 6.8
    Backlogs: float = 1.0
    Sleep_Hours: float = 6.2
    Screen_Time_Hours: float = 5.5
    Exercise_Freq_Per_Week: int = 2
    Social_Activity_Score: float = 5.0
    Part_Time_Job: str = "No"
    Family_Income_Bracket: str = "Middle"
    Financial_Stress_Score: float = 5.0
    Family_Support_Score: float = 6.5
    Stress_Level: float = 5.5
    Anxiety_Score: float = 5.0
    Motivation_Score: float = 5.5
    Peer_Pressure_Score: float = 4.5
    Counseling_Access: str = "No"
    Burnout_Level: str = "Auto"


@app.post("/api/predict")
def predict_student(data: StudentInput):
    row = data.model_dump()
    burnout_mode = row.pop("Burnout_Level", "Auto")

    df_in = pd.DataFrame([row])

    # 1. Predict or use provided Burnout_Level
    burn_labels = {0: "Low", 1: "Medium", 2: "High"}
    if model_burnout is not None:
        burn_probs = model_burnout.predict_proba(df_in)[0]
        pred_burn_idx = int(burn_probs.argmax())
        pred_burn_label = burn_labels[pred_burn_idx]
        burn_prob_dict = {
            "Low": round(float(burn_probs[0]) * 100, 1),
            "Medium": round(float(burn_probs[1]) * 100, 1),
            "High": round(float(burn_probs[2]) * 100, 1),
        }
    else:
        pred_burn_label = "Medium"
        burn_prob_dict = {"Low": 30.0, "Medium": 45.0, "High": 25.0}

    effective_burnout = pred_burn_label if burnout_mode == "Auto" else burnout_mode
    df_in["Burnout_Level"] = effective_burnout

    # 2. Predict Dropout Risk (Full & Early Warning models)
    prob_full = float(model_full.predict_proba(df_in)[0][1]) if model_full is not None else 0.5
    df_early = df_in.drop(columns=["Burnout_Level"])
    prob_early = float(model_early.predict_proba(df_early)[0][1]) if model_early is not None else prob_full

    final_prob = prob_full if burnout_mode != "Auto" else (0.65 * prob_full + 0.35 * prob_early)
    dropout_pct = round(final_prob * 100, 1)
    dropout_label = "CÓ NGUY CƠ BỎ HỌC (YES)" if final_prob >= 0.5 else "AN TOÀN / THẤP (NO)"

    # 3. Compute engineered indices for explainability
    fe = DomainFeatureEngineer(include_burnout=True)
    fe.fit(df_in)
    df_fe = fe.transform(df_in)

    eng_idx = round(float(df_fe["Academic_Engagement_Index"].iloc[0]), 2)
    str_idx = round(float(df_fe["Psychological_Stress_Index"].iloc[0]), 2)
    ratio_idx = round(float(df_fe["Stress_Engagement_Ratio"].iloc[0]), 2)

    # 4. Generate recommendations
    recs = []
    if data.Attendance_Percent < 70:
        recs.append("Cảnh báo chuyên cần thấp (<70%): Cần cố vấn học tập liên hệ nhắc nhở và tìm hiểu nguyên nhân vắng học.")
    if data.Study_Hours_Per_Day < 2.0:
        recs.append("Thời lượng tự học thấp (<2h/ngày): Khuyến nghị tham gia nhóm học tập (Study Group) hoặc phụ đạo.")
    if data.Stress_Level >= 6.5 or data.Anxiety_Score >= 6.5:
        recs.append("Chỉ số Căng thẳng / Lo âu cao: Đề xuất kết nối Phòng Tham vấn Tâm lý học đường ngay lập tức.")
    if data.Financial_Stress_Score >= 6.5:
        recs.append("Áp lực tài chính lớn: Tư vấn học bổng hỗ trợ hoặc quỹ hỗ trợ sinh viên khó khăn.")
    if data.Screen_Time_Hours >= 7.5 or data.Sleep_Hours < 5.5:
        recs.append("Mất cân bằng sinh hoạt (Ngủ ít / Dùng thiết bị điện tử nhiều): Khuyến nghị điều chỉnh nhịp sinh học.")
    if not recs:
        recs.append("Hồ sơ học tập và tâm lý ổn định: Duy trì lộ trình học tập và hoạt động ngoại khóa hiện tại.")

    return {
        "dropout_probability": dropout_pct,
        "dropout_label": dropout_label,
        "is_high_risk": final_prob >= 0.5,
        "effective_burnout": effective_burnout,
        "predicted_burnout": pred_burn_label,
        "burnout_probabilities": burn_prob_dict,
        "indices": {
            "Academic_Engagement_Index": eng_idx,
            "Psychological_Stress_Index": str_idx,
            "Stress_Engagement_Ratio": ratio_idx,
        },
        "recommendations": recs,
    }


HTML_PAGE = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SIC AI Capstone - Dự đoán Nguy cơ Bỏ học & Burnout (Nhóm 03)</title>
<style>
  :root {
    --primary: #1d3557;
    --secondary: #457b9d;
    --teal: #2a9d8f;
    --amber: #f4a261;
    --danger: #e63946;
    --bg: #f4f7fb;
  }
  * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
  body { margin: 0; background: var(--bg); color: #1e293b; }
  header {
    background: linear-gradient(135deg, #1d3557 0%, #264653 100%);
    color: white; padding: 18px 32px;
    display: flex; justify-content: space-between; align-items: center;
    box-shadow: 0 3px 12px rgba(0,0,0,0.15);
  }
  header h1 { margin: 0; font-size: 21px; }
  header p { margin: 4px 0 0; font-size: 13px; opacity: 0.9; }
  .badge-team {
    background: rgba(255,255,255,0.15); padding: 8px 16px;
    border-radius: 8px; font-size: 13px; font-weight: 600;
  }
  .container {
    max-width: 1360px; margin: 20px auto; padding: 0 20px;
    display: grid; grid-template-columns: 1.35fr 1fr; gap: 22px;
  }
  .card {
    background: white; border-radius: 12px; padding: 20px 24px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.06);
  }
  .card h2 {
    margin: 0 0 14px; font-size: 16.5px; color: var(--primary);
    border-bottom: 2px solid #e2e8f0; padding-bottom: 8px;
  }
  .presets { display: flex; gap: 10px; margin-bottom: 16px; flex-wrap: wrap; }
  .btn-preset {
    border: none; padding: 8px 14px; border-radius: 6px; cursor: pointer;
    font-weight: 600; font-size: 12.5px; color: white; transition: 0.2s;
  }
  .btn-low { background: var(--teal); }
  .btn-med { background: #e07a5f; }
  .btn-high { background: var(--danger); }
  .btn-preset:hover { opacity: 0.9; transform: translateY(-1px); }
  .grid-form {
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px 16px;
  }
  .form-group label {
    display: flex; justify-content: space-between;
    font-size: 12.5px; font-weight: 600; margin-bottom: 4px; color: #334155;
  }
  .form-group input, .form-group select {
    width: 100%; padding: 7px 9px; border: 1px solid #cbd5e1;
    border-radius: 6px; font-size: 13px;
  }
  .btn-submit {
    margin-top: 18px; width: 100%; padding: 12px;
    background: var(--primary); color: white; border: none;
    border-radius: 8px; font-size: 15px; font-weight: 700; cursor: pointer;
  }
  .btn-submit:hover { background: #14253d; }
  .result-banner {
    text-align: center; padding: 20px; border-radius: 10px;
    color: white; margin-bottom: 18px;
  }
  .risk-prob { font-size: 42px; font-weight: 800; margin: 6px 0; }
  .kpi-row {
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 16px;
  }
  .kpi-box {
    background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;
    padding: 10px; text-align: center;
  }
  .kpi-val { font-size: 20px; font-weight: 700; color: var(--primary); }
  .kpi-lbl { font-size: 11.5px; color: #64748b; }
  .bar-wrap { margin-bottom: 10px; }
  .bar-label { display: flex; justify-content: space-between; font-size: 12.5px; font-weight: 600; margin-bottom: 3px; }
  .bar-bg { height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden; }
  .bar-fill { height: 100%; transition: width 0.4s; }
  ul.rec-list { padding-left: 18px; margin: 8px 0 0; font-size: 13.5px; line-height: 1.55; }
  ul.rec-list li { margin-bottom: 6px; }
</style>
</head>
<body>
<header>
  <div>
    <h1>🎓 HỆ THỐNG DỰ ĐOÁN NGUY CƠ BỎ HỌC & KIỆT SỨC HỌC ĐƯỜNG (STUDENT BURNOUT & DROPOUT AI)</h1>
    <p>Samsung Innovation Campus (SIC) AI Capstone Project — Mô hình Ensemble Learning trên các Chỉ số Căng thẳng & Tham gia Học tập</p>
  </div>
  <div class="badge-team">Nhóm 03: Đỗ Quang Minh • Vũ Quang Huy</div>
</header>

<div class="container">
  <div class="card">
    <h2>📋 Nhập Chỉ số Sinh viên (Student Engagement & Stress Indicators)</h2>
    <div class="presets">
      <button type="button" class="btn-preset btn-low" onclick="loadPreset('low')">🟢 Mẫu SV Tích cực (Nguy cơ Thấp)</button>
      <button type="button" class="btn-preset btn-med" onclick="loadPreset('med')">🟠 Mẫu SV Căng thẳng TB (Nguy cơ Vừa)</button>
      <button type="button" class="btn-preset btn-high" onclick="loadPreset('high')">🔴 Mẫu SV Kiệt sức (Nguy cơ Bỏ học Cao)</button>
    </div>

    <div class="grid-form">
      <div class="form-group">
        <label>Chuyên cần (Attendance %): <span id="v_att">75%</span></label>
        <input type="range" id="Attendance_Percent" min="30" max="100" step="1" value="75" oninput="syncVal('Attendance_Percent','v_att','%')">
      </div>
      <div class="form-group">
        <label>Giờ tự học/ngày (Study Hours): <span id="v_stu">3.0h</span></label>
        <input type="range" id="Study_Hours_Per_Day" min="0" max="8.5" step="0.1" value="3.0" oninput="syncVal('Study_Hours_Per_Day','v_stu','h')">
      </div>
      <div class="form-group">
        <label>Điểm động lực (Motivation 0-10): <span id="v_mot">5.5</span></label>
        <input type="range" id="Motivation_Score" min="0" max="10" step="0.1" value="5.5" oninput="syncVal('Motivation_Score','v_mot','')">
      </div>

      <div class="form-group">
        <label>GPA kỳ trước (Previous GPA): <span id="v_gpa">6.8</span></label>
        <input type="range" id="Previous_GPA" min="2.4" max="10.0" step="0.1" value="6.8" oninput="syncVal('Previous_GPA','v_gpa','')">
      </div>
      <div class="form-group">
        <label>Số môn nợ (Backlogs): <span id="v_bac">1</span></label>
        <input type="range" id="Backlogs" min="0" max="6" step="1" value="1" oninput="syncVal('Backlogs','v_bac','')">
      </div>
      <div class="form-group">
        <label>Mức căng thẳng (Stress 0-10): <span id="v_str">5.5</span></label>
        <input type="range" id="Stress_Level" min="0" max="10" step="0.1" value="5.5" oninput="syncVal('Stress_Level','v_str','')">
      </div>

      <div class="form-group">
        <label>Điểm lo âu (Anxiety 0-10): <span id="v_anx">5.0</span></label>
        <input type="range" id="Anxiety_Score" min="0" max="10" step="0.1" value="5.0" oninput="syncVal('Anxiety_Score','v_anx','')">
      </div>
      <div class="form-group">
        <label>Áp lực tài chính (Financial 0-10): <span id="v_fin">5.0</span></label>
        <input type="range" id="Financial_Stress_Score" min="0" max="10" step="0.1" value="5.0" oninput="syncVal('Financial_Stress_Score','v_fin','')">
      </div>
      <div class="form-group">
        <label>Áp lực bạn bè (Peer Pressure): <span id="v_pee">4.5</span></label>
        <input type="range" id="Peer_Pressure_Score" min="0" max="10" step="0.1" value="4.5" oninput="syncVal('Peer_Pressure_Score','v_pee','')">
      </div>

      <div class="form-group">
        <label>Hỗ trợ gia đình (Family Support): <span id="v_fam">6.5</span></label>
        <input type="range" id="Family_Support_Score" min="0" max="10" step="0.1" value="6.5" oninput="syncVal('Family_Support_Score','v_fam','')">
      </div>
      <div class="form-group">
        <label>Giờ ngủ/ngày (Sleep Hours): <span id="v_slp">6.2h</span></label>
        <input type="range" id="Sleep_Hours" min="2" max="10" step="0.1" value="6.2" oninput="syncVal('Sleep_Hours','v_slp','h')">
      </div>
      <div class="form-group">
        <label>Giờ màn hình (Screen Time): <span id="v_scr">5.5h</span></label>
        <input type="range" id="Screen_Time_Hours" min="0.5" max="12.5" step="0.1" value="5.5" oninput="syncVal('Screen_Time_Hours','v_scr','h')">
      </div>

      <div class="form-group">
        <label>Tập thể dục (buổi/tuần): <span id="v_exe">2</span></label>
        <input type="range" id="Exercise_Freq_Per_Week" min="0" max="7" step="1" value="2" oninput="syncVal('Exercise_Freq_Per_Week','v_exe','')">
      </div>
      <div class="form-group">
        <label>Hoạt động xã hội (Social 0-10): <span id="v_soc">5.0</span></label>
        <input type="range" id="Social_Activity_Score" min="0" max="10" step="0.1" value="5.0" oninput="syncVal('Social_Activity_Score','v_soc','')">
      </div>
      <div class="form-group">
        <label>Trạng thái Burnout_Level:</label>
        <select id="Burnout_Level">
          <option value="Auto">🤖 AI Tự động Dự báo (Auto)</option>
          <option value="Low">Low (Kiệt sức Thấp)</option>
          <option value="Medium">Medium (Kiệt sức Trung bình)</option>
          <option value="High">High (Kiệt sức Cao)</option>
        </select>
      </div>

      <div class="form-group">
        <label>Khoa / Ngành (Department):</label>
        <select id="Department">
          <option value="Engineering">Engineering</option>
          <option value="Business">Business</option>
          <option value="Science">Science</option>
          <option value="Arts">Arts</option>
          <option value="Medicine">Medicine</option>
          <option value="Law">Law</option>
        </select>
      </div>
      <div class="form-group">
        <label>Thu nhập GĐ (Income Bracket):</label>
        <select id="Family_Income_Bracket">
          <option value="Low">Low</option>
          <option value="Lower-Middle">Lower-Middle</option>
          <option value="Middle" selected>Middle</option>
          <option value="Upper-Middle">Upper-Middle</option>
          <option value="High">High</option>
        </select>
      </div>
      <div class="form-group">
        <label>Làm thêm (Part-Time Job):</label>
        <select id="Part_Time_Job">
          <option value="No">No (Không)</option>
          <option value="Yes">Yes (Có)</option>
        </select>
      </div>
    </div>

    <button class="btn-submit" onclick="runPrediction()">⚡ PHÂN TÍCH & DỰ ĐOÁN NGUY CƠ BỎ HỌC (RUN AI PREDICTION)</button>
  </div>

  <div class="card">
    <h2>📊 Kết quả Dự báo từ Hệ thống AI (AI Diagnostic Output)</h2>
    <div id="banner" class="result-banner" style="background: #2a9d8f;">
      <div style="font-size: 13px; font-weight: 600; opacity: 0.9;">XÁC SUẤT NGUY CƠ BỎ HỌC (DROPOUT RISK PROBABILITY)</div>
      <div id="out_prob" class="risk-prob">--%</div>
      <div id="out_label" style="font-size: 16px; font-weight: 700;">Đang tải dự báo...</div>
    </div>

    <div class="kpi-row">
      <div class="kpi-box">
        <div class="kpi-val" id="idx_burn">--</div>
        <div class="kpi-lbl">Mức độ Burnout (AI)</div>
      </div>
      <div class="kpi-box">
        <div class="kpi-val" id="idx_eng">--</div>
        <div class="kpi-lbl">Chỉ số Gắn kết Học tập</div>
      </div>
      <div class="kpi-box">
        <div class="kpi-val" id="idx_str">--</div>
        <div class="kpi-lbl">Chỉ số Căng thẳng Tâm lý</div>
      </div>
    </div>

    <h3 style="font-size: 14px; color: var(--primary); margin: 14px 0 8px;">Phân bố Xác suất Mức độ Kiệt sức (Burnout Level Probabilities)</h3>
    <div class="bar-wrap">
      <div class="bar-label"><span>Low Burnout (Thấp)</span><span id="p_low">0%</span></div>
      <div class="bar-bg"><div id="b_low" class="bar-fill" style="width:0%; background:#2a9d8f;"></div></div>
    </div>
    <div class="bar-wrap">
      <div class="bar-label"><span>Medium Burnout (Trung bình)</span><span id="p_med">0%</span></div>
      <div class="bar-bg"><div id="b_med" class="bar-fill" style="width:0%; background:#f4a261;"></div></div>
    </div>
    <div class="bar-wrap">
      <div class="bar-label"><span>High Burnout (Cao)</span><span id="p_high">0%</span></div>
      <div class="bar-bg"><div id="b_high" class="bar-fill" style="width:0%; background:#e63946;"></div></div>
    </div>

    <h3 style="font-size: 14px; color: var(--primary); margin: 16px 0 6px;">💡 Khuyến nghị Can thiệp Sớm (Personalized Recommendations)</h3>
    <ul id="rec_list" class="rec-list"></ul>
  </div>
</div>

<script>
function syncVal(id, spanId, unit) {
  document.getElementById(spanId).innerText = document.getElementById(id).value + unit;
}

function loadPreset(type) {
  const presets = {
    low: {Attendance_Percent: 93, Study_Hours_Per_Day: 4.8, Motivation_Score: 8.2, Previous_GPA: 8.1, Backlogs: 0, Stress_Level: 3.1, Anxiety_Score: 2.8, Financial_Stress_Score: 2.5, Peer_Pressure_Score: 3.0, Family_Support_Score: 8.5, Sleep_Hours: 7.5, Screen_Time_Hours: 3.8, Exercise_Freq_Per_Week: 4, Social_Activity_Score: 6.8, Burnout_Level: 'Low'},
    med: {Attendance_Percent: 74, Study_Hours_Per_Day: 2.8, Motivation_Score: 5.4, Previous_GPA: 6.5, Backlogs: 1, Stress_Level: 5.8, Anxiety_Score: 5.4, Financial_Stress_Score: 5.5, Peer_Pressure_Score: 4.8, Family_Support_Score: 6.0, Sleep_Hours: 6.0, Screen_Time_Hours: 6.0, Exercise_Freq_Per_Week: 2, Social_Activity_Score: 4.8, Burnout_Level: 'Medium'},
    high: {Attendance_Percent: 52, Study_Hours_Per_Day: 1.1, Motivation_Score: 2.8, Previous_GPA: 5.1, Backlogs: 3, Stress_Level: 8.6, Anxiety_Score: 8.2, Financial_Stress_Score: 7.9, Peer_Pressure_Score: 6.8, Family_Support_Score: 3.5, Sleep_Hours: 4.5, Screen_Time_Hours: 9.2, Exercise_Freq_Per_Week: 0, Social_Activity_Score: 2.5, Burnout_Level: 'High'}
  };
  const p = presets[type];
  for (const k in p) {
    const el = document.getElementById(k);
    if (el) el.value = p[k];
  }
  syncVal('Attendance_Percent','v_att','%');
  syncVal('Study_Hours_Per_Day','v_stu','h');
  syncVal('Motivation_Score','v_mot','');
  syncVal('Previous_GPA','v_gpa','');
  syncVal('Backlogs','v_bac','');
  syncVal('Stress_Level','v_str','');
  syncVal('Anxiety_Score','v_anx','');
  syncVal('Financial_Stress_Score','v_fin','');
  syncVal('Peer_Pressure_Score','v_pee','');
  syncVal('Family_Support_Score','v_fam','');
  syncVal('Sleep_Hours','v_slp','h');
  syncVal('Screen_Time_Hours','v_scr','h');
  syncVal('Exercise_Freq_Per_Week','v_exe','');
  syncVal('Social_Activity_Score','v_soc','');
  runPrediction();
}

async function runPrediction() {
  const payload = {
    Attendance_Percent: parseFloat(document.getElementById('Attendance_Percent').value),
    Study_Hours_Per_Day: parseFloat(document.getElementById('Study_Hours_Per_Day').value),
    Motivation_Score: parseFloat(document.getElementById('Motivation_Score').value),
    Previous_GPA: parseFloat(document.getElementById('Previous_GPA').value),
    Backlogs: parseFloat(document.getElementById('Backlogs').value),
    Stress_Level: parseFloat(document.getElementById('Stress_Level').value),
    Anxiety_Score: parseFloat(document.getElementById('Anxiety_Score').value),
    Financial_Stress_Score: parseFloat(document.getElementById('Financial_Stress_Score').value),
    Peer_Pressure_Score: parseFloat(document.getElementById('Peer_Pressure_Score').value),
    Family_Support_Score: parseFloat(document.getElementById('Family_Support_Score').value),
    Sleep_Hours: parseFloat(document.getElementById('Sleep_Hours').value),
    Screen_Time_Hours: parseFloat(document.getElementById('Screen_Time_Hours').value),
    Exercise_Freq_Per_Week: parseInt(document.getElementById('Exercise_Freq_Per_Week').value),
    Social_Activity_Score: parseFloat(document.getElementById('Social_Activity_Score').value),
    Burnout_Level: document.getElementById('Burnout_Level').value,
    Department: document.getElementById('Department').value,
    Family_Income_Bracket: document.getElementById('Family_Income_Bracket').value,
    Part_Time_Job: document.getElementById('Part_Time_Job').value
  };

  const res = await fetch('/api/predict', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload)
  });
  const data = await res.json();

  document.getElementById('out_prob').innerText = data.dropout_probability + '%';
  document.getElementById('out_label').innerText = data.dropout_label;
  const banner = document.getElementById('banner');
  if (data.dropout_probability >= 65) banner.style.background = '#e63946';
  else if (data.dropout_probability >= 45) banner.style.background = '#e07a5f';
  else banner.style.background = '#2a9d8f';

  document.getElementById('idx_burn').innerText = data.effective_burnout;
  document.getElementById('idx_eng').innerText = data.indices.Academic_Engagement_Index;
  document.getElementById('idx_str').innerText = data.indices.Psychological_Stress_Index;

  document.getElementById('p_low').innerText = data.burnout_probabilities.Low + '%';
  document.getElementById('b_low').style.width = data.burnout_probabilities.Low + '%';
  document.getElementById('p_med').innerText = data.burnout_probabilities.Medium + '%';
  document.getElementById('b_med').style.width = data.burnout_probabilities.Medium + '%';
  document.getElementById('p_high').innerText = data.burnout_probabilities.High + '%';
  document.getElementById('b_high').style.width = data.burnout_probabilities.High + '%';

  const ul = document.getElementById('rec_list');
  ul.innerHTML = '';
  data.recommendations.forEach(r => {
    const li = document.createElement('li');
    li.innerText = r;
    ul.appendChild(li);
  });
}

window.onload = () => runPrediction();
</script>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
def index():
    return HTML_PAGE


if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)
