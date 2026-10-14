from flask import Flask, render_template_string, request, send_file, redirect, url_for
import io, json
from engine import make_excel_report

app = Flask(__name__)

# Весь шаблон сайта в одной переменной
FULL_HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <title>СоциоКласс РБ Pro</title>
    <meta charset="utf-8">
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 30px; background-color: #f8fafc; color: #1e293b; line-height: 1.5; }
        .container { max-width: 1100px; margin: 0 auto; }
        .header-panel { background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); color: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); margin-bottom: 25px; }
        .header-panel h2 { margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.5px; }
        .header-panel p { margin: 10px 0 0 0; opacity: 0.9; font-size: 15px; }
        h3 { color: #0f172a; font-size: 18px; font-weight: 600; margin-top: 30px; margin-bottom: 15px; padding-left: 5px; }
        .card { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; margin-bottom: 20px; }
        .param-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 20px; }
        label { font-weight: 600; font-size: 13px; color: #475569; display: block; margin-bottom: 6px; text-transform: uppercase; letter-spacing: 0.5px; }
        input, select { width: 100%; padding: 10px 12px; border: 1px solid #cbd5e1; border-radius: 8px; box-sizing: border-box; font-size: 14px; color: #0f172a; background-color: #fff; }
        input:focus, select:focus { border-color: #3b82f6; outline: none; box-shadow: 0 0 0 3px rgba(59,130,246,0.15); }
        .btn { display: inline-flex; align-items: center; justify-content: center; background-color: #10b981; color: white; padding: 12px 24px; border: none; border-radius: 8px; cursor: pointer; font-size: 15px; font-weight: 600; box-shadow: 0 1px 2px rgba(0,0,0,0.05); transition: background-color 0.2s; }
        .btn:hover { background-color: #059669; }
        .btn-panel { background-color: #1e3a8a; width: 100%; margin-top: 10px; }
        .btn-panel:hover { background-color: #172554; }
        .student-block { background: white; padding: 25px; margin-top: 20px; border-radius: 12px; border: 1px solid #e2e8f0; border-left: 6px solid #3b82f6; transition: border-color 0.3s; }
        .student-title { font-size: 16px; font-weight: 600; color: #1e3a8a; margin-bottom: 15px; padding-bottom: 8px; border-bottom: 1px solid #f1f5f9; display: flex; align-items: center; }
        .student-title::before { content: ""; display: inline-block; width: 4px; height: 16px; background: #3b82f6; border-radius: 2px; margin-right: 8px; }
        .tile-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(115px, 1fr)); gap: 12px; }
        .tile-card { border: 1px solid #e2e8f0; border-radius: 10px; padding: 10px; text-align: center; background: #f8fafc; display: flex; flex-direction: column; align-items: center; justify-content: center; }
        .tile-card.active-p { background-color: #ecfdf5; border-color: #10b981; color: #065f46; box-shadow: 0 0 0 1px #10b981; }
        .tile-card.active-m { background-color: #fef2f2; border-color: #ef4444; color: #991b1b; box-shadow: 0 0 0 1px #ef4444; }
        .tile-card.self-block { background-color: #f1f5f9; border-color: #e2e8f0; color: #94a3b8; opacity: 0.6; pointer-events: none; }
        .tile-num { font-weight: 700; font-size: 15px; margin-bottom: 8px; }
        .sub-btn-grid { display: flex; gap: 6px; width: 100%; }
        .btn-sub { flex: 1; height: 28px; border: 1px solid #cbd5e1; background: #fff; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; color: #475569; }
        .btn-sub:hover { background: #f1f5f9; color: #0f172a; border-color: #94a3b8; }
        .info-bar { color: #64748b; font-style: italic; font-size: 13px; margin-bottom: 15px; display: flex; align-items: center; }
        .info-bar::before { content: "ℹ️"; margin-right: 6px; }
        .bot-bar { display: flex; justify-content: space-between; margin-top: 35px; padding-top: 20px; border-top: 2px solid #f1f5f9; gap: 15px; }
    </style>
    <script>
        function toggleChoice(studentFrom, studentTo, type) {
            let card = document.getElementById(`card_${studentFrom}_${studentTo}`);
            let pInput = document.getElementById(`in_p_${studentFrom}_${studentTo}`);
            let mInput = document.getElementById(`in_m_${studentFrom}_${studentTo}`);
            let limP = parseInt(document.getElementById('limit_p').value) || 0;
            let limM = parseInt(document.getElementById('limit_m').value) || 0;
            
            if (type === 'plus') {
                if (!card.classList.contains('active-p')) {
                    if (limP > 0) {
                        let activePluses = document.querySelectorAll(`[id^="in_p_${studentFrom}_"][value="1"]`).length;
                        if (activePluses >= limP) {
                            alert(`Превышен лимит! Не более ${limP} выборов (+) для одного ученика.`);
                            return;
                        }
                    }
                    card.classList.remove('active-m'); card.classList.add('active-p');
                    pInput.value = "1"; mInput.value = "0";
                } else {
                    card.classList.remove('active-p'); pInput.value = "0";
                }
            } else {
                if (!card.classList.contains('active-m')) {
                    if (limM > 0) {
                        let activeMinuses = document.querySelectorAll(`[id^="in_m_${studentFrom}_"][value="1"]`).length;
                        if (activeMinuses >= limM) {
                            alert(`Превышен лимит! Не более ${limM} отклонений (-) для одного ученика.`);
                            return;
                        }
                    }
                    card.classList.remove('active-p'); card.classList.add('active-m');
                    mInput.value = "1"; pInput.value = "0";
                } else {
                    card.classList.remove('active-m'); mInput.value = "0";
                }
            }
        }
        
        function clearAllData() {
            if (confirm("Вы уверены, что хотите полностью сбросить всю матрицу класса?")) {
                document.querySelectorAll('.tile-card').forEach(c => {
                    if(!c.classList.contains('self-block')) c.classList.remove('active-p', 'active-m');
                });
                document.querySelectorAll('input[id^="in_"]').forEach(i => i.value = "0");
            }
        }

        function validateForm() {
            let totalStudents = parseInt(document.getElementById('student_count_val').value) || 0;
            let unfilled = [];
            document.querySelectorAll('.student-block').forEach(b => b.style.borderColor = '#e2e8f0');

            for (let i = 1; i <= totalStudents; i++) {
                let hasPlus = document.querySelectorAll(`[id^="in_p_${i}_"][value="1"]`).length > 0;
                let hasMinus = document.querySelectorAll(`[id^="in_m_${i}_"][value="1"]`).length > 0;
                if (!hasPlus && !hasMinus) {
                    unfilled.push(i);
                    let block = document.getElementById(`block_student_${i}`);
                    if (block) block.style.borderColor = '#ef4444';
                }
            }

            if (unfilled.length > 0) {
                let msg = `Внимание! Ученики под номерами: ${unfilled.join(', ')} остались незаполненными.\\n\\n`;
                msg += `Вы уверены, что хотите сформировать и скачать отчет с пустыми строками?`;
                return confirm(msg);
            }
            return true;
        }
    </script>
</head>
<body>
<div class="container">
    <div class="header-panel">
        <h2>🧠 СоциоКласс РБ Pro</h2>
        <p>Профессиональный автоматический расчет школьной социометрии Морено-Коломинского без сбора персональных данных</p>
    </div>
    
    <form method="POST" onsubmit="if(event.submitter && event.submitter.value === 'export') { return validateForm(); }">
        <div class="card">
            <div class="param-grid">
                <div>
                    <label>Страна отчета</label>
                    <select name="country">
                        <option value="РБ" {% if country=='РБ' %}selected{% endif %}>Республика Беларусь</option>
                        <option value="РФ" {% if country=='РФ' %}selected{% endif %}>Российская Федерация</option>
                    </select>
                </div>
                <div>
                    <label>Название учреждения</label>
                    <input type="text" name="school" value="{{ school }}" required>
                </div>
                <div>
                    <label>Класс</label>
                    <input type="text" name="cls" value="{{ cls }}" required>
                </div>
                <div>
                    <label>ФИО психолога</label>
                    <input type="text" name="expert" value="{{ expert }}" required>
                </div>
            </div>
        </div>
        
        <div class="card" style="border-top: 3px solid #3b82f6;">
            <div class="param-grid">
                <div>
                    <label>Учеников в классе (2-40)</label>
                    <input type="number" id="student_count_val" name="count" value="{{ count }}" min="2" max="40" required>
                </div>
                <div>
                    <label>Лимит выборов (+)</label>
                    <input type="number" id="limit_p" name="limit_p" value="{{ limit_p }}">
                </div>
                <div>
                    <label>Лимит отклонений (-)</label>
                    <input type="number" id="limit_m" name="limit_m" value="{{ limit_m }}">
                </div>
                <div>
                    <label>Порог «Звезд» (>=)</label>
                    <input type="number" name="thresh_star" value="{{ thresh_star }}">
                </div>
                <div>
                    <label>Порог «Отвергаемых» (<=)</label>
                    <input type="number" name="thresh_rej" value="{{ thresh_rej }}">
                </div>
            </div>
            
            <div style="margin-top: 20px; display: flex; align-items: center;">
                <input type="checkbox" name="calc_indices" value="1" id="calc_ind" {% if calc_indices %}checked{% endif %} style="width:auto; margin-right:10px; cursor:pointer;">
                <label for="calc_ind" style="display:inline; font-weight:500; color:#334155; text-transform:none; cursor:pointer; font-size:14px;">Рассчитать и вывести групповые индексы (ИБВ, ИГС)</label>
            </div>
            <button type="submit" name="action" value="set_count" class="btn btn-panel">🔄 Создать / Обновить список класса</button>
        </div>
        {% if count_blocks > 0 %}
        <h3>📊 Интерактивная матрица выборов класса</h3>
        <p class="info-bar">Нажимайте кнопки [+] и [-] внутри плиток одноклассников. Все персональные данные полностью защищены!</p>
        
        {% for i in range(1, count_blocks + 1) %}
        <div class="student-block" id="block_student_{{ i }}">
            <div class="student-title">Ученик №{{ i }} оценивает коллектив:</div>
            <div class="tile-grid">
                {% for j in range(1, count_blocks + 1) %}
                    {% set is_active_p = j in saved_data.get(i, {}).get('plus', []) %}
                    {% set is_active_m = j in saved_data.get(i, {}).get('minus', []) %}
                    
                    {% if i == j %}
                    <div class="tile-card self-block" id="card_{{i}}_{{j}}">
                        <div class="tile-num">№{{ j }}</div>
                        <div style="font-size:11px; font-style:italic;">Сам себя</div>
                    </div>
                    {% else %}
                    <div class="tile-card {% if is_active_p %}active-p{% elif is_active_m %}active-m{% endif %}" id="card_{{i}}_{{j}}">
                        <div class="tile-num">№{{ j }}</div>
                        <div class="sub-btn-grid">
                            <button type="button" class="btn-sub" onclick="toggleChoice({{i}}, {{j}}, 'plus')">+</button>
                            <button type="button" class="btn-sub" onclick="toggleChoice({{i}}, {{j}}, 'minus')">-</button>
                        </div>
                        <input type="hidden" id="in_p_{{i}}_{{j}}" name="in_p_{{i}}_{{j}}" value="{% if is_active_p %}1{% else %}0{% endif %}">
                        <input type="hidden" id="in_m_{{i}}_{{j}}" name="in_m_{{i}}_{{j}}" value="{% if is_active_m %}1{% else %}0{% endif %}">
                    </div>
                    {% endif %}
                {% endfor %}
            </div>
        </div>
        {% endfor %}
        
        <!-- ИСПРАВЛЕНО: Добавлено JavaScript-предупреждение на кнопку «Начальный экран» -->
        <div class="bot-bar">
            <div>
                <button type="button" class="btn" style="background-color:#ef4444; margin-right: 10px;" onclick="clearAllData()">🗑️ Очистить матрицу</button>
                <button type="submit" name="action" value="reset_home" class="btn" style="background-color:#3b82f6;" onclick="return confirm('Вы уверены, что хотите вернуться на начальный экран?\\n\\nВсе текущие внесенные данные этого класса будут безвозвратно удалены!')">🏠 Начальный экран</button>
            </div>
            <button type="submit" name="action" value="export" class="btn">🟢 Сформировать и скачать отчет Excel</button>
        </div>
        {% endif %}
    </form>
</div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    school, cls, expert, count = "Школа №1", "10-А", "И.И. Иванов", "10"
    limit_p, limit_m, thresh_star, thresh_rej = "3", "3", "3", "0"
    country = "РБ"
    calc_indices = True
    count_blocks = 0
    saved_data = {}

    if request.method == "POST":
        action = request.form.get("action")
        
        if action == "reset_home":
            return redirect(url_for('index'))
            
        country = request.form.get("country")
        school = request.form.get("school")
        cls = request.form.get("cls")
        expert = request.form.get("expert")
        count = request.form.get("count")
        limit_p = request.form.get("limit_p", "0")
        limit_m = request.form.get("limit_m", "0")
        thresh_star = request.form.get("thresh_star", "3")
        thresh_rej = request.form.get("thresh_rej", "0")
        calc_indices = True if request.form.get("calc_indices") else False
        
        c_int = int(count)
        count_blocks = c_int
        
        for i in range(1, c_int + 1):
            saved_data[i] = {
                'plus': [j for j in range(1, c_int + 1) if request.form.get(f"in_p_{i}_{j}") == "1"],
                'minus': [j for j in range(1, c_int + 1) if request.form.get(f"in_m_{i}_{j}") == "1"]
            }

        if action == "export":
            ts = int(thresh_star) if thresh_star.isdigit() else 3
            tr = int(thresh_rej) if thresh_rej.isdigit() else 0
            memory_file = io.BytesIO()
            make_excel_report(school, cls, expert, saved_data, memory_file, ts, tr, country, calc_indices)
            memory_file.seek(0)
            return send_file(memory_file, download_name=f"Sociometry_{cls}.xlsx", as_attachment=True, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return render_template_string(FULL_HTML_TEMPLATE, school=school, cls=cls, expert=expert, count=count, count_blocks=count_blocks, limit_p=limit_p, limit_m=limit_m, thresh_star=thresh_star, thresh_rej=thresh_rej, country=country, calc_indices=calc_indices, saved_data=saved_data)

if __name__ == "__main__":
    app.run(debug=True)
