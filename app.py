from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import HTMLResponse, FileResponse
import subprocess
import os
import pandas as pd
import io

app = FastAPI()

# 1. 로고 이미지 라우터
@app.get("/logo.png")
async def get_logo():
    if os.path.exists("logo.png"):
        return FileResponse("logo.png")
    return HTMLResponse("")

# 2. 토스 풍 UI + 단가표 업로드 기능 추가
@app.get("/", response_class=HTMLResponse)
async def get_index():
    return """
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8">
        <title>SPH 정산서 자동화</title>
        <link rel="stylesheet" as="style" crossorigin href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.8/dist/web/static/pretendard.css" />
        <style>
            body { 
                font-family: 'Pretendard', sans-serif; 
                background-color: #f2f4f6; 
                display: flex; justify-content: center; align-items: center; 
                min-height: 100vh; margin: 0; padding: 20px;
            }
            .container {
                width: 100%; max-width: 500px;
            }
            .card { 
                background: white; 
                border-radius: 24px; 
                padding: 40px; 
                box-shadow: 0 8px 20px rgba(0,0,0,0.04); 
                box-sizing: border-box; margin-bottom: 20px;
            }
            .logo { height: 35px; margin-bottom: 24px; display: block; }
            h2 { 
                font-size: 24px; font-weight: 700; color: #191f28; 
                margin-top: 0; margin-bottom: 24px; letter-spacing: -0.5px; 
            }
            .sub-title { font-size: 18px; color: #333d4b; margin-bottom: 16px; font-weight: 600; }
            
            .form-group { margin-bottom: 24px; }
            label { 
                display: block; font-size: 15px; font-weight: 600; 
                color: #4e5968; margin-bottom: 10px; 
            }
            input[type="file"], select, input[type="number"] { 
                width: 100%; padding: 16px; font-size: 16px; 
                border: none; border-radius: 12px; 
                background-color: #f2f4f6; color: #191f28; 
                box-sizing: border-box; outline: none; transition: 0.2s; 
                font-family: 'Pretendard', sans-serif;
            }
            input[type="file"]:focus, select:focus, input[type="number"]:focus { 
                box-shadow: 0 0 0 2px #1A3C6B inset; 
            }
            select:disabled { color: #8b95a1; cursor: not-allowed; }
            
            /* 메인 액션 버튼 */
            .btn-primary { 
                width: 100%; padding: 18px; font-size: 17px; font-weight: 700; 
                color: white; background-color: #1A3C6B; 
                border: none; border-radius: 16px; cursor: pointer; transition: 0.2s; 
            }
            .btn-primary:hover { background-color: #122a4d; }
            .btn-primary:disabled { background-color: #d1d6db; cursor: not-allowed; }
            
            /* 서브 액션 버튼 (단가표용) */
            .btn-secondary {
                width: 100%; padding: 14px; font-size: 15px; font-weight: 600; 
                color: #1A3C6B; background-color: #e8f0fe; 
                border: none; border-radius: 12px; cursor: pointer; transition: 0.2s;
            }
            .btn-secondary:hover { background-color: #d2e3fc; }
            
            /* 구분선 */
            .divider { height: 1px; background-color: #e5e8eb; margin: 30px 0; }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="card" style="padding: 30px 40px;">
                <div class="sub-title">⚙️ 기준 단가표(템플릿) 설정</div>
                <p style="font-size: 13px; color: #8b95a1; margin-top: -10px; margin-bottom: 20px;">
                    한 번 업로드하면 수정하기 전까지 계속 유지됩니다.
                </p>
                <form id="templateForm">
                    <div class="form-group" style="margin-bottom: 16px;">
                        <input type="file" id="template_file" name="template_file" accept=".xlsx" required>
                    </div>
                    <button type="submit" class="btn-secondary" id="templateBtn">단가표 서버에 저장하기</button>
                </form>
            </div>

            <div class="card">
                <img src="/logo.png" alt="SPH Logo" class="logo" onerror="this.style.display='none'">
                <h2>월별 정산서 추출하기</h2>
                
                <form id="mainForm" action="/generate" method="post" enctype="multipart/form-data">
                    <div class="form-group">
                        <label>1. 구글 빌링 CSV 파일</label>
                        <input type="file" id="csv_file" name="csv_file" accept=".csv" required>
                    </div>
                    
                    <div class="form-group">
                        <label>2. 타겟 고객사</label>
                        <select id="project_name" name="project_name" disabled required>
                            <option value="">CSV 파일을 먼저 업로드해주세요</option>
                        </select>
                    </div>
                    
                    <div class="form-group">
                        <label>3. 이번 달 환율 (₩)</label>
                        <input type="number" step="0.01" name="fx_rate" value="1525.3" required>
                    </div>
                    
                    <button type="submit" class="btn-primary" id="submitBtn">정산서 엑셀 다운로드</button>
                </form>
            </div>
        </div>

        <script>
            // [기능 1] 단가표 업로드 비동기 처리
            document.getElementById('templateForm').addEventListener('submit', async function(e) {
                e.preventDefault(); // 페이지 새로고침 방지
                const formData = new FormData(this);
                const btn = document.getElementById('templateBtn');
                btn.textContent = '서버에 저장 중...';
                btn.disabled = true;

                try {
                    const response = await fetch('/upload-template', { method: 'POST', body: formData });
                    const data = await response.json();
                    if (data.error) {
                        alert('오류: ' + data.error);
                    } else {
                        alert('✅ ' + data.message);
                    }
                } catch (err) {
                    alert('업로드 중 통신 오류가 발생했습니다.');
                } finally {
                    btn.textContent = '단가표 서버에 저장하기';
                    btn.disabled = false;
                    document.getElementById('template_file').value = ''; // 입력창 초기화
                }
            });

            // [기능 2] CSV 스캔해서 고객사 목록 가져오기
            document.getElementById('csv_file').addEventListener('change', async function(e) {
                const file = e.target.files[0];
                if (!file) return;

                const select = document.getElementById('project_name');
                select.innerHTML = '<option value="">고객사 목록을 스캔하는 중...</option>';
                select.disabled = true;

                const formData = new FormData();
                formData.append('csv_file', file);

                try {
                    const response = await fetch('/get-accounts', { method: 'POST', body: formData });
                    const data = await response.json();

                    if (data.error) {
                        alert(data.error);
                        select.innerHTML = '<option value="">데이터를 불러오지 못했습니다</option>';
                    } else if (data.accounts && data.accounts.length > 0) {
                        select.innerHTML = '';
                        data.accounts.forEach(acc => {
                            const opt = document.createElement('option');
                            opt.value = acc;
                            opt.textContent = acc;
                            select.appendChild(opt);
                        });
                        select.disabled = false;
                    }
                } catch (err) {
                    alert('파일 분석 중 오류가 발생했습니다.');
                    select.innerHTML = '<option value="">오류 발생</option>';
                }
            });

            // [기능 3] 정산서 다운로드 버튼 로딩 처리
            document.getElementById('mainForm').addEventListener('submit', function() {
                const btn = document.getElementById('submitBtn');
                btn.textContent = '엑셀 생성 중... 잠시만 기다려주세요';
                btn.disabled = true;
                setTimeout(() => {
                    btn.textContent = '정산서 엑셀 다운로드';
                    btn.disabled = false;
                }, 4000);
            });
        </script>
    </body>
    </html>
    """

# [신규 API] 단가표(템플릿) 서버에 영구 저장하기
@app.post("/upload-template")
async def upload_template(template_file: UploadFile = File(...)):
    try:
        # 업로드된 파일을 "template.xlsx"라는 이름으로 서버(현재 폴더)에 덮어씁니다.
        with open("template.xlsx", "wb") as f:
            f.write(await template_file.read())
        return {"message": "단가표 템플릿이 성공적으로 업데이트되었습니다. (다음 정산부터 자동 적용)"}
    except Exception as e:
        return {"error": str(e)}

# [기존 API] 고객사 목록 스캔
@app.post("/get-accounts")
async def get_accounts(csv_file: UploadFile = File(...)):
    try:
        content = await csv_file.read()
        try:
            df = pd.read_csv(io.BytesIO(content), skiprows=8, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(content), skiprows=8, encoding='cp949')

        if '결제 계정 이름' in df.columns:
            accounts = df['결제 계정 이름'].dropna().unique().tolist()
            return {"accounts": sorted([str(a).strip() for a in accounts if str(a).strip()])}
        else:
            return {"error": "CSV 파일에 '결제 계정 이름' 열이 없습니다. 올바른 청구서인지 확인해주세요."}
    except Exception as e:
        return {"error": str(e)}

# [기존 API] 엑셀 생성 로직
@app.post("/generate")
async def generate_excel(csv_file: UploadFile = File(...), project_name: str = Form(...), fx_rate: float = Form(...)):
    # 1. 사용자가 올린 CSV는 1회용 임시 파일로 저장
    temp_csv_path = f"temp_{csv_file.filename}"
    with open(temp_csv_path, "wb") as f:
        f.write(await csv_file.read())
    
    out_excel_path = f"Invoice_Result_{project_name}.xlsx"
    
    # 2. 서버에 저장되어 있는 가장 최신 단가표(template.xlsx)를 무기로 사용!
    template_path = "template.xlsx"
    if not os.path.exists(template_path):
        return {"error": "서버에 단가표 템플릿이 없습니다. 상단의 '단가표 설정'에서 먼저 업로드해주세요!"}
    
    command = [
        "python", "dynamic_invoice.py", 
        template_path, temp_csv_path, out_excel_path, 
        "--project", project_name, 
        "--fx", str(fx_rate)
    ]
    
    try:
        subprocess.run(command, check=True)
    except Exception as e:
        return {"error": f"엑셀 생성 중 오류가 발생했습니다. 원인: {str(e)}"}
        
    return FileResponse(
        out_excel_path, 
        filename=f"정산서_{project_name}.xlsx", 
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )