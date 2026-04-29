# deploy/ — nginx 리버스 프록시 패키지

Streamlit 앱(`webapp.py`)을 nginx 뒤에 두고 **외부 포트 8084**로 노출합니다.

```
브라우저  ──▶  nginx (0.0.0.0:8084)  ──▶  streamlit (127.0.0.1:8501)
```

## 구성 파일

| 경로                       | 역할                                                 |
|---------------------------|------------------------------------------------------|
| `nginx/nginx.exe`         | Windows용 nginx 1.26.2 (mainline stable)             |
| `nginx/conf/nginx.conf`   | 8084→8501 프록시 + WebSocket 업그레이드 설정          |
| `start.bat`               | streamlit + nginx 한 번에 기동                        |
| `stop.bat`                | nginx graceful quit + streamlit 프로세스 종료         |

## 로컬 실행

```
deploy\start.bat        # 기동
deploy\stop.bat         # 종료
```

기동 후 브라우저에서 **http://localhost:8084** 접속.

- streamlit은 별도 콘솔(제목: `yoon-streamlit`)에서 동작 — 로그 확인 가능.
- nginx는 백그라운드. 로그는 `deploy\nginx\logs\access.log`, `error.log`.

## 다른 서버로 옮길 때

`deploy/` 폴더는 자족 패키지라 그대로 복사하면 됩니다. 새 서버에서 한 번만 해야 할 일:

1. **프로젝트 전체 복사**: 프로젝트 루트(`webapp.py`, `billing/`, `requirements.txt` 등 포함)를 통째로 복사.
2. **Python 가상환경 재생성** — `.venv`는 OS/경로에 종속되어 복사하면 안 됨:
   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
3. **방화벽**: 외부에서 접속하려면 inbound TCP 8084 허용.
   ```
   netsh advfirewall firewall add rule name="yoon-8084" dir=in action=allow protocol=TCP localport=8084
   ```
4. `deploy\start.bat` 실행.

## 포트 / 설정 변경

- **외부 포트 변경**: `nginx/conf/nginx.conf`의 `listen 8084;` 수정 후 `nginx -s reload`.
- **streamlit 포트 변경**: `start.bat`의 `--server.port`와 `nginx.conf`의 `upstream streamlit_backend { server 127.0.0.1:8501; }` 둘 다 수정.
- **업로드 크기**: 기본 200MB. `nginx.conf`의 `client_max_body_size` 수정.

## nginx 직접 제어 (디버깅용)

```
cd deploy\nginx
nginx -t              # 설정 문법 검사
nginx                 # 기동
nginx -s reload       # 설정 리로드
nginx -s quit         # graceful 종료
nginx -s stop         # 즉시 종료
```
