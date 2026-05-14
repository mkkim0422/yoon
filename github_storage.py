"""GitHub Contents API 를 통한 영구 저장소.

Streamlit Community Cloud 는 컨테이너 휴면 후 재기동 시 로컬 디스크가 초기화
되므로, 사용자가 업로드한 단가표 같은 파일은 휘발된다. 이 모듈은 업로드된
파일을 그대로 GitHub repo 에 커밋해 두고, 앱 시작 시 로컬에 없으면 repo 에서
받아오는 헬퍼를 제공한다.

설정 (`.streamlit/secrets.toml` 또는 Streamlit Cloud Secrets):

    [github]
    token  = "ghp_xxx"            # contents:write 권한 PAT
    repo   = "owner/repo-name"
    branch = "main"               # 선택, 기본 main

설정이 비어 있으면 모든 함수가 조용히 no-op 한다 (로컬 개발 환경에서도 안전).
"""
from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Optional
from urllib import error, request

import streamlit as st

_API = "https://api.github.com"


def _config() -> Optional[dict]:
    """st.secrets 에서 github 설정을 읽는다. 없으면 None."""
    try:
        cfg = st.secrets.get("github")  # type: ignore[attr-defined]
    except Exception:
        return None
    if not cfg:
        return None
    token = cfg.get("token")
    repo = cfg.get("repo")
    if not token or not repo:
        return None
    return {
        "token": token,
        "repo": repo,
        "branch": cfg.get("branch", "main"),
    }


def is_configured() -> bool:
    return _config() is not None


def _request(method: str, url: str, *, token: str, body: Optional[dict] = None) -> Optional[dict]:
    data = None if body is None else json.dumps(body).encode("utf-8")
    req = request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("X-GitHub-Api-Version", "2022-11-28")
    if body is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with request.urlopen(req, timeout=15) as resp:
            raw = resp.read()
            return json.loads(raw) if raw else {}
    except error.HTTPError as e:
        if e.code == 404:
            return None
        raise
    except error.URLError:
        return None


def _get_sha(remote_path: str, cfg: dict) -> Optional[str]:
    url = f"{_API}/repos/{cfg['repo']}/contents/{remote_path}?ref={cfg['branch']}"
    info = _request("GET", url, token=cfg["token"])
    if info is None:
        return None
    return info.get("sha")


def pull(remote_path: str, local_path: Path) -> bool:
    """Repo 의 remote_path 파일을 local_path 로 내려받는다.

    이미 local_path 가 존재하면 아무 것도 하지 않는다. 성공/스킵 모두 True,
    설정 없음·원격 없음·네트워크 실패 시 False.
    """
    if local_path.exists():
        return True
    cfg = _config()
    if cfg is None:
        return False
    url = f"{_API}/repos/{cfg['repo']}/contents/{remote_path}?ref={cfg['branch']}"
    try:
        info = _request("GET", url, token=cfg["token"])
    except Exception:
        return False
    if not info or "content" not in info:
        return False
    try:
        raw = base64.b64decode(info["content"])
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(raw)
        return True
    except Exception:
        return False


def push(local_path: Path, remote_path: str, message: str) -> bool:
    """local_path 의 내용을 repo 의 remote_path 에 커밋한다."""
    cfg = _config()
    if cfg is None or not local_path.exists():
        return False
    try:
        content_b64 = base64.b64encode(local_path.read_bytes()).decode("ascii")
        body = {
            "message": message,
            "content": content_b64,
            "branch": cfg["branch"],
        }
        sha = _get_sha(remote_path, cfg)
        if sha:
            body["sha"] = sha
        url = f"{_API}/repos/{cfg['repo']}/contents/{remote_path}"
        _request("PUT", url, token=cfg["token"], body=body)
        return True
    except Exception:
        return False


def delete(remote_path: str, message: str) -> bool:
    """repo 의 remote_path 파일을 삭제한다. 없으면 True 반환."""
    cfg = _config()
    if cfg is None:
        return False
    try:
        sha = _get_sha(remote_path, cfg)
        if sha is None:
            return True
        body = {"message": message, "sha": sha, "branch": cfg["branch"]}
        url = f"{_API}/repos/{cfg['repo']}/contents/{remote_path}"
        _request("DELETE", url, token=cfg["token"], body=body)
        return True
    except Exception:
        return False
