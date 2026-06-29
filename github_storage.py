"""
GitHub 저장소를 영구 저장소로 사용하는 헬퍼 모듈
set_config.json, coupon_config.json, 상품코드 CSV를 GitHub에 저장/로드
"""
import json
import base64
import io
import streamlit as st
import pandas as pd

def _get_repo():
    try:
        from github import Github
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["GITHUB_REPO"]
        g = Github(token)
        return g.get_repo(repo_name)
    except Exception as e:
        return None

def gh_load(filename: str, default):
    """GitHub에서 JSON 파일 로드. 실패 시 default 반환"""
    repo = _get_repo()
    if repo is None:
        return default
    try:
        f = repo.get_contents(filename)
        return json.loads(base64.b64decode(f.content).decode("utf-8"))
    except Exception:
        return default

def gh_save(filename: str, data):
    """GitHub에 JSON 파일 저장 (없으면 생성, 있으면 업데이트)"""
    repo = _get_repo()
    if repo is None:
        return False
    content = json.dumps(data, ensure_ascii=False, indent=2)
    try:
        try:
            f = repo.get_contents(filename)
            repo.update_file(filename, f"Update {filename}", content, f.sha)
        except Exception:
            repo.create_file(filename, f"Create {filename}", content)
        return True
    except Exception as e:
        st.warning(f"GitHub 저장 실패: {e}")
        return False

def gh_save_df(filename: str, df: pd.DataFrame) -> bool:
    """DataFrame을 CSV로 GitHub에 저장"""
    repo = _get_repo()
    if repo is None:
        return False
    content = df.to_csv(index=False, encoding="utf-8")
    try:
        try:
            f = repo.get_contents(filename)
            repo.update_file(filename, f"Update {filename}", content, f.sha)
        except Exception:
            repo.create_file(filename, f"Create {filename}", content)
        return True
    except Exception as e:
        st.warning(f"GitHub 저장 실패: {e}")
        return False

def gh_load_df(filename: str) -> pd.DataFrame:
    """GitHub에서 CSV 파일을 DataFrame으로 로드"""
    repo = _get_repo()
    if repo is None:
        return pd.DataFrame()
    try:
        f = repo.get_contents(filename)
        content = base64.b64decode(f.content).decode("utf-8")
        return pd.read_csv(io.StringIO(content), dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()
