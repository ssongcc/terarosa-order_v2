"""
테라로사 자사몰 — 무료원두 쿠폰 치환 관리 페이지
pages/무료원두쿠폰.py 로 저장하세요.
"""

import json
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

CONFIG_PATH = Path("coupon_config.json")

# ──────────────────────────────────────────────
# 설정 로드/저장
# ──────────────────────────────────────────────
def load_config() -> list:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []

def save_config(cfg: list):
    CONFIG_PATH.write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
    )

# ──────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────
st.set_page_config(page_title="무료원두 쿠폰 관리", page_icon="🎁", layout="wide")

st.markdown("""
<style>
[data-testid="stSidebar"] { background: #FAF3F0; }
h1, h2, h3 { color: #8B3A2A !important; }
.stButton > button {
    background: #8B3A2A; color: white; border: none;
    border-radius: 6px; font-weight: 600;
}
.stButton > button:hover { background: #C4644A; color: white; }
.stDownloadButton > button {
    background: #8B3A2A; color: white; border: none;
    border-radius: 6px; font-weight: 600; width: 100%;
}
.info-box {
    background: #FAF3F0; border: 1px solid #EDE5DC;
    border-radius: 8px; padding: 12px 16px; margin-bottom: 16px;
    font-size: 13px; color: #2C2C2C;
}
</style>
""", unsafe_allow_html=True)

# 상태 초기화
if "coupon_items" not in st.session_state:
    st.session_state.coupon_items = load_config()

st.title("🎁 무료원두 쿠폰 치환 관리")
st.caption("'무료원두 쿠폰 250g' 행을 아래 설정한 품목들로 자동 교체합니다. 원본 행은 결과에서 사라집니다.")
st.divider()

left, right = st.columns([1, 1], gap="large")

# ════════════════════════
# 왼쪽: 치환 품목 설정
# ════════════════════════
with left:
    st.subheader("치환 품목 설정")
    st.markdown("""
<div class="info-box">
무료원두 쿠폰 250g → 아래 품목들로 교체됩니다.<br>
수량은 쿠폰 1개당 수량이 아니라 <b>직접 입력한 절대 수량</b>이 사용됩니다.
</div>
""", unsafe_allow_html=True)

    # 헤더
    hc1, hc2, hc3, hc4 = st.columns([4, 2, 2, 1])
    with hc1: st.caption("품목명 (A열)")
    with hc2: st.caption("중량 (B열)")
    with hc3: st.caption("수량 (D열)")
    with hc4: st.caption("")
    st.divider()

    items = st.session_state.coupon_items

    # 기존 품목 목록
    for i, item in enumerate(items):
        c1, c2, c3, c4 = st.columns([4, 2, 2, 1])
        with c1:
            new_name = st.text_input(
                "품목명", value=item.get("name", ""),
                key=f"name_{i}", label_visibility="collapsed",
                placeholder="예: 하우스 드립 블렌드"
            )
        with c2:
            new_weight = st.text_input(
                "중량", value=item.get("weight", "250g"),
                key=f"weight_{i}", label_visibility="collapsed",
                placeholder="예: 250g"
            )
        with c3:
            new_qty = st.number_input(
                "수량", min_value=1, value=item.get("qty", 1),
                key=f"qty_{i}", label_visibility="collapsed"
            )
        with c4:
            if st.button("✕", key=f"del_{i}"):
                items.pop(i)
                save_config(items)
                st.rerun()

        # 변경사항 실시간 반영
        items[i] = {
            "name": new_name.strip(),
            "weight": new_weight.strip(),
            "option": "증정 원두",
            "qty": int(new_qty)
        }

    # 새 품목 추가 행
    st.divider()
    na1, na2, na3, na4 = st.columns([4, 2, 2, 1])
    with na1:
        add_name = st.text_input(
            "품목명", placeholder="품목명 입력",
            key="add_name", label_visibility="collapsed"
        )
    with na2:
        add_weight = st.text_input(
            "중량", placeholder="예: 250g", value="250g",
            key="add_weight", label_visibility="collapsed"
        )
    with na3:
        add_qty = st.number_input(
            "수량", min_value=1, value=1,
            key="add_qty", label_visibility="collapsed"
        )
    with na4:
        if st.button("＋", key="add_btn"):
            if add_name.strip():
                items.append({
                    "name": add_name.strip(),
                    "weight": add_weight.strip(),
                    "option": "증정 원두",
                    "qty": int(add_qty)
                })
                save_config(items)
                st.rerun()

    st.divider()
    if st.button("💾 저장", use_container_width=True):
        save_config(items)
        st.success("저장 완료!")

# ════════════════════════
# 오른쪽: 미리보기
# ════════════════════════
with right:
    st.subheader("치환 결과 미리보기")

    if not st.session_state.coupon_items:
        st.info("← 왼쪽에서 치환 품목을 추가하세요.")
    else:
        st.markdown("**무료원두 쿠폰 250g** → 아래 품목으로 교체됩니다:")
        st.divider()

        preview_data = []
        total_qty = 0
        for item in st.session_state.coupon_items:
            if item.get("name"):
                preview_data.append({
                    "A열 (품목명)": item["name"],
                    "B열 (중량)": item.get("weight", "250g"),
                    "C열 (옵션)": item.get("option", "증정 원두"),
                    "D열 (수량)": item["qty"]
                })
                total_qty += item["qty"]

        if preview_data:
            st.dataframe(
                pd.DataFrame(preview_data),
                use_container_width=True,
                hide_index=True
            )
            st.caption(f"총 수량: {total_qty}개")

st.divider()

# ════════════════════════
# 쿠폰 치환 확인
# ════════════════════════
st.subheader("파일에서 쿠폰 확인")
st.caption("주문취합 파일을 업로드하면 무료원두 쿠폰 수량을 미리 확인할 수 있습니다.")

check_file = st.file_uploader("📄 주문취합 Excel 업로드", type=["xlsx"], key="check_file")
if check_file:
    try:
        df = pd.read_excel(check_file, sheet_name="취합용", header=0, dtype=str)
        df = df.iloc[:, :2].copy()
        df.columns = ["품목명", "수량"]
        mask = df["품목명"].str.contains("무료원두 쿠폰", na=False)
        coupon_rows = df[mask]

        if coupon_rows.empty:
            st.success("✅ 이 파일에는 무료원두 쿠폰이 없습니다.")
        else:
            total_coupon = pd.to_numeric(coupon_rows["수량"], errors="coerce").fillna(0).astype(int).sum()
            st.warning(f"🎁 무료원두 쿠폰 발견: **{total_coupon}개**")

            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**원본 쿠폰 행**")
                st.dataframe(coupon_rows, use_container_width=True, hide_index=True)
            with col_b:
                st.markdown("**치환 후 결과**")
                if st.session_state.coupon_items:
                    replaced = []
                    for item in st.session_state.coupon_items:
                        if item.get("name"):
                            replaced.append({
                                "품목명": item["name"],
                                "중량": item.get("weight", "250g"),
                                "옵션": item.get("option", "증정 원두"),
                                "수량": item["qty"]
                            })
                    st.dataframe(
                        pd.DataFrame(replaced),
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("치환 품목이 설정되지 않았습니다.")
    except Exception as e:
        st.error(f"파일 읽기 오류: {e}")

with st.expander("ℹ️ 사용 방법"):
    st.markdown("""
**설정 방법**
1. 왼쪽에서 치환할 품목명, 중량, 수량 입력 후 **＋** 클릭
2. **💾 저장** 클릭 → `coupon_config.json`에 자동 저장

**처리 결과**
- 주문취합 처리 시 `무료원두 쿠폰 250g` 행이 설정한 품목들로 자동 교체
- 원본 쿠폰 행은 결과에서 완전히 제거됨
- 수량은 설정값 그대로 사용 (쿠폰 수량과 무관)

**페이지 삭제**
- GitHub에서 `pages/무료원두쿠폰.py` 파일만 삭제하면 기능 비활성화
""")
