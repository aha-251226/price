import streamlit as st
import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
import io
from datetime import datetime
import logging
import re
import os

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

st.set_page_config(page_title="네이버 부동산 매물 수집기", layout="centered")

st.title("🏠 네이버 부동산 매물 수집기")
st.markdown("**📌 지역 이름을 입력하면 자동으로 법정동 코드를 찾아 매물 정보를 수집하고 엑셀로 저장합니다.**")

# 세션 상태 초기화
if 'cortarNo' not in st.session_state:
    st.session_state.cortarNo = ""
if 'search_results' not in st.session_state:
    st.session_state.search_results = []

# 법정동코드 파일 자동 생성 함수
def create_sample_legal_code():
    sample_data = [
        ["1168010100", "서울특별시", "강남구", "역삼동"],
        ["1168010200", "서울특별시", "강남구", "논현동"],
        ["1168010300", "서울특별시", "강남구", "압구정동"],
        ["1168010400", "서울특별시", "강남구", "신사동"],
        ["1168010500", "서울특별시", "강남구", "청담동"],
        ["1168010600", "서울특별시", "강남구", "삼성동"],
        ["1165010100", "서울특별시", "서초구", "서초동"],
        ["1165010200", "서울특별시", "서초구", "잠원동"],
        ["1165010300", "서울특별시", "서초구", "반포동"],
        ["1111013100", "서울특별시", "종로구", "종로1가"],
    ]
    
    df = pd.DataFrame(sample_data, columns=['법정동코드', '시도', '시군구', '읍면동'])
    df.to_csv('법정동코드.csv', index=False, encoding='utf-8-sig')
    return df

# 법정동코드 파일 확인 및 생성
if not os.path.exists('법정동코드.csv'):
    st.warning("법정동코드.csv 파일이 없습니다. 샘플 파일을 자동 생성합니다.")
    create_sample_legal_code()
    st.success("샘플 법정동코드.csv 파일이 생성되었습니다!")

# --- 1. 법정동 코드 자동 검색기 ---
st.subheader("1️⃣ 지역 정보 입력")

col1, col2, col3 = st.columns(3)
with col1:
    sido = st.text_input("시/도", placeholder="예: 서울특별시")
with col2:
    sigungu = st.text_input("시/군/구", placeholder="예: 강남구")
with col3:
    eupmyeondong = st.text_input("읍/면/동", placeholder="예: 삼성동")

def search_legal_code(sido, sigungu, eupmyeondong):
    """법정동 코드를 검색하는 함수"""
    try:
        # 여러 인코딩으로 시도
        encodings = ['utf-8', 'utf-8-sig', 'cp949', 'euc-kr']
        law_df = None
        
        for encoding in encodings:
            try:
                law_df = pd.read_csv("법정동코드.csv", dtype=str, encoding=encoding)
                break
            except UnicodeDecodeError:
                continue
        
        if law_df is None:
            return "", False, "❌ 법정동코드.csv 파일 읽기 실패 (인코딩 문제)"
        
        # 컬럼명 확인 및 표준화
        columns = law_df.columns.tolist()
        st.write(f"🔍 CSV 파일 컬럼: {columns}")  # 디버깅용
        
        # 다양한 컬럼명 패턴 지원
        column_mapping = {}
        for col in columns:
            col_lower = col.strip().lower()
            if '법정동' in col or 'code' in col_lower or '코드' in col:
                column_mapping['법정동코드'] = col
            elif '시도' in col or 'sido' in col_lower:
                column_mapping['시도'] = col
            elif '시군구' in col or 'sigungu' in col_lower or '구군' in col:
                column_mapping['시군구'] = col
            elif '읍면동' in col or 'dong' in col_lower or '동' in col:
                column_mapping['읍면동'] = col
        
        # 필수 컬럼 확인
        required_cols = ['법정동코드', '시도', '시군구', '읍면동']
        missing_cols = [col for col in required_cols if col not in column_mapping]
        
        if missing_cols:
            return "", False, f"❌ CSV 파일에 필요한 컬럼이 없습니다: {missing_cols}. 현재 컬럼: {columns}"
        
        # 컬럼명 매핑 적용
        if column_mapping:
            law_df = law_df.rename(columns={v: k for k, v in column_mapping.items()})
        
        # 공백 제거 및 정규화
        for col in ['시도', '시군구', '읍면동']:
            if col in law_df.columns:
                law_df[col] = law_df[col].astype(str).str.strip()
        
        # 검색 실행
        sido_clean = sido.strip()
        sigungu_clean = sigungu.strip()
        eupmyeondong_clean = eupmyeondong.strip()
        
        # 정확한 매칭
        exact_match = law_df[
            (law_df["시도"] == sido_clean) &
            (law_df["시군구"] == sigungu_clean) &
            (law_df["읍면동"] == eupmyeondong_clean)
        ]
        
        if not exact_match.empty:
            return exact_match.iloc[0]["법정동코드"], True, "✅ 법정동 코드를 찾았습니다!"
        
        # 부분 매칭
        partial_match = law_df[
            (law_df["시도"].str.contains(sido_clean, na=False)) &
            (law_df["시군구"].str.contains(sigungu_clean, na=False)) &
            (law_df["읍면동"].str.contains(eupmyeondong_clean, na=False))
        ]
        
        if not partial_match.empty:
            return partial_match.iloc[0]["법정동코드"], True, "✅ 부분 일치로 법정동 코드를 찾았습니다!"
        
        # 시도, 시군구만 매칭하여 사용 가능한 동 표시
        area_match = law_df[
            (law_df["시도"].str.contains(sido_clean, na=False)) &
            (law_df["시군구"].str.contains(sigungu_clean, na=False))
        ]
        
        if not area_match.empty:
            available_dongs = area_match["읍면동"].unique()[:10]
            return "", False, f"❗ '{eupmyeondong_clean}'을 찾을 수 없습니다. 사용 가능한 동: {', '.join(available_dongs)}"
        else:
            return "", False, f"❗ '{sido_clean} {sigungu_clean}'를 찾을 수 없습니다."
                
    except FileNotFoundError:
        return "", False, "❌ '법정동코드.csv' 파일이 없습니다."
    except Exception as e:
        logger.error(f"법정동 코드 검색 오류: {e}")
        return "", False, f"❌ 오류가 발생했습니다: {str(e)}"

if st.button("🔍 법정동 코드 자동 검색"):
    if sido and sigungu and eupmyeondong:
        cortarNo, success, message = search_legal_code(sido, sigungu, eupmyeondong)
        if success:
            st.session_state.cortarNo = cortarNo
            st.success(f"{message} (코드: {cortarNo})")
        else:
            st.session_state.cortarNo = ""
            st.warning(message)
    else:
        st.warning("모든 지역 정보를 입력해주세요.")

# --- 2. 매물 유형 및 거래 방식 선택 ---
st.subheader("2️⃣ 검색 조건 설정")

col4, col5 = st.columns(2)
with col4:
    # 올바른 네이버 부동산 매물 유형 코드 사용
    property_types = {
        "APT": "아파트",
        "OPST": "오피스텔", 
        "VL": "빌라/연립/다세대",
        "ABYG": "아파트분양권",
        "OBYG": "오피스텔분양권",
        "SG": "상가",
        "SMS": "사무실",
        "GJCG": "공장/창고",
        "TJ": "토지",
        "JGC": "재개발/재건축"
    }
    
    rletTpCd = st.selectbox("매물 유형", options=list(property_types.keys()), format_func=lambda x: property_types[x])
with col5:
    tradTpCd = st.selectbox("거래 유형", options=["A1", "B1", "B2"], format_func=lambda x: {"A1": "매매", "B1": "전세", "B2": "월세"}[x])

# 면적 조건 설정
st.subheader("📐 면적 조건 (선택사항)")
col6, col7 = st.columns(2)
with col6:
    area_filter_enabled = st.checkbox("면적 조건 사용", value=False)
    min_area = st.number_input("최소 면적 (㎡)", min_value=0, max_value=10000, value=0, step=10, disabled=not area_filter_enabled)
with col7:
    st.write("")  # 공간 확보
    max_area = st.number_input("최대 면적 (㎡)", min_value=0, max_value=10000, value=1000, step=10, disabled=not area_filter_enabled)

# 가격 조건 설정
st.subheader("💰 가격 조건 (선택사항)")
col8, col9 = st.columns(2)
with col8:
    price_filter_enabled = st.checkbox("가격 조건 사용", value=False)
    min_price = st.number_input("최소 가격 (만원)", min_value=0, max_value=1000000, value=0, step=1000, disabled=not price_filter_enabled)
with col9:
    st.write("")  # 공간 확보
    max_price = st.number_input("최대 가격 (만원)", min_value=0, max_value=1000000, value=100000, step=1000, disabled=not price_filter_enabled)

# 추가 검색 옵션
st.subheader("⚙️ 고급 설정 (선택사항)")
col10, col11 = st.columns(2)
with col10:
    max_pages = st.slider("최대 페이지 수", min_value=1, max_value=10, value=3)
with col11:
    delay_time = st.slider("요청 간격 (초)", min_value=0.1, max_value=2.0, value=0.5, step=0.1)

def get_coordinates_from_legal_code(cortarNo):
    """법정동 코드로부터 대략적인 좌표를 얻는 함수"""
    # 주요 지역별 좌표 매핑 (더 정확한 주소 검색을 위해)
    coord_mapping = {
        # 서울 강남구
        "1168010100": ("37.5009", "127.0374"),  # 역삼동
        "1168010200": ("37.5139", "127.0379"),  # 논현동  
        "1168010300": ("37.5271", "127.0276"),  # 압구정동
        "1168010400": ("37.5175", "127.0203"),  # 신사동
        "1168010500": ("37.5197", "127.0486"),  # 청담동
        "1168010600": ("37.5089", "127.0637"),  # 삼성동
        "1168010700": ("37.4946", "127.0619"),  # 대치동
        "1168010800": ("37.4782", "127.0761"),  # 개포동
        
        # 서울 서초구  
        "1165010100": ("37.4833", "127.0327"),  # 서초동
        "1165010200": ("37.5229", "127.0114"),  # 잠원동
        "1165010300": ("37.5035", "127.0070"),  # 반포동
        "1165010400": ("37.4817", "126.9965"),  # 방배동
        "1165010500": ("37.4845", "127.0371"),  # 양재동
        
        # 서울 종로구
        "1111013100": ("37.5701", "126.9835"),  # 종로1가
        "1111013200": ("37.5658", "126.9859"),  # 종로2가
        "1111013300": ("37.5700", "126.9910"),  # 종로3가
    }
    
    # 법정동 코드에 해당하는 좌표가 있으면 사용, 없으면 기본값
    if cortarNo in coord_mapping:
        return coord_mapping[cortarNo]
    else:
        # 기본 좌표 (서울 시청 기준)
        return ("37.5665", "126.9780")

def scrape_property_details(atclNo, headers):
    """매물 상세 정보를 스크래핑하는 함수"""
    try:
        detail_url = f"https://m.land.naver.com/article/info/{atclNo}"
        response = requests.get(detail_url, headers=headers, timeout=15)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, "html.parser")
        
        zoning = purpose = management_fee = ""
        detailed_address = ""
        
        # 상세 주소 정보 추출
        address_selectors = [
            "div.article_header",
            "div.article_info",
            "div.location_info", 
            "div.addr_info",
            "h1",
            "h2",
            ".article_title"
        ]
        
        for selector in address_selectors:
            elements = soup.select(selector)
            for elem in elements:
                text = elem.get_text(strip=True)
                if text and any(keyword in text for keyword in ["동", "로", "길", "가", "번지", "번지"]):
                    if len(text) > len(detailed_address) and len(text) < 100:  # 너무 긴 텍스트 제외
                        detailed_address = text
        
        # 다양한 셀렉터로 정보 추출 시도
        detail_selectors = [
            "div.detail_box",
            "div.info_detail",
            "div.article_detail",
            "div.detail_info",
            "div.item_detail"
        ]
        
        detail_boxes = []
        for selector in detail_selectors:
            boxes = soup.find_all("div", class_=selector.split('.')[1] if '.' in selector else selector)
            if boxes:
                detail_boxes.extend(boxes)
        
        # 추가로 모든 div 태그에서 키워드 검색
        if not detail_boxes:
            all_divs = soup.find_all("div")
            detail_boxes = [div for div in all_divs if div.get_text() and any(keyword in div.get_text() for keyword in ["용도지역", "건물용도", "관리비"])]
        
        for box in detail_boxes:
            text = box.get_text(strip=True, separator="\n")
            
            # 용도지역 추출
            if "용도지역" in text and not zoning:
                try:
                    lines = text.split("\n")
                    for i, line in enumerate(lines):
                        if "용도지역" in line and i + 1 < len(lines):
                            zoning = lines[i + 1].strip()
                            break
                except (IndexError, AttributeError):
                    pass
            
            # 건물용도 추출  
            if "건물용도" in text and not purpose:
                try:
                    lines = text.split("\n")
                    for i, line in enumerate(lines):
                        if "건물용도" in line and i + 1 < len(lines):
                            purpose = lines[i + 1].strip()
                            break
                except (IndexError, AttributeError):
                    pass
            
            # 관리비 추출
            if "관리비" in text and not management_fee:
                try:
                    lines = text.split("\n")
                    for i, line in enumerate(lines):
                        if "관리비" in line and i + 1 < len(lines):
                            management_fee = lines[i + 1].strip()
                            break
                except (IndexError, AttributeError):
                    pass
        
        # 공장/창고의 경우 특별 처리
        if not purpose:
            # 제목이나 설명에서 용도 추출 시도
            title_elem = soup.find("h1") or soup.find("title")
            if title_elem:
                title_text = title_elem.get_text()
                if any(keyword in title_text for keyword in ["공장", "창고", "물류", "제조", "생산"]):
                    purpose = "공장/창고"
                    
        return zoning, purpose, management_fee, detailed_address
        
    except requests.RequestException as e:
        logger.error(f"매물 상세 정보 요청 오류 (ID: {atclNo}): {e}")
        return "", "", "", ""
    except Exception as e:
        logger.error(f"매물 상세 정보 파싱 오류 (ID: {atclNo}): {e}")
        return "", "", "", ""

def filter_by_conditions(results, area_filter_enabled, min_area, max_area, price_filter_enabled, min_price, max_price):
    """수집된 데이터를 조건에 따라 필터링하는 함수"""
    if not results:
        return results
    
    filtered_results = []
    
    for result in results:
        # 면적 필터링
        if area_filter_enabled:
            area_str = result.get("전용면적(㎡)", "") or result.get("임대면적(㎡)", "")
            if area_str:
                try:
                    # 숫자만 추출 (예: "85.5㎡" -> 85.5)
                    area_match = re.search(r'[\d,]+\.?\d*', str(area_str).replace(',', ''))
                    if area_match:
                        area = float(area_match.group())
                        if not (min_area <= area <= max_area):
                            continue
                    else:
                        continue  # 면적 정보가 없으면 제외
                except (ValueError, TypeError):
                    continue  # 면적 파싱 실패시 제외
        
        # 가격 필터링  
        if price_filter_enabled:
            price_str = result.get("보증금/매매가", "")
            if price_str:
                try:
                    # 숫자만 추출하고 만원 단위로 변환
                    # 억, 만원 단위 처리
                    price_text = str(price_str).replace(',', '')
                    
                    # 억원 처리
                    if '억' in price_text:
                        eok_match = re.search(r'(\d+(?:\.\d+)?)억', price_text)
                        man_match = re.search(r'(\d+(?:\.\d+)?)만', price_text)
                        
                        eok_value = float(eok_match.group(1)) * 10000 if eok_match else 0
                        man_value = float(man_match.group(1)) if man_match else 0
                        price = eok_value + man_value
                    # 만원만 있는 경우
                    elif '만' in price_text:
                        man_match = re.search(r'(\d+(?:\.\d+)?)만', price_text)
                        price = float(man_match.group(1)) if man_match else 0
                    # 숫자만 있는 경우 (만원 단위로 가정)
                    else:
                        num_match = re.search(r'(\d+(?:\.\d+)?)', price_text)
                        price = float(num_match.group(1)) if num_match else 0
                    
                    if not (min_price <= price <= max_price):
                        continue
                except (ValueError, TypeError, AttributeError):
                    continue  # 가격 파싱 실패시 제외
        
        filtered_results.append(result)
    
    return filtered_results

def search_properties(cortarNo, rletTpCd, tradTpCd, max_pages, delay_time, sido="", sigungu="", eupmyeondong=""):
    """매물을 검색하는 함수"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
        "Referer": "https://m.land.naver.com/",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Cache-Control": "no-cache"
    }
    
    # 법정동 코드에 따른 좌표 설정
    lat, lon = get_coordinates_from_legal_code(cortarNo)
    z = "15"
    
    # 검색 범위 확장 (더 넓은 범위에서 검색)
    range_offset = 0.02  # 범위를 더 넓게 설정
    btm, top = float(lat) - range_offset, float(lat) + range_offset
    lft, rgt = float(lon) - range_offset, float(lon) + range_offset
    
    all_results = []
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        for page in range(1, max_pages + 1):
            status_text.text(f"페이지 {page}/{max_pages} 검색 중...")
            progress_bar.progress(page / max_pages)
            
            list_url = (
                f"https://m.land.naver.com/cluster/ajax/articleList?"
                f"itemId={cortarNo}&rletTpCd={rletTpCd}&tradTpCd={tradTpCd}&"
                f"z={z}&lat={lat}&lon={lon}&btm={btm}&lft={lft}&top={top}&rgt={rgt}&page={page}"
            )
            
            response = requests.get(list_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            data = response.json()
            items = data.get("body", [])
            
            if not items:
                status_text.text(f"페이지 {page}에서 더 이상 매물이 없습니다.")
                break
            
            for i, item in enumerate(items):
                atclNo = item.get("atclNo")
                if not atclNo:
                    continue
                    
                status_text.text(f"페이지 {page}/{max_pages} - 매물 {i+1}/{len(items)} 상세정보 수집 중...")
                
                # 면적 정보 개선 - 다양한 형식 처리
                area_info = {
                    "전용면적(㎡)": item.get("spc2", ""),
                    "임대면적(㎡)": item.get("spc1", ""),
                    "연면적(㎡)": item.get("spc3", ""),
                    "대지면적(㎡)": item.get("spc4", ""),
                }
                
                # 가격 정보 개선 - 형식 표준화
                price_info = {
                    "보증금/매매가": item.get("hanPrc", ""),
                    "월세": item.get("rentPrc", ""),
                    "전세금": item.get("rentPrc", "") if tradTpCd == "B1" else "",
                }
                
                basic_info = {
                    "매물번호": atclNo,
                    "층수": item.get("flrInfo", ""),
                    **area_info,
                    **price_info,
                    "건물명": item.get("bildNm", ""),
                    "방향": item.get("direction", ""),
                    "매물타입": property_types.get(rletTpCd, rletTpCd),
                    "거래타입": {"A1": "매매", "B1": "전세", "B2": "월세"}.get(tradTpCd, tradTpCd),
                }
                
                # 상세 정보 수집 (웹 스크래핑)
                zoning, purpose, management_fee, scraped_address = scrape_property_details(atclNo, headers)
                
                # 주소 정보 개선 - 법정동 기반 주소 생성
                # 세션에서 입력된 지역 정보 활용
                base_address = ""
                detailed_address = ""
                
                # 입력된 지역 정보로 기본 주소 생성
                if sido and sigungu and eupmyeondong:
                    base_address = f"{sido} {sigungu} {eupmyeondong}"
                
                # API에서 제공되는 주소 정보들
                address_candidates = [
                    item.get("atclNm", ""),
                    item.get("addr1", ""),
                    item.get("addr2", ""),
                    item.get("bildNm", ""),
                    scraped_address
                ]
                
                # 상세 주소 선택 로직
                for addr in address_candidates:
                    if addr and addr.strip():
                        addr_clean = addr.strip()
                        # "일반상가", "상가", "오피스텔" 등의 일반적인 단어만 있는 경우 제외
                        if not any(only_word in addr_clean for only_word in ["일반상가", "상가", "오피스텔", "아파트", "빌라"]):
                            # 실제 주소가 포함된 경우 (도로명, 지번 등)
                            if any(addr_keyword in addr_clean for addr_keyword in ["로", "길", "동", "가", "번지", "번", "-"]):
                                detailed_address = addr_clean
                                break
                
                # 주소가 여전히 없으면 기본 법정동 정보 사용
                if not base_address:
                    base_address = f"법정동코드: {cortarNo}"
                
                # 상세주소가 없으면 빈 문자열
                if not detailed_address:
                    detailed_address = ""
                
                result = {
                    **basic_info,
                    "주소지": base_address,
                    "상세주소": detailed_address,
                    "용도": purpose,
                    "지역지구": zoning,
                    "관리비": management_fee,
                    "매물 링크": f"https://m.land.naver.com/article/info/{atclNo}",
                    "수집일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
                all_results.append(result)
                time.sleep(delay_time)
            
            # 페이지 간 딜레이
            if page < max_pages:
                time.sleep(delay_time * 2)
                
    except requests.RequestException as e:
        st.error(f"❌ 네이버 접속 오류: {str(e)}")
        logger.error(f"네이버 접속 오류: {e}")
    except Exception as e:
        st.error(f"❌ 예상치 못한 오류: {str(e)}")
        logger.error(f"예상치 못한 오류: {e}")
    finally:
        progress_bar.empty()
        status_text.empty()
    
    return all_results

# --- 3. 검색 실행 버튼 ---
st.subheader("3️⃣ 매물 검색 및 엑셀 저장")

if st.button("🚀 매물 검색 시작"):
    if st.session_state.cortarNo:
        if sido and sigungu and eupmyeondong:
            st.info("🔄 데이터 수집 중... 잠시만 기다려주세요.")
            
            results = search_properties(
                st.session_state.cortarNo, 
                rletTpCd, 
                tradTpCd, 
                max_pages, 
                delay_time,
                sido,
                sigungu, 
                eupmyeondong
            )
            
            if results:
                # 조건 필터링 적용
                if area_filter_enabled or price_filter_enabled:
                    original_count = len(results)
                    results = filter_by_conditions(
                        results, 
                        area_filter_enabled, min_area, max_area,
                        price_filter_enabled, min_price, max_price
                    )
                    filtered_count = len(results)
                    
                    if filtered_count < original_count:
                        st.info(f"📊 필터링 결과: {original_count}개 중 {filtered_count}개 매물이 조건에 맞습니다.")
                
                if results:
                    st.session_state.search_results = results
                    df = pd.DataFrame(results)
                    
                    st.success(f"🎉 총 {len(results)}개의 매물 정보를 수집했습니다!")
                    
                    # 미리보기
                    st.subheader("📊 수집된 데이터 미리보기")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    if len(results) > 10:
                        st.info(f"처음 10개만 표시됩니다. 전체 {len(results)}개 데이터는 엑셀 파일에서 확인하세요.")
                    
                    # 필터 조건 요약 표시
                    if area_filter_enabled or price_filter_enabled:
                        st.subheader("🔍 적용된 필터 조건")
                        filter_info = []
                        if area_filter_enabled:
                            filter_info.append(f"📐 면적: {min_area}㎡ ~ {max_area}㎡")
                        if price_filter_enabled:
                            filter_info.append(f"💰 가격: {min_price:,}만원 ~ {max_price:,}만원")
                        st.info(" | ".join(filter_info))
                    
                    # 엑셀 파일 생성
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{sigungu}_{eupmyeondong}_매물정보_{timestamp}.xlsx"
                    
                    # 메모리에서 엑셀 파일 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='매물정보')
                        
                        # 워크시트 스타일링
                        worksheet = writer.sheets['매물정보']
                        for column in worksheet.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)
                            worksheet.column_dimensions[column_letter].width = adjusted_width
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 엑셀 파일 다운로드",
                        data=output.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ 설정한 조건에 맞는 매물이 없습니다. 조건을 완화해보세요.")
                    st.info("💡 **조건 완화 제안:**\n- 면적 범위를 넓혀보세요\n- 가격 범위를 조정해보세요\n- 다른 지역을 시도해보세요")
            else:
                st.warning("⚠️ 검색 결과가 없습니다. 검색 조건을 변경해보세요.")
        else:
            st.warning("모든 지역 정보를 입력해주세요.")
    else:
        st.warning("❗ 법정동 코드를 먼저 검색해주세요.")

# --- 기존 검색 결과 표시 ---
if st.session_state.search_results:
    st.subheader("📋 최근 검색 결과")
    df = pd.DataFrame(st.session_state.search_results)
    
    col12, col13 = st.columns(2)
    with col12:
        st.metric("총 매물 수", len(df))
    with col13:
        if not df.empty and '수집일시' in df.columns:
            st.metric("마지막 검색", df['수집일시'].iloc[0][:16])
    
    # 간단한 통계
    if not df.empty:
        st.subheader("📈 간단 통계")
        col14, col15, col16 = st.columns(3)
        
        with col14:
            if '전용면적(㎡)' in df.columns:
                try:
                    areas = pd.to_numeric(df['전용면적(㎡)'], errors='coerce').dropna()
                    if not areas.empty:
                        st.metric("평균 전용면적", f"{areas.mean():.1f}㎡")
                except:
                    pass
        
        with col15:
            if '층수' in df.columns:
                floor_counts = df['층수'].value_counts()
                if not floor_counts.empty:
                    st.metric("가장 많은 층수", floor_counts.index[0])
        
        with col16:
            if '용도' in df.columns:
                purpose_counts = df['용도'].value_counts()
                if not purpose_counts.empty:
                    st.metric("가장 많은 용도", purpose_counts.index[0])

# --- 도움말 및 주의사항 ---
with st.expander("💡 사용법 및 주의사항"):
    st.markdown("""
    ### 📖 사용 방법
    1. **지역 정보 입력**: 정확한 시/도, 시/군/구, 읍/면/동 이름을 입력하세요.
    2. **법정동 코드 검색**: 입력한 지역의 법정동 코드를 자동으로 찾습니다.
    3. **검색 조건 설정**: 매물 유형(건물/토지/공장)과 거래 유형을 선택하세요.
    4. **매물 검색**: 검색을 시작하면 자동으로 매물 정보를 수집합니다.
    5. **결과 확인 및 다운로드**: 수집된 데이터를 확인하고 엑셀 파일로 다운로드하세요.
    
    ### 🏭 매물 유형 설명
    - **아파트**: 아파트 매매/전세/월세
    - **오피스텔**: 오피스텔 매매/전세/월세  
    - **빌라/연립/다세대**: 다가구 주택
    - **분양권**: 아파트/오피스텔 분양권
    - **상가**: 상업용 건물
    - **사무실**: 오피스 공간
    - **공장/창고**: 공장, 창고, 물류센터 등
    - **토지**: 대지, 전, 답, 임야 등
    - **재개발/재건축**: 재개발/재건축 관련 매물
    
    ### 📐 면적 및 가격 필터링
    - **면적 조건**: 원하는 면적 범위로 매물을 필터링할 수 있습니다
    - **가격 조건**: 예산에 맞는 매물만 선별할 수 있습니다
    - **조건 조합**: 면적과 가격 조건을 동시에 적용 가능합니다
    
    ### ⚠️ 주의사항
    - 과도한 요청은 네이버 서버에 부하를 줄 수 있으니 적절한 간격을 두고 사용하세요.
    - 법정동코드.csv 파일이 필요합니다.
    - 일부 매물의 상세 정보가 수집되지 않을 수 있습니다.
    - 수집된 정보는 참고용으로만 사용하세요.
    
    ### 🔧 문제 해결
    - **법정동 코드를 찾을 수 없는 경우**: 지역명을 정확히 입력했는지 확인하세요.
    - **검색 결과가 없는 경우**: 다른 매물 유형이나 거래 유형을 시도해보세요.
    - **오류가 발생하는 경우**: 네트워크 연결을 확인하고 잠시 후 다시 시도하세요.
    """)

st.markdown("---")
st.markdown("**⚡ 개발: AI Assistant | 문의사항이 있으시면 개발자에게 연락하세요.**")