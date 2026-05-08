import os
import time
import requests
import feedparser
import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from urllib.parse import quote
from datetime import datetime

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


# =========================
# 기본 설정
# =========================

st.set_page_config(
    page_title="ICN AI Travel Guide",
    page_icon="✈️",
    layout="wide"
)

DEFAULT_EXCEL_PATH = "boarding_pass_manual_input_examples_v3.xlsx"


# =========================
# 항공/공항 관련 제외 키워드
# =========================

EXCLUDE_KEYWORDS = [
    "항공", "항공사", "항공권", "비행기", "공항", "인천공항",
    "출국", "입국", "환승", "탑승", "수하물", "기내식", "라운지",
    "대한항공", "아시아나", "제주항공", "진에어", "티웨이",
    "에어부산", "에어서울", "델타항공", "카타르항공",
    "에미레이트", "싱가포르항공", "캐세이퍼시픽",
    "운항", "노선", "마일리지", "터미널", "게이트"
]


CITY_GUIDE = {
    "일본": {
        "주의사항": "교통카드, 현금 결제 가능성, 지하철 환승 동선을 미리 확인하는 것이 좋습니다.",
        "추천활동": ["현지 교통패스 확인", "맛집 대기시간 확인", "관광지 예약 여부 확인"],
        "음식": ["라멘", "스시", "규카츠", "편의점 간식"]
    },
    "태국": {
        "주의사항": "더운 날씨와 교통 체증을 고려해 이동 시간을 넉넉히 잡는 것이 좋습니다.",
        "추천활동": ["그랩 또는 공항철도 확인", "야시장 방문", "마사지 예약"],
        "음식": ["팟타이", "똠얌꿍", "망고스티키라이스"]
    },
    "미국": {
        "주의사항": "도시별 치안과 이동수단을 미리 확인하는 것이 좋습니다.",
        "추천활동": ["우버/리프트 확인", "현지 맛집 예약", "주요 관광지 운영시간 확인"],
        "음식": ["버거", "스테이크", "타코", "브런치"]
    },
    "프랑스": {
        "주의사항": "소매치기와 대중교통 지연 정보를 확인하는 것이 좋습니다.",
        "추천활동": ["박물관 예약", "카페 투어", "야경 명소 확인"],
        "음식": ["크루아상", "에스카르고", "스테이크 프리트", "마카롱"]
    },
    "카타르": {
        "주의사항": "복장 문화와 현지 규정을 확인하는 것이 좋습니다.",
        "추천활동": ["수크 와키프 방문", "사막 투어 확인", "현지 문화시설 방문"],
        "음식": ["마크부스", "샤와르마", "아랍식 커피"]
    },
    "홍콩": {
        "주의사항": "옥토퍼스 카드와 대중교통 동선을 미리 확인하는 것이 좋습니다.",
        "추천활동": ["빅토리아 피크", "침사추이 야경", "로컬 맛집 탐방"],
        "음식": ["딤섬", "완탕면", "에그타르트"]
    },
    "아랍에미리트": {
        "주의사항": "실내외 온도 차가 크고, 관광지 예약이 필요한 경우가 많습니다.",
        "추천활동": ["두바이몰", "부르즈 할리파 예약", "사막 투어"],
        "음식": ["샤와르마", "만디", "중동식 그릴"]
    },
    "싱가포르": {
        "주의사항": "벌금 규정이 엄격하므로 음식물, 흡연, 쓰레기 관련 규정을 주의하세요.",
        "추천활동": ["마리나베이샌즈", "호커센터 방문", "가든스 바이 더 베이"],
        "음식": ["치킨라이스", "락사", "칠리크랩"]
    },
    "대만": {
        "주의사항": "이지카드 사용 가능 여부와 야시장 운영 시간을 확인하는 것이 좋습니다.",
        "추천활동": ["야시장 방문", "타이베이101", "카페 투어"],
        "음식": ["우육면", "지파이", "버블티", "샤오롱바오"]
    }
}


# =========================
# 엑셀 로드
# =========================

@st.cache_data
def load_excel(file):
    df = pd.read_excel(file, header=2)
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df.columns = [str(col).strip() for col in df.columns]

    if "번호" in df.columns:
        df = df[df["번호"].notna()]

    return df


# =========================
# 텍스트 처리
# =========================

def clean_html(text):
    if not text:
        return ""

    soup = BeautifulSoup(str(text), "html.parser")
    return soup.get_text(" ", strip=True)


def clean_destination_text(text):
    if pd.isna(text):
        return ""

    text = str(text)
    text = text.replace("/", " ")
    text = text.replace("\\", " ")
    text = text.replace("|", " ")
    return text.strip()


def simplify_city_name(city):
    city = clean_destination_text(city)

    mapping = {
        "도쿄 나리타": "도쿄",
        "도쿄/나리타": "도쿄",
        "방콕 수완나품": "방콕",
        "방콕/수완나품": "방콕",
        "파리 샤를드골": "파리",
        "파리/샤를드골": "파리",
        "타이베이 타오위안": "타이베이",
        "타이베이/타오위안": "타이베이",
        "오키나와 나하": "오키나와",
        "오키나와/나하": "오키나와"
    }

    for key, value in mapping.items():
        if key in city:
            return value

    if "도쿄" in city:
        return "도쿄"
    if "방콕" in city:
        return "방콕"
    if "파리" in city:
        return "파리"
    if "타이베이" in city:
        return "타이베이"
    if "오키나와" in city or "나하" in city:
        return "오키나와"
    if "로스앤젤레스" in city:
        return "로스앤젤레스"
    if "싱가포르" in city:
        return "싱가포르"
    if "홍콩" in city:
        return "홍콩"
    if "두바이" in city:
        return "두바이"
    if "도하" in city:
        return "도하"

    return city.split()[0].strip()


def is_valid_travel_content(title, summary):
    text = f"{title} {summary}".lower()

    for keyword in EXCLUDE_KEYWORDS:
        if keyword.lower() in text:
            return False

    return True


# =========================
# 이미지 추출
# =========================

def extract_image_from_summary(summary):
    if not summary:
        return None

    soup = BeautifulSoup(str(summary), "html.parser")
    img = soup.find("img")

    if img and img.get("src"):
        return img.get("src")

    return None


@st.cache_data(ttl=3600)
def get_article_image(url):
    if not url:
        return None

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    try:
        response = requests.get(url, headers=headers, timeout=5)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        og_image = soup.find("meta", property="og:image")
        if og_image and og_image.get("content"):
            return og_image.get("content")

        twitter_image = soup.find("meta", attrs={"name": "twitter:image"})
        if twitter_image and twitter_image.get("content"):
            return twitter_image.get("content")

        first_img = soup.find("img")
        if first_img and first_img.get("src"):
            return first_img.get("src")

    except Exception:
        return None

    return None


# =========================
# Google RSS 여행 정보 검색
# =========================

@st.cache_data(ttl=900)
def get_google_travel_results(query, max_items=5):
    """
    Google News RSS 기반 검색.
    - 최신 14일 -> 30일 -> 90일 -> 기간 제한 없음 순서로 검색 확장
    - 항공/공항 관련 결과 제외
    - 검색 결과 상위 순위와 최신성을 우선 반영
    """

    search_ranges = [
        "14d",
        "30d",
        "90d",
        None
    ]

    all_items = []
    seen_links = set()

    for days in search_ranges:
        if days:
            search_query = f"{query} when:{days}"
            freshness_label = f"최근 {days}"
        else:
            search_query = query
            freshness_label = "기간 제한 없음"

        encoded_query = quote(search_query)

        url = (
            "https://news.google.com/rss/search?"
            f"q={encoded_query}&hl=ko&gl=KR&ceid=KR:ko"
        )

        try:
            feed = feedparser.parse(url)
        except Exception:
            continue

        for rank, entry in enumerate(feed.entries):
            title = clean_html(entry.get("title", ""))
            raw_summary = entry.get("summary", "")
            summary = clean_html(raw_summary)
            link = entry.get("link", "")

            if not link or link in seen_links:
                continue

            if not is_valid_travel_content(title, summary):
                continue

            seen_links.add(link)

            published_text = entry.get("published", "")
            published_dt = None

            if hasattr(entry, "published_parsed") and entry.published_parsed:
                try:
                    published_dt = datetime.fromtimestamp(
                        time.mktime(entry.published_parsed)
                    )
                except Exception:
                    published_dt = None

            image_url = extract_image_from_summary(raw_summary)
            if not image_url:
                image_url = get_article_image(link)

            item = {
                "title": title,
                "link": link,
                "published": published_text,
                "published_dt": published_dt,
                "summary": summary,
                "freshness_label": freshness_label,
                "search_rank": rank,
                "image_url": image_url
            }

            all_items.append(item)

        if len(all_items) >= max_items:
            break

    def score_item(item):
        date_score = 0

        if item["published_dt"]:
            age_days = (datetime.now() - item["published_dt"]).days
            date_score = max(0, 365 - age_days)

        rank_score = max(0, 100 - item["search_rank"])

        return date_score + rank_score

    all_items = sorted(
        all_items,
        key=score_item,
        reverse=True
    )

    return all_items[:max_items]


# =========================
# 공항 행동 가이드
# =========================

def make_airport_action_guide(row):
    flight = row["항공편명"]
    dep_airport = row["출발 공항"]
    dep_terminal = row["출발 터미널"]
    arr_city = row["도착 도시"]
    arr_country = row["도착 국가"]

    return {
        "현재 해야 할 일": [
            f"{dep_airport} {dep_terminal} 출발 항공편 {flight} 기준으로 체크인 및 출국 준비를 진행하세요.",
            "수하물이 있다면 먼저 체크인 카운터로 이동하세요.",
            "수하물이 없다면 모바일 체크인 여부를 확인한 뒤 바로 보안검색대로 이동할 수 있습니다."
        ],
        "어디로 가야 하는지": [
            f"출발 터미널은 {dep_terminal}입니다.",
            "체크인 완료 후 보안검색 → 출국심사 → 면세구역 → 탑승구 순서로 이동하세요.",
            "탑승구가 멀 수 있으므로 면세구역 진입 후 게이트 위치를 먼저 확인하세요."
        ],
        "시간 부족 판단": [
            "탑승 시작 60분 이상 전이면 비교적 여유가 있습니다.",
            "탑승 시작 30~60분 전이면 면세점은 짧게 이용하는 것을 권장합니다.",
            "탑승 시작 30분 이내라면 라운지/면세점 이용보다 바로 탑승구 이동을 권장합니다."
        ],
        "면세점/라운지": [
            "면세점은 보안검색 통과 후 이용 가능합니다.",
            "라운지는 최소 45분 이상 여유가 있을 때만 추천합니다.",
            "탑승구가 먼 경우 라운지 이용 가능 시간은 더 짧아질 수 있습니다."
        ],
        "도착지 요약": [
            f"도착지는 {arr_city}, 도착 국가는 {arr_country}입니다.",
            "도착 후 입국심사, 수하물 수령, 시내 이동수단을 미리 확인하세요."
        ]
    }


# =========================
# 요약 생성
# =========================

def make_rule_based_summary(row, trend_items, travel_items):
    country = row["도착 국가"]
    city = simplify_city_name(row["도착 도시"])
    flight = row["항공편명"]

    local = CITY_GUIDE.get(country, {
        "주의사항": "현지 교통, 날씨, 결제수단, 관광지 운영시간을 확인하는 것이 좋습니다.",
        "추천활동": ["시내 이동수단 확인", "숙소 체크인 시간 확인", "관광지 운영시간 확인"],
        "음식": ["현지 대표 음식"]
    })

    trend_title = trend_items[0]["title"] if trend_items else "관련 트렌드 정보 없음"
    travel_title = travel_items[0]["title"] if travel_items else "관련 여행 정보 없음"

    return f"""
### ✈️ 맞춤 여행 가이드 요약

현재 선택된 항공편은 **{flight}**이며, 도착지는 **{city} / {country}**입니다.

#### 1. 지금 공항에서 해야 할 일
출발 터미널은 **{row['출발 터미널']}**입니다.  
수하물이 있다면 체크인 카운터를 먼저 이용하고, 수하물이 없다면 모바일 체크인 여부를 확인한 뒤 보안검색대로 이동하는 것이 좋습니다.

#### 2. 도착지 활용 정보
현재 참고할 만한 지역 트렌드 정보는 **{trend_title}**입니다.  
여행 정보로는 **{travel_title}**을 우선 확인해볼 수 있습니다.

#### 3. 현지 주의사항
{country} 여행 시에는 **{local['주의사항']}**

#### 4. 추천 활동
- {local['추천활동'][0]}
- {local['추천활동'][1]}
- {local['추천활동'][2]}

#### 5. 추천 음식
- {", ".join(local["음식"])}
"""


def make_llm_summary(row, trend_items, travel_items):
    api_key = os.getenv("OPENAI_API_KEY")

    if not api_key or OpenAI is None:
        return make_rule_based_summary(row, trend_items, travel_items)

    client = OpenAI(api_key=api_key)

    prompt = f"""
너는 인천공항 탑승권 기반 여행 가이드 AI다.

아래 정보로 사용자에게 실용적인 여행 가이드를 작성해라.
항공사, 항공권, 공항 뉴스가 아니라 도착 도시에서 실제 여행에 활용할 수 있는 정보만 중심으로 설명해라.

[탑승권 정보]
날짜: {row['날짜']}
항공편명: {row['항공편명']}
출발 공항: {row['출발 공항']}
출발 터미널: {row['출발 터미널']}
도착 공항: {row['도착 공항']}
도착 도시: {row['도착 도시']}
도착 국가: {row['도착 국가']}

[지역 트렌드 정보]
{trend_items}

[여행 추천 정보]
{travel_items}

반드시 포함:
1. 지금 공항에서 해야 할 일
2. 면세점/라운지 이용 판단 기준
3. 도착 도시의 트렌드
4. 현지에서 바로 활용 가능한 장소/음식/동선 팁
5. 주의사항

한국어로 간결하게 작성해라.
"""

    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            input=prompt
        )
        return response.output_text

    except Exception as e:
        return f"""
LLM 호출 실패로 규칙 기반 안내를 표시합니다.

{make_rule_based_summary(row, trend_items, travel_items)}

오류 내용:
{e}
"""


# =========================
# 카드 렌더링
# =========================

def render_info_card(item):
    title = item.get("title", "")
    link = item.get("link", "")
    published = item.get("published", "")
    summary = item.get("summary", "")
    freshness_label = item.get("freshness_label", "")
    image_url = item.get("image_url", "")

    with st.container(border=True):
        col_img, col_text = st.columns([1, 3])

        with col_img:
            if image_url:
                try:
                    st.image(image_url, use_container_width=True)
                except Exception:
                    st.markdown(
                        """
                        <div style="
                            width:100%;
                            height:130px;
                            background-color:#f0f2f6;
                            border-radius:10px;
                            display:flex;
                            align-items:center;
                            justify-content:center;
                            color:#777;">
                            이미지 표시 실패
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
            else:
                st.markdown(
                    """
                    <div style="
                        width:100%;
                        height:130px;
                        background-color:#f0f2f6;
                        border-radius:10px;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        color:#777;">
                        이미지 없음
                    </div>
                    """,
                    unsafe_allow_html=True
                )

        with col_text:
            if link:
                st.markdown(f"### [{title}]({link})")
            else:
                st.markdown(f"### {title}")

            if published:
                st.caption(published)

            if freshness_label:
                st.caption(f"정보 기준: {freshness_label}")

            if summary:
                st.write(summary[:300] + "...")


# =========================
# Streamlit UI
# =========================

st.title("✈️ ICN AI Travel Guide")
st.caption("탑승권 정보를 입력하면 공항 행동 가이드와 도착지 중심 여행/트렌드 정보를 제공합니다.")

uploaded_file = st.sidebar.file_uploader("탑승권 정보 엑셀 업로드", type=["xlsx"])

if uploaded_file is not None:
    df = load_excel(uploaded_file)
else:
    if os.path.exists(DEFAULT_EXCEL_PATH):
        df = load_excel(DEFAULT_EXCEL_PATH)
    else:
        st.error("엑셀 파일이 없습니다. 사이드바에서 파일을 업로드하세요.")
        st.stop()


required_cols = [
    "번호",
    "날짜",
    "항공편명",
    "출발 공항",
    "출발 터미널",
    "도착 공항",
    "도착 도시",
    "도착 국가"
]

missing_cols = [c for c in required_cols if c not in df.columns]

if missing_cols:
    st.error(f"엑셀에 필요한 컬럼이 없습니다: {missing_cols}")
    st.write("현재 인식된 컬럼:")
    st.write(list(df.columns))
    st.stop()


# =========================
# 사이드바
# =========================

st.sidebar.subheader("항공권 선택")

flight_options = [
    f"{int(row['번호'])}번 | {row['항공편명']} | {row['도착 도시']}({row['도착 국가']})"
    for _, row in df.iterrows()
]

selected = st.sidebar.selectbox("항공권 선택", flight_options)
selected_index = flight_options.index(selected)
row = df.iloc[selected_index]

st.sidebar.divider()
st.sidebar.subheader("검색 설정")

trend_count = st.sidebar.slider("지역 트렌드 정보 개수", 3, 10, 5)
travel_count = st.sidebar.slider("여행 추천 정보 개수", 3, 10, 6)

city = simplify_city_name(row["도착 도시"])
country = clean_destination_text(row["도착 국가"])

trend_query = f"{city} 여행 트렌드 맛집 축제 관광"
travel_query = f"{city} 여행 추천 명소 핫플 맛집 코스"

trend_items = get_google_travel_results(trend_query, trend_count)
travel_items = get_google_travel_results(travel_query, travel_count)


# =========================
# 상단 카드
# =========================

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("항공편", row["항공편명"])

with col2:
    st.metric("출발", f"{row['출발 공항']} {row['출발 터미널']}")

with col3:
    st.metric("도착", row["도착 공항"])

with col4:
    st.metric("도착 지역", city)

st.divider()


# =========================
# 탑승권 정보
# =========================

st.subheader("📌 입력된 탑승권 정보")

info_col1, info_col2 = st.columns(2)

with info_col1:
    st.write(f"**날짜:** {row['날짜']}")
    st.write(f"**항공편명:** {row['항공편명']}")
    st.write(f"**출발 공항:** {row['출발 공항']}")
    st.write(f"**출발 터미널:** {row['출발 터미널']}")

with info_col2:
    st.write(f"**도착 공항:** {row['도착 공항']}")
    st.write(f"**도착 도시:** {row['도착 도시']}")
    st.write(f"**도착 국가:** {row['도착 국가']}")

st.divider()


# =========================
# 공항 행동 가이드
# =========================

st.subheader("🧭 공항 내 행동 가이드")

guide = make_airport_action_guide(row)

g1, g2 = st.columns(2)

with g1:
    with st.expander("지금 무엇을 해야 하는지", expanded=True):
        for item in guide["현재 해야 할 일"]:
            st.write(f"- {item}")

    with st.expander("어디로 가야 하는지", expanded=True):
        for item in guide["어디로 가야 하는지"]:
            st.write(f"- {item}")

with g2:
    with st.expander("시간이 부족한지 판단", expanded=True):
        for item in guide["시간 부족 판단"]:
            st.write(f"- {item}")

    with st.expander("면세점/라운지 이용 가능 여부", expanded=True):
        for item in guide["면세점/라운지"]:
            st.write(f"- {item}")

st.divider()


# =========================
# LLM 여행 가이드
# =========================

st.subheader("🤖 LLM 여행 가이드 요약")

if st.button("여행 가이드 생성", type="primary"):
    with st.spinner("여행 가이드를 생성 중입니다..."):
        summary = make_llm_summary(row, trend_items, travel_items)
        st.markdown(summary)

st.divider()


# =========================
# 지역 트렌드 정보
# =========================

st.subheader(f"📰 {city} 지역 트렌드 정보")
st.caption("항공사/공항 관련 정보는 제외하고, 관광·축제·맛집·현지 트렌드 중심으로 표시합니다.")

if trend_items:
    for item in trend_items:
        render_info_card(item)
else:
    st.info("지역 트렌드 정보를 찾지 못했습니다. 검색어를 조정하거나 잠시 후 다시 시도하세요.")


# =========================
# 여행 추천 정보
# =========================

st.subheader(f"🌍 {city} 여행 추천 정보")
st.caption("추천 명소, 핫플, 맛집, 여행 코스 중심으로 표시합니다.")

if travel_items:
    for item in travel_items:
        render_info_card(item)
else:
    st.info("여행 추천 정보를 찾지 못했습니다. 검색어를 조정하거나 잠시 후 다시 시도하세요.")


# =========================
# 전체 데이터
# =========================

st.divider()

with st.expander("전체 탑승권 데이터 보기"):
    st.dataframe(df, use_container_width=True)
