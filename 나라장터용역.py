import json
from json.decoder import JSONDecoder
from os import error
from datetime import datetime
from dateutil.relativedelta import *
import requests
import pandas as pd

# ----------------------------------------------------------------------------------------------------------
SERVICE_KEY = "mpu3qIrJ7SuwS4LTIjh2aSrCYMCMNWiLFdy3k7HB/wEAD1gkHA+cUih69ttBhGo0Od2tQtreNv6bCGAv6qM3Sw=="
DMINSTTNM = "국세청"  # 기관명 : 일부만 입력해도 됨
START_YEAR = 2015  # 공고시작년도
END_YEAR = 2024  # 공고종료년도
bidNtceNm = "소득정보"  # 공고명
now = datetime.now()  # 파일명 시간

FILENAME = f"조달청입찰공고_{DMINSTTNM}_({START_YEAR}~{END_YEAR})_{now.strftime('%Y%m%d_%H%M%S')}.xlsx"  # 파일명
# ----------------------------------------------------------------------------------------------------------


def get_data(
    dminsttNm: str, year: int, start_month: int, end_month: int
) -> pd.DataFrame:
    start_date = datetime(year, start_month, 1)
    end_date = datetime(year, end_month, 1) + relativedelta(months=1)

    url = "http://apis.data.go.kr/1230000/BidPublicInfoService04/"
    api = "getBidPblancListInfoServcPPSSrch01"

    parameters = {
        "serviceKey": SERVICE_KEY,  # 공공데이터포탈에서 받은 인증키
        "pageNo": 1,  # 페이지 번호
        "numOfRows": 100,  # 한 페이지 결과 수
        "dminsttNm": DMINSTTNM,  # ※ 수요기관명 일부 입력시에도 조회 가능(방위사업청 연계건의 경우 : 발주기관(ornt)) 으로 검색
        "bidNtceNm": bidNtceNm,  # 공고명
        "inqryDiv": 1,  # 검색하고자하는 조회구분 1:공고게시일시, 2:개찰일시 (방위사업청 연계건의 경우 조회구분) 1. 공고게시일시 : 공고일자(pblancDate)
        "inqryBgnDt": start_date.strftime("%Y%m%d")
        + "0000",  # 검색하고자 하는 조회시작일시
        "inqryEndDt": end_date.strftime("%Y%m%d")
        + "0000",  # 검색하고자 하는 조회종료일시
        "intrntnlDivCd": 1,  # 검색하고자하는 국제구분코드 국내:1, 국제:2(방위사업청 연계건의 경우 아래 내용 참고하여 검색) 국내/시설 입찰 공고일 경우 : 1, 국외 입찰 공고일 경우 : 2
        "type": "json",  # 오픈API 리턴 타입을 JSON으로 받고 싶을 경우 'json' 으로 지정
    }

    # "ServiceKey" : 'mpu3qIrJ7SuwS4LTIjh2aSrCYMCMNWiLFdy3k7HB/wEAD1gkHA+cUih69ttBhGo0Od2tQtreNv6bCGAv6qM3Sw=='
    # response = requests.post(url=url, data=parameters)
    response = requests.get(url=url + api, params=parameters)
    print("-------------")
    try:
        if response.json()["response"]["body"]["totalCount"] == 0:
            print("0")
            return pd.DataFrame()

        df = pd.DataFrame(response.json()["response"]["body"]["items"])
        df = df.loc[
            :,
            [
                "bidNtceNo",
                "bidNtceOrd",
                "ntceKindNm",
                "infoBizYn",
                "bidNtceDt",
                "bidClseDt",
                "bidNtceNm",
                "asignBdgtAmt",
                "ntceInsttNm",
                "dminsttNm",
                "cntrctCnclsMthdNm",
                "srvceDivNm",
                "bidNtceUrl",
                "bidNtceDtlUrl",
                "stdNtceDocUrl",
                "ntceSpecDocUrl1",
                "ntceSpecDocUrl2",
                "ntceSpecFileNm1",
                "ntceSpecFileNm2",
            ],
        ]
        df.rename(
            columns={
                "bidNtceNo": "입찰공고번호",
                "bidNtceOrd": "입찰공고차수",
                "ntceKindNm": "공고상태",
                "infoBizYn": "정보화사업여부",
                "bidNtceDt": "공고일시",
                "bidClseDt": "입찰마감일시",
                "bidNtceNm": "입찰공고명",
                "asignBdgtAmt": "예산",
                "ntceInsttNm": "공고기관",
                "dminsttNm": "수요기관",
                "cntrctCnclsMthdNm": "계약방법",
                "srvceDivNm": "용역구분",
                "bidNtceDtlUrl": "입찰공고상세화면",
                "bidNtceUrl": "입찰공고URL",
                "stdNtceDocUrl": "표준공고문 URL",
                "ntceSpecDocUrl1": "공고문",
                "ntceSpecDocUrl2": "제안요청서",
                "ntceSpecFileNm1": "공고서파일명",
                "ntceSpecFileNm2": "과업지시서파일명",
            },
            inplace=True,
        )
        print(df["입찰공고번호"].size)

        df["입찰공고명"] = (
            '=HYPERLINK("' + df["입찰공고상세화면"] + '", "' + df["입찰공고명"] + '")'
        )

        df["공고문"] = (
            '=HYPERLINK("' + df["공고문"] + '", "' + df["공고서파일명"] + '")'
        )
        df["제안요청서"] = (
            '=HYPERLINK("' + df["제안요청서"] + '", "' + df["과업지시서파일명"] + '")'
        )
        df = df.drop(
            [
                "입찰공고상세화면",
                "입찰공고URL",
                "표준공고문 URL",
                "공고서파일명",
                "과업지시서파일명",
            ],
            axis=1,
        )
        return df
    except Exception as e:
        print(e)
        return None

    else:
        pass
    finally:
        pass


def save_file(df: pd.DataFrame) -> None:
    ## XlsxWriter 엔진으로 Pandas writer 객체 만들기
    writer = pd.ExcelWriter(
        FILENAME,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_numbers": False}},
    )
    ## DataFrame을 xlsx에 쓰기
    df.to_excel(writer, sheet_name="입찰공고", index=False)
    ## Pandas writer 객체에서 xlsxwriter 객체 가져오기
    workbook = writer.book
    worksheet = writer.sheets["입찰공고"]
    ## 포맷 만들기
    # format = workbook.add_format({"num_format": "#,##0.00"})
    format = workbook.add_format()
    format.set_font_color("red")
    format.set_num_format("#,##0")
    ## 예산금액 포매팅
    worksheet.write_column("H2:H10000", df["예산"], format)

    worksheet.autofit()
    ## Pandas writer 객체 닫기
    writer.close()
    print(FILENAME + "을 저장하였습니다.")

    return None


df = pd.DataFrame()

years = list(range(START_YEAR, END_YEAR + 1))
for year in years:
    print(f"{year}년...")
    df1 = get_data(DMINSTTNM, year, 1, 12)
    df = pd.concat([df, df1])

print("-------------------------------------------------------------")

save_file(df)
# df.to_excel(FILENAME, index=False)
