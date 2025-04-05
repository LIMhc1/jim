
import streamlit as st
import pandas as pd
import io

st.title("짐패스 양식 자동 생성기 (색상, 사이즈, 우편번호 보정 포함)")

uploaded_product = st.file_uploader("상품정리 파일 업로드 (list_url.xlsx)", type=["xlsx"])
uploaded_order = st.file_uploader("스마트스토어 주문관리 파일 업로드 (s_order.xlsx)", type=["xlsx"])
uploaded_template = st.file_uploader("짐패스 원본 양식 업로드 (jim.xlsx)", type=["xls", "xlsx"])

required_product_columns = ['옵션관리코드', '영문상품명', '짐패스품목코드', '현지가격', '구매url', '이미지url', '색상(영문)', '사이즈']
required_order_columns = ['수취인명', '수취인연락처1', '우편번호', '통합배송지', '개인통관고유부호', '배송메세지', '수량', '옵션관리코드', '주문번호']

if uploaded_product and uploaded_order and uploaded_template:
    try:
        product_df = pd.read_excel(uploaded_product, sheet_name=0)
        order_df = pd.read_excel(uploaded_order, skiprows=1)
        template_df = pd.read_excel(uploaded_template)

        # 필수 컬럼 확인
        if not all(col in product_df.columns for col in required_product_columns):
            st.error("❌ 상품정리 파일에 필수 컬럼이 누락되었습니다.")
            st.stop()
        if not all(col in order_df.columns for col in required_order_columns):
            st.error("❌ 주문관리 파일에 필수 컬럼이 누락되었습니다.")
            st.stop()

        # 옵션관리코드 중복 검사 (NaN 제외)
        valid_codes = product_df['옵션관리코드'].dropna()
        duplicate_codes = valid_codes[valid_codes.duplicated(keep=False)].unique().tolist()
        if duplicate_codes:
            st.error(f"❌ 상품정리 파일에 중복된 옵션관리코드가 있습니다: {duplicate_codes}")
            st.stop()

        # 병합
        product_df = product_df[required_product_columns]
        order_df = order_df[required_order_columns]

        merged_df = pd.merge(order_df, product_df, on='옵션관리코드', how='left', indicator=True)

        # 매칭 실패 행
        error_rows = merged_df[merged_df['_merge'] == 'left_only'].index.tolist()
        if error_rows:
            human_rows = [i + 2 for i in error_rows]
            st.error(f"❌ 옵션관리코드가 상품정리에 존재하지 않는 주문 행이 있습니다: {human_rows}")
            st.stop()

        # 수량 NaN/0 검사
        invalid_qty = merged_df['수량'].isna() | (merged_df['수량'] == 0)
        if invalid_qty.any():
            rows = [i + 2 for i in merged_df[invalid_qty].index.tolist()]
            st.error(f"❌ 수량이 비었거나 0인 주문 행이 있습니다: {rows}")
            st.stop()

        # 컬럼명 매핑
        merged_df.rename(columns={
            '수취인연락처1': '수취인 연락처',
            '통합배송지': '주소',
            '개인통관고유부호': '세관신고정보',
            '배송메세지': '택배사요청메모',
            '짐패스품목코드': '품목분류코드',
            '현지가격': '해외 단가',
            '구매url': '제품URL',
            '이미지url': '이미지URL'
        }, inplace=True)

        # 우편번호 5자리 문자열 처리
        merged_df['우편번호'] = merged_df['우편번호'].astype(str).str.zfill(5)

        # 고정값 추가
        merged_df['실물검수'] = 'Y'
        merged_df['해외 구매물품 보상보험'] = 'Y'
        merged_df['특이사항'] = ''

        # 최종 컬럼 순서
        final_columns = [
            '수취인명', '수취인 연락처', '우편번호', '주소', '세관신고정보',
            '택배사요청메모', '수량', '영문상품명', '품목분류코드', '색상(영문)', '사이즈',
            '해외 단가', '제품URL', '이미지URL',
            '실물검수', '해외 구매물품 보상보험', '특이사항'
        ]
        merged_data = merged_df[final_columns]

        # 템플릿 행 확장
        if len(template_df) < len(merged_data):
            extra = len(merged_data) - len(template_df)
            empty = pd.DataFrame('', index=range(extra), columns=template_df.columns)
            result_df = pd.concat([template_df, empty], ignore_index=True)
        else:
            result_df = template_df.copy()

        # 데이터 덮어쓰기
        for col in merged_data.columns:
            if col in result_df.columns:
                result_df.loc[:len(merged_data)-1, col] = merged_data[col]

        # 다운로드
        st.success("🎉 변환 성공! 아래에서 결과 파일을 다운로드하세요.")
        output = io.BytesIO()
        result_df.to_excel(output, index=False)
        st.download_button("📥 결과 파일 다운로드", output.getvalue(), file_name="짐패스_출력결과.xlsx")

    except Exception as e:
        st.error(f"오류 발생: {e}")
