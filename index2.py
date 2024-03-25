import streamlit as st
import pandas as pd
import openpyxl  # Excel 파일로 저장하기 위해 필요
from io import BytesIO  # 파일 다운로드를 위해 필요

# 데이터프레임을 저장할 빈 데이터프레임 생성
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=['페이지 번호', '애니메이션 적용 대상', '사용할 대본'])

course_name = st.text_input("강의명", "도레미파이썬")

# 사용자 입력 받기
with st.form(key='record_form'):
    page_number = st.number_input("페이지 번호", value=None, step=1)
    
    animation_target = st.multiselect(
        '애니메이션 적용 대상',
        ['⛔ 없음', "🔠 텍스트", '🆚 도형을 포함한 텍스트', '🟪 도형', '🖼️ 이미지(아이콘)/코드', '✨ 효과', '👩‍🎨 애니메이션 제작 필요','🎸 기타'],
        default = ['⛔ 없음']
        )
    
    script = st.text_area('사용할 대본', height = 200)
    
    submit_button = st.form_submit_button(label='스크립트 추가')

# 정보 추가 버튼이 클릭되었을 때
if submit_button:
    # 입력받은 정보를 데이터프레임에 추가
    new_data = {'페이지 번호': page_number, '애니메이션 적용 대상': animation_target, '사용할 대본': script}
    st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([new_data])], ignore_index=True)


# 데이터프레임을 화면에 표시
st.checkbox("Use container width", value=True, key="use_container_width")

st.data_editor(
    st.session_state.df, 
    use_container_width=st.session_state.use_container_width,
    hide_index=True
    )

# 데이터프레임을 Excel 파일로 변환하는 함수
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    # 파일 데이터를 가져와서 반환
    return output.getvalue()

# 데이터프레임에서 마지막 행을 삭제하는 버튼
if st.button('마지막 스크립트 삭제'):
    if not st.session_state.df.empty:
        st.session_state.df = st.session_state.df[:-1]
    else:
        st.warning('데이터프레임이 비어 있습니다.')


if st.button('엑셀 다운로드'):
    val = to_excel(st.session_state.df)
    st.download_button(label='현재 데이터 다운로드', data=val, file_name=f"{course_name}_script.xlsx", mime='application/vnd.ms-excel')


from docx import Document  # 워드 파일로 저장하기 위해 필요

def to_word(df):
    doc = Document()
    
    # 데이터프레임의 컬럼명 추가
    doc.add_paragraph("\n".join(df.columns))
    doc.add_paragraph("==========")
    
    # 데이터프레임의 데이터 추가
    for index, row in df.iterrows():
        row_data = "\n".join(str(value) for value in row)
        doc.add_paragraph(row_data)
        doc.add_paragraph("==========")


    # 파일 데이터를 BytesIO 객체로 저장하고 반환
    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.getvalue()

# 데이터프레임을 워드 파일로 다운로드하는 버튼
if st.button('워드 다운로드'):
    word_val = to_word(st.session_state.df)
    st.download_button(label='현재 데이터 워드로 다운로드', data=word_val, file_name=f"{course_name}_script.docx", mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

st.divider()

# 데이터프레임 초기화 버튼
if st.button('스크립트 초기화'):
    st.session_state.df = pd.DataFrame(columns=['페이지 번호', '애니메이션 적용 대상', '사용할 대본'])