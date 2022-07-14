import streamlit as st
from dtod_function import *
from io import BytesIO
import openpyxl
import os
import datetime
import re
import webbrowser

st.set_page_config(page_title="D2D", layout="wide")
st.sidebar.header('설정')
url1 = 'http://10.51.160.87:18555/'
url2 = 'http://localhost:8080/'
mode = st.sidebar.radio('모드를 선택해 주세요.', ('Document to Data', 'Data to Document'), index=0)

if mode == 'Document to Data':

    st.header('Mode : Document to Data')
    st.subheader('Step 1. 양식파일 업로드 & 시트선택')
    form_file = st.file_uploader("양식 파일을 업로드 해주세요.(생략가능)", type=['xls', 'xlsx'])

    if form_file:
        xl = pd.ExcelFile(form_file)
        sheet_list = st.multiselect('처리할 Sheet를 선택하세요.', xl.sheet_names, xl.sheet_names)

    st.subheader('Step 2. 데이터시트 업로드')
    uploaded_file = st.file_uploader("엑셀 파일을 업로드 해주세요.", type=['xls', 'xlsx'], accept_multiple_files=True)

    cell_detail = st.checkbox("Cell detail 표시")
    start = st.button('실행')

    if start:
        if form_file is None:
            st.warning('양식파일이 업로드 되지 않았습니다. 데이터시트파일들을 서로 비교하여, 양식을 확인합니다.')
        if len(uploaded_file) == 0: st.warning('데이터시트파일이 업로드 되지 않았습니다. 데이터시트파일 업로드가 필요합니다.')

    if start & (len(uploaded_file) > 0):
        if form_file is not None:
            form_list = make_list_all(form_file, sheet_list)
        else:
            form_list = make_form_list(uploaded_file)
            sheet_list = form_list['sheet_name'].unique()
        st.success('- 양식 파일의 행 / 열 / 셀값을 모두 확인 하였습니다.')

        data_table = make_final_table(uploaded_file, form_list, sheet_list, cell_detail)
        st.success('- 처리가 완료 되었습니다. 결과 파일을 다운로드 받으세요.')
        st.write(data_table)
        in_memory_fp = BytesIO()
        data_table.to_excel(in_memory_fp)
        st.download_button(label="Download data as xlsx", data=in_memory_fp.getvalue(),
                           file_name='summary.xlsx', mime='application/vnd.ms-excel', )

        ## form_list를 엑셀로 받는 기능
        # in_memory_fp1 = BytesIO()
        # form_list.to_excel(in_memory_fp1)
        # st.download_button(label="Download form_list as xlsx", data=in_memory_fp1.getvalue(),
        #                    file_name='form_list.xlsx', mime='application/vnd.ms-excel', )

if mode == 'Data to Document':

    st.header('Mode : Data to Document')
    st.subheader('Step 1. 양식파일 업로드')
    form_file = st.file_uploader("양식 파일을 업로드 해주세요.", type=['xlsx'])
    st.subheader('Step 2. 데이터 파일 업로드')
    data_file = st.file_uploader("데이터 파일을 업로드 해주세요.", type=['xls', 'xlsx'])
    pw_required = st.checkbox("결과물 ZIP 파일에 암호를 보호 하고 싶으면 여기를 클릭하세요.")
    if pw_required:
        password = st.text_input("ZIP 파일에 사용 할 비밀번호를 입력 하세요.", type="password")
    start = st.button("실행")

    # 필요한 파일이 없을 경우 Warning
    if start:
        if form_file is None: st.warning('양식파일이 업로드 되지 않았습니다. 양식파일 업로드가 필요합니다.')
        if data_file is None: st.warning('데이터파일이 업로드 되지 않았습니다. 데이터파일 업로드가 필요합니다.')

    if start & (form_file is not None) & (data_file is not None):

        now = datetime.datetime.now()
        folder_name = now.strftime("%Y%m%d_%H%M%S")
        os.mkdir('output\\'+folder_name)

        df = pd.read_excel(data_file, index_col=None, header=0)
        header_list = list(df.columns)
        header_list = [item for item in header_list if item[-1] != '!']

        if 'sheet_name' not in header_list: st.warning('데이터 파일에 sheet_name 컬럼이 없습니다.')
        if 'address' not in header_list: st.warning('데이터 파일에 address 컬럼이 없습니다.')
        if len(header_list) <= 2: st.warning('하나 이상의 문서를 작성 할 데이터가 필요 합니다. 데이터 파일을 확인해 주세요.')

        if ('sheet_name' in header_list) & ('address' in header_list) & (len(header_list) > 2):
            header_list.remove('sheet_name')
            header_list.remove('address')
            st.write(str(len(header_list)) + '개의 문서를 작성하겠습니다.')

            c = 0
            txt1 = st.text('')
            prog = st.progress(0)
            txt2 = st.text('')

            for header in header_list:
                txt1.text(str(c) + '/' + str(len(header_list)) + '(' + str(int(c / len(header_list) * 100)) + '%)')
                txt2.text(header + '를 처리 중입니다.')
                wb = openpyxl.load_workbook(form_file)
                df_1 = df[['sheet_name', 'address', header]]
                df_1 = df_1[df_1[header].notna()]
                sheet_list = list(df_1.sheet_name.unique())

                for sheet_name in sheet_list:
                    sheet = wb[sheet_name]
                    df_2 = df_1[df_1.sheet_name == sheet_name]

                    for i in df_2.index:
                        a = df_2.loc[i, 'address']
                        b = df_2.loc[i, header]
                        sheet[a].value = b
                fn = re.sub('[\/:*?"<>|]', '', header)  # 파일이름에 사용 할 수 없는 특수문자를 정규식을 이용해 삭제함.
                wb.save('output\\'+folder_name +'\\' +fn + '.xlsx')
                wb.close()

                c += 1
                prog.progress(c / len(header_list))
                txt1.text(str(len(header_list)) + '/' + str(len(header_list)) + '(100%)')

            st.success('처리가 완료 되었습니다. 다운로드 창이 열립니다.')
            webbrowser.open_new_tab(url2+folder_name+'.zip')
