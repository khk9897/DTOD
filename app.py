import streamlit as st
from dtod_function import *

import time
from io import BytesIO
import openpyxl
import os
import datetime
import re
from smb.SMBConnection import SMBConnection
import shutil

st.set_page_config(page_title="D2D", layout="wide")
st.sidebar.header('설정')
mode = st.sidebar.selectbox('모드를 선택해 주세요.', ('Document to Data', 'Data to Document'), index=0)

if mode == 'Document to Data':

    st.header('Mode : Document to Data')
    st.subheader('Step 1. 양식파일 업로드 & 시트선택')
    form_file = st.file_uploader("양식 파일을 업로드 해주세요.", type=['xls', 'xlsx'])

    if form_file:
        xl = pd.ExcelFile(form_file)
        sheet_list = st.multiselect('처리할 Sheet를 선택하세요.', xl.sheet_names, xl.sheet_names)

    st.subheader('Step 2. 데이터시트 업로드')
    uploaded_file = st.file_uploader("엑셀 파일을 업로드 해주세요.", type=['xls', 'xlsx'], accept_multiple_files=True)

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

        data_table = make_final_table(uploaded_file)
        st.success('- 처리가 완료 되었습니다. 결과 파일을 다운로드 받으세요.')
        st.write(data_table)
        in_memory_fp = BytesIO()
        data_table.to_excel(in_memory_fp)
        st.download_button(label="Download data as xlsx", data=in_memory_fp.getvalue(),
                           file_name='summary.xlsx', mime='application/vnd.ms-excel', )
        in_memory_fp1 = BytesIO()
        form_list.to_excel(in_memory_fp1)
        st.download_button(label="Download form_list as xlsx", data=in_memory_fp1.getvalue(),
                           file_name='form_list.xlsx', mime='application/vnd.ms-excel', )
if mode == 'Data to Document':

    st.header('Mode : Data to Document')
    st.subheader('Step 1. 양식파일 업로드')
    form_file = st.file_uploader("양식 파일을 업로드 해주세요.", type=['xls', 'xlsx'])
    st.subheader('Step 2. 데이터 파일 업로드')
    data_file = st.file_uploader("데이터 파일을 업로드 해주세요.", type=['xls', 'xlsx'])
    start = st.button("실행")

    # 필요한 파일이 없을 경우 Warning
    if start:
        if form_file is None: st.warning('양식파일이 업로드 되지 않았습니다. 양식파일 업로드가 필요합니다.')
        if data_file is None: st.warning('데이터파일이 업로드 되지 않았습니다. 데이터파일 업로드가 필요합니다.')

    if start & (form_file is not None) & (data_file is not None):

        now = datetime.datetime.now()
        folder_name = '[999_temp]\\' + now.strftime("%Y%m%d_%H%M%S")
        os.mkdir(folder_name)

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

            # 공유 폴더 업로드를 위한 부분
            conn = SMBConnection(userID, password, client_machine_name, server_name, domain=domain_name,
                                 use_ntlm_v2=True,
                                 is_direct_tcp=True)
            connected = conn.connect(server_ip, 445)
            conn.createDirectory(service, folder_name, timeout=30)
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
                wb.save(folder_name + fn + '.xlsx')
                wb.close()
                c += 1
                prog.progress(c / len(header_list))
                txt1.text(str(len(header_list)) + '/' + str(len(header_list)) + '(100%)')

                # 공유 폴더 업로드를 위한 부분
                with open(folder_name + header + '.xlsx', 'rb') as file_obj:
                    conn.storeFile(service, folder_name + r'\\' + header + '.xlsx', file_obj)

            st.success('처리가 완료 되었습니다. 다음 공유폴더에서 결과물을 확인 하세요.')
            st.success(r'\\\\' + server_ip + r'\\' + service + r'\\' + folder_name)
            # shutil.rmtree(folder_name) #공유폴더에 Upload 후 서버에서 결과 파일을 삭제 하기 위함.
