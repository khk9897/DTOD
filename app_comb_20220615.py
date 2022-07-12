import streamlit as st
import pandas as pd
import time
from io import BytesIO
import openpyxl
import os
import datetime
import re
from smb.SMBConnection import SMBConnection
import shutil

# 공유폴더 접근을 위한 정보
userID = 'kanghk1'
password = 'rotating'
client_machine_name = 'localpcname'
server_name = 'servername'
server_ip = '172.23.36.22'
domain_name = 'domainname'
service = 'RO_Mech_Root'

# table의 head 값을 제한하여, 보여주기한 문자열
# 하기 문자열안에 포함 되어 있는 경우에만 Above 값에 표시 됨.
above_limit = 'MaximumMinimumRatedNormalRev'


def ColIdxToXlName(idx):
    # excel의 columns no(1,2,3)을 letter(A,B,C)로 변환하는 함수 from stackoverflow

    if idx < 1:
        raise ValueError("Index is too small")
    result = ""
    while True:
        if idx > 26:
            idx, r = divmod(idx - 1, 26)
            result = chr(r + ord('A')) + result
        else:
            return chr(idx + ord('A') - 1) + result


def make_list_one(df):
    # excel에서 load한 한개 sheet의 df을 [row no,column no, cell value]의 column을 가지는 df으로 변환
    start_one = time.time()
    list_from_df = df.values.tolist()
    list_rcv = []
    row_no = 0
    for row_from_list in list_from_df:
        col_no = 0
        for value_from_row in row_from_list:
            if not (pd.isna(value_from_row)):
                list_rcv.append((int(row_no), int(col_no), str(value_from_row)))
            col_no += 1
        row_no += 1
    df_list = pd.DataFrame(list_rcv)
    df_list.columns = ['row', 'col', 'value']
    # st.write(round(time.time() - start_one, 3), '초')
    return df_list


def make_list_all(file_path, sheet_list):
    # make_list_one 함수를 반복 이용하여, 모든 sheet를 처리 하여, [sheet_name, row no,column no, cell value]의 column을 가지는 df으로 변환
    start_all = time.time()
    df_combined = pd.DataFrame()
    df_all = pd.read_excel(file_path, sheet_name=None, header=None)
    for sheet_name in sheet_list:
        df = df_all[sheet_name]
        df_list = make_list_one(df)
        df_list['sheet_name'] = sheet_name
        df_combined = df_combined.append(df_list)
        # print(sheet_name)

    df_combined = df_combined[['sheet_name', 'row', 'col', 'value']]
    df_combined = df_combined.reset_index(drop=True)
    # st.write(round(time.time() - start_all, 2), '초')
    return df_combined


def make_table(df_F, df_list, df_table, file_name):
    # make_list_all을 이용해 만든 각 input file을 하나의 table로 만듬
    # [sheet_name, row no,column no, file1, file2, filex~] 의 column을 가지는 df으로 변환
    df = pd.merge(df_F, df_list, on=['sheet_name', 'row', 'col', 'value'], how="outer", indicator=True)
    df = df[df._merge == 'right_only']
    df = df.drop(['_merge'], axis=1)
    df.columns = ['sheet_name', 'row', 'col', file_name]
    if df_table.shape[1] == 0:
        df_table = df
    else:
        df_table = pd.merge(df_table, df, how='outer', on=['sheet_name', 'row', 'col'])
    return df_table


def make_form_list(uploaded_file):
    # 양식 파일이 없을 경우, 데이터시트 파일들을 비교하며, 반복되는 부분을 양식이라 간주 함
    form_list = pd.DataFrame()
    count = 0
    for file in uploaded_file:
        xl = pd.ExcelFile(file)
        df = make_list_all(file, xl.sheet_names)
        if count == 0:
            form_list = df
        else:
            form_list = pd.merge(form_list, df, on=['sheet_name', 'row', 'col', 'value'], how="inner")
        count += 1
    return form_list


def make_final_table(uploaded_file):
    # make_table로 만든 df에 cell address, left1,left2,above 추가 하여 최종 df로 만듬
    data_table = pd.DataFrame()
    for file in uploaded_file:
        data_list = make_list_all(file, sheet_list)
        data_table = make_table(form_list, data_list, data_table, file.name)
    data_table.insert(1, 'address', None)
    data_table.insert(2, 'left1!', None)
    data_table.insert(3, 'left2!', None)
    data_table.insert(4, 'above!', None)

    for i in data_table.index:
        row = data_table.loc[i, 'row']
        col = data_table.loc[i, 'col']
        sheet = data_table.loc[i, 'sheet_name']
        address = str(ColIdxToXlName(int(col) + 1)) + str(int(row) + 1)
        data_table.loc[i, 'address'] = address
        left = list(form_list['value'][
                        (form_list.row == row) & (form_list.col < col) & (form_list.sheet_name == sheet)])
        above = list(form_list['value'][
                         (form_list.row < row) & (form_list.col == col) & (form_list.sheet_name == sheet)])
        if len(left) > 0: data_table.loc[i, 'left1!'] = left[-1]
        if len(left) > 1: data_table.loc[i, 'left2!'] = left[-2]
        if len(above) > 0: data_table.loc[i, 'above!'] = above[-1]
        # above로 표시 할 단어들을 제한하기 위한, 조건문 (ex. Rated / Normal / Min. 등등)
        # if above[-1].upper() in above_limit.upper():
        #     data_table.loc[i, 'above'] = above[-1]
    data_table = data_table.reset_index(drop=True)
    data_table.index.name = 'No!'
    data_table.rename(columns={'row': 'row!'}, inplace=True)
    data_table.rename(columns={'col': 'col!'}, inplace=True)
    return data_table


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
