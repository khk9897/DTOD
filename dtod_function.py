import pandas as pd
# import streamlit as st


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
    # st.write('make_list_one')
    # excel에서 load한 한개 sheet의 df을 [row no,column no, cell value]의 column을 가지는 df으로 변환
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
    # st.write(df_list)
    # st.write(len(df_list))
    if len(df_list) > 0:
        df_list.columns = ['row', 'col', 'value']
    # st.write(round(time.time() - start_one, 3), '초')
    return df_list


def make_list_all(file_path, sheet_list):
    # st.write('make_list_all')
    # make_list_one 함수를 반복 이용하여, 모든 sheet를 처리 하여, [sheet_name, row no,column no, cell value]의 column을 가지는 df으로 변환
    # start_all = time.time()
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
    # st.write('make_table')
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


def make_final_table(uploaded_file,form_list):
    # make_table로 만든 df에 cell address, left1,left2,above 추가 하여 최종 df로 만듬
    data_table = pd.DataFrame()
    for file in uploaded_file:
        xl = pd.ExcelFile(file)
        data_list = make_list_all(file, xl.sheet_names)
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
