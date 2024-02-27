import streamlit as st
import pandas as pd
import openpyxl
import io
from datetime import datetime


now = datetime.now()
## need improvment##
## Ahmed Alsrehy##

st.set_page_config(layout="wide", page_title="Beta")

file_upload = st.sidebar.file_uploader(
    "**البحث عن أسرة بالعينة.xlsx**", type=['xlsx'])
if file_upload is not None:
    df = pd.read_excel(file_upload)

    # st.write(df)
    df = df.iloc[1:, :]

    supervisor = 'مشرف'
    inspector = 'مفتش'
    researcher = 'باحث'
    su_status = 'حالة الإستمارة'
    occ = 'وصف حالة المعاينة'
    researcher_name = 'إسم الباحث'
    occ_code = 'رمز المعاينة'
    sample_id = 'معرف العينة'

    df[supervisor] = df[supervisor].astype(int)
    df[inspector] = df[inspector].astype(int)
    df[researcher] = df[researcher].astype(int)
    df[sample_id] = df[sample_id].astype(int)

    code_list = []
    df_code = df
    df_code = df_code.dropna(subset=occ_code)
    for i in sorted(df_code[occ_code].unique()):
        df_code1 = df_code[df_code[occ_code] == i]
        code_list.append({
            "رمز": df_code1[occ_code].unique()[0],
            "الوصف": df_code1[occ].unique()[0]
        })
    code_table = pd.DataFrame(code_list)
    col1, col2 = st.columns(2)
    with col1:
        with st.expander("**رموز المعاينة - إضافة**"):
            st.dataframe(code_table, use_container_width=True, hide_index=True)

    df2 = df

    tab1, tab2 = st.tabs(['عام', 'مراجعة للمكتملة'])
    with tab1:

        col1, col2, col3 = st.columns(3)
        with col1:
            inspector_select = st.selectbox(
                "المفتش", sorted(df[inspector].unique()))
            if len(df[inspector].unique()) > 0:
                if st.toggle("تفعيل فلتر المفتش",):
                    df = df[df[inspector] == inspector_select]
        with col2:
            researcher_select = st.selectbox(" الباحث", sorted(
                df[researcher].unique()), disabled=True)
            togg = st.toggle("تفعيل فلتر الباحث", disabled=True)
            if togg:
                df = df[df[researcher_name] == researcher_select]
        with col3:
            dd = st.multiselect("حالة الإكتمال", df[su_status].unique(
            ), df[su_status].unique(), disabled=True)
            df = df[df[su_status].isin(dd)]

        per_list = []
        for i in sorted(df[inspector].unique()):
            df_i = df[df[inspector] == i]
            for i in sorted(df_i[researcher].unique()):
                df_r = df_i[df_i[researcher] == i]

                full = df_r[df_r[su_status] == 'مكتمل']
                uncomplate = df_r[df_r[su_status] == 'غير مكتمل']
                new = df_r[df_r[su_status] == 'جديد']
                total = len(full) + len(uncomplate) + len(new)

                per_list.append({
                    "المشرف": df_r[supervisor].unique()[0],
                    'المفتش': df_r[inspector].unique()[0],
                    "الباحث": df_r[researcher].unique()[0],
                    "إسم الباحث": df_r[researcher_name].unique()[0],
                    # "الإجمالي": total,
                    "الإنتاجية": round((len(full) / total)*100, 2),
                    "مكتمل": len(full),
                    "غير مكتمل": len(uncomplate),
                    "جديد": len(new)

                })

        per_data = pd.DataFrame(per_list)
        per_data = per_data.iloc[:, ::-1]

        llist = []
        for i in sorted(df[inspector].unique()):
            df_i = df[df[inspector] == i]
            for i in sorted(df_i[researcher].unique()):
                df_r = df_i[df_i[researcher] == i]

                code1 = df_r[df_r[occ_code] == 1]
                code2 = df_r[df_r[occ_code] == 2]
                code3 = df_r[df_r[occ_code] == 3]
                code4 = df_r[df_r[occ_code] == 4]
                code5 = df_r[df_r[occ_code] == 5]
                code6 = df_r[df_r[occ_code] == 6]
                code7 = df_r[df_r[occ_code] == 7]
                code8 = df_r[df_r[occ_code] == 8]
                code9 = df_r[df_r[occ_code] == 9]
               
                llist.append({
                    "المشرف": int(df_r[supervisor].unique()[0]),
                    "المفتش": int(df_r[inspector].unique()[0]),
                    "الباحث": int(df_r[researcher].unique()[0]),
                    "أعطت كامل البيانات": len(code1),
                    "اعطت بعض البيانات": len(code2),
                    "يطلب الزيارة في وقت آخر": len(code3),
                    "المستوجب رفض الإستجابة": len(code4),
                    "لايوجد فرد مؤهل للرد": len(code5),
                    "الأسرة غير متواجدة وقت الزيارة (المسكن مغلق)": len(code6),
                    "المسكن خالي": len(code7),
                    "المسكن تحت التشييد أو الترميم أو مهدوم أو ازيل": len(code8),
                    "أخرى": len(code9),
                   


                })

        dash = pd.DataFrame(llist)
        dash = dash.iloc[:, ::-1]

        with st.expander("**الإنتاجية**", expanded=True):

            st.data_editor(
                per_data,
                column_config={
                    "الإنتاجية": st.column_config.ProgressColumn(
                        "الإنتاجية",
                        help="None",
                        format="%f",
                        min_value=0,
                        max_value=100,
                    ),
                },
                hide_index=True, use_container_width=True
            )

        with st.expander("**حالة المعاينة**", expanded=True):
            st.markdown(dash.style.hide(axis="index").to_html(),
                        unsafe_allow_html=True)

    with tab2:
        col1, col2, col3 = st.columns(3)
        with col1:
            inspector_select = st.selectbox(
                "المفتش ", sorted(df2[inspector].unique()))
            if len(df2[inspector].unique()) > 0:
                if st.toggle("تفعيل فلتر المفتش ",):
                    df2 = df2[df2[inspector] == inspector_select]

        df3 = df2[df2[su_status].isin(['مكتمل'])]
        df3 = df3[df3[occ].isin(
            ['يطلب الزيارة في وقت أخر', 'الأسرة غير متواجدة وقت الزيارة (المسكن مغلق)', "لا يوجد فرد مؤهل لإعطاء البيانات"])]
        id_table = []
        for i in df3[sample_id].unique():
            df2_id = df3[df3[sample_id] == i]
            id_table.append({
                "المشرف": df2_id[supervisor].unique()[0],
                "المفتش": df2_id[inspector].unique()[0],
                "الباحث": df2_id[researcher].unique()[0],
                "حالة المعاينة": df2_id[occ].unique()[0],
                "رقم العينة": round(df2_id[sample_id].unique()[0]),
                "حالة الإكتمال": df2_id[su_status].unique()[0]

            })
        id_data = pd.DataFrame(id_table)
        id_data = id_data.iloc[:, ::-1]

        col1, col2 = st.columns(2)
        with st.expander("Table_1", expanded=True):
            if len(id_data) > 0:
                st.dataframe(id_data.style.format(
                    thousands="", precision=0), hide_index=True, use_container_width=True)
            else:
                st.title("لاتوجد بيانات")

    buffer = io.BytesIO()
    file = "التقرير - نسخة أولية{}.XLSX".format(now.strftime("%d-%m-%Y"))
    with pd.ExcelWriter(buffer) as writer:
        per_data = per_data.iloc[:, ::-1]
        per_data.to_excel(
            writer, sheet_name='الإنتاجية', index=False)
        dash = dash.iloc[:, ::-1]
        dash.to_excel(
            writer, sheet_name='حالات المعاينة', index=False)
        if len(id_data) > 0:
            id_data.to_excel(
                writer, sheet_name='المكتملة', index=False)

    with st.sidebar.expander("Download", expanded=True):
        btn = st.download_button(
            label=":open_file_folder: xlsx - تحميل التقرير",
            data=buffer,
            file_name=file,
            mime="application/vnd.ms-excel"
        )
        if btn:
         
            st.success("تم التحميل ")
