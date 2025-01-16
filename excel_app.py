import streamlit as st
import pandas as pd

st.title("Ứng dụng xử lý file Excel")

uploaded_file = st.file_uploader("Chọn file Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Đọc file Excel với header ở hàng thứ 2 (index 1)
        df = pd.read_excel(uploaded_file, engine='openpyxl', header=1)
        st.success("Đọc file Excel thành công!")

        # Hiển thị dữ liệu
        st.subheader("Dữ liệu gốc")
        st.dataframe(df)

        # Các tùy chọn xử lý dữ liệu
        st.subheader("Tùy chọn xử lý")

        # Hiển thị thông tin cơ bản
        if st.checkbox("Hiển thị thông tin cơ bản về dữ liệu"):
            st.write("Số dòng:", df.shape[0])
            st.write("Số cột:", df.shape[1])
            st.write("Tên các cột:", list(df.columns))
            st.write("Kiểu dữ liệu của từng cột:")
            st.write(df.dtypes)

        # Hiển thị thống kê mô tả
        if st.checkbox("Hiển thị thống kê mô tả"):
            st.write(df.describe())

        # Lựa chọn cột để hiển thị
        st.subheader("Chọn cột để hiển thị")
        selected_columns = st.multiselect("Chọn các cột:", df.columns.tolist(), default=df.columns.tolist())
        if selected_columns:
            st.write(df[selected_columns])

        # Lọc dữ liệu
        st.subheader("Lọc dữ liệu")
        filter_column = st.selectbox("Chọn cột để lọc:", df.columns.tolist())
        filter_value = st.text_input(f"Nhập giá trị để lọc trong cột '{filter_column}':")
        if filter_value:
            try:
                filtered_df = df[df[filter_column].astype(str).str.contains(filter_value, case=False)]
                st.write(f"Dữ liệu sau khi lọc theo '{filter_column}' với giá trị '{filter_value}':")
                st.dataframe(filtered_df)
            except Exception as e:
                st.error(f"Lỗi khi lọc dữ liệu: {e}")

        # Tải xuống dữ liệu đã xử lý
        st.subheader("Tải xuống dữ liệu đã xử lý")
        if st.button("Tải xuống dữ liệu dưới dạng CSV"):
            csv_data = df.to_csv(index=False)
            st.download_button(
                label="Tải xuống CSV",
                data=csv_data,
                file_name="processed_data.csv",
                mime="text/csv",
            )

        if st.button("Tải xuống dữ liệu dưới dạng Excel"):
            excel_data = df.to_excel(index=False, engine='openpyxl')
            st.download_button(
                label="Tải xuống Excel",
                data=excel_data,
                file_name="processed_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"Lỗi khi đọc file Excel: {e}")