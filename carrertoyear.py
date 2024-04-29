import pandas as pd

def main():
    # Đọc dữ liệu từ file Excel
    file_path = r'D:\stockcode\code2\data.xlsx'
    df = pd.read_excel(file_path)

    # Sắp xếp dữ liệu dựa trên cột 'Category Name'
    df.sort_values(by=['Category Name'], inplace=True)

    # Tạo một Writer object để ghi vào tệp Excel
    with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:

        df.to_excel(writer, sheet_name='1', index=False)

    print("Done:", file_path)

if __name__ == "__main__":
    main()
