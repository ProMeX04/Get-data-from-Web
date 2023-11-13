"""Cài đặt python và Cài đặt các thư viện phía dưới mới có thể sử dụng được"""
import subprocess
from tkinter import *
import requests
from bs4 import BeautifulSoup
# Danh sách các buổi
list_lesson = (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14,
               15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26)
# Số lượng bài tập
num_exercise = (29, 20, 27, 0, 7, 17, 0, 6, 26, 13, 5, 10,
                19, 0, 20, 6, 6, 10, 0, 6, 0, 0, 11, 0, 5, 258)

# Session ID của bạn (cần được cung cấp)
session_id = 'fe8ijbpm7v0r6hookq4coniwnqmjksdg'

# Đường dẫn lưu trữ dữ liệu dưới dạng CSV
__path__ = "BXH.csv"

# Danh sách các thành viên cần kiểm tra (có thể thay đổi mặc định theo nhóm)
group = ["alisegaunt", "Hoanghieu2603", "", "", "kienbnvn2004", "", "", "",
         "", "", "", "", "", "", "nguyenvanchinh28", "Nguyenduytan",
         "phanhhoccode", "anhkhoi16", "",
         "Tram_le2003", "huynhduong152", "", "QuocThang",
         "nguyenduchuong300903fullhouse", "", "nhi9999", "",
         "gia_huy922004", "daoduykhai77", "HoanganhIT2K4",
         "anhquan", "", "",
         "Phuong_10", "Huongphanminh", "giaudeptrai", "Ducthinh20",
         "ninhlong21", "duongngan1810", "duongngan1810"]


def user_interface():
    """function show windows for user"""
    window = Tk()
    window.geometry("500x100")
    window.title("Lấy dữ liệu bài tập")
    label = Label(window, text="Nhập ID hoặc danh sách ID cần kiểm tra cách nhau bởi dấu cách", 
                font="calibri")
    label.pack(side='top')
    entry = Entry(window, width=100, foreground="green")
    entry.pack(side=TOP)

    def get():
        if (entry.get().split()):
            global group
            group = entry.get().split()
        window.destroy()
    button = Button(window, command=get, text="Collect", font="calibri")
    button.pack(side=TOP)
    window.mainloop()


def request(session_id) -> None:
    """_summary_

    Args:
        session_id (_string_): _session id from web_

    Returns:
        _type_: _string_
    """
    session = requests.Session()
    session.cookies.set('sessionid', session_id)
    response = session.get("https://laptrinh24h.vn/contest/cpp33/ranking/")
    return response


def get_data(response) -> list:
    res = response.text
    soup = BeautifulSoup(res, 'html.parser')
    table = soup.find('table')
    User = []
    if table:
        for index_row, row in enumerate(table.find_all("tr"), start=-1):
            cnt = 0
            for index_column, column in enumerate(row.find_all("td")):
                if index_column == 0:
                    User.append({"rank": column.text})
                elif index_column == 2:
                    User[index_row]["ID"] = column.find('a').text
                    for i in list_lesson:
                        User[index_row][str(i)] = 0
                elif index_column > 2:
                    anchor = column.find('a')
                    if anchor:
                        href = anchor.get('href')
                        if href:
                            bt = href.split(
                                '/')[-2].rstrip("cpp").split('b')
                            if len(bt) == 2:
                                if anchor.text[:2] == '10':
                                    User[index_row][bt[1]] += 1
                                    cnt += 1
            if index_row > -1:
                User[index_row]["Total"] = cnt
    return User


def write_to_csv(__path__, group, User) -> None:
    with open(__path__, "w", encoding="utf-8") as OutFile:
        OutFile.write("Rank,ID,")
        for i in list_lesson:
            OutFile.write("Lesson "+str(i)+",")
        OutFile.write("Total,")
        OutFile.write("\n")
        for ID in group:
            for x in User:
                if x["ID"] == ID:
                    for index, j in enumerate(x):
                        if (x[j] != 0):
                            OutFile.write(" "+str(x[j]))
                            if index > 1:
                                OutFile.write('/'+str(num_exercise[index-2]))
                        OutFile.write(',')
            OutFile.write("\n")


if __name__ == "__main__":
    # bật / tắt giao diện nhập
    user_interface()
    res = request(session_id)
    if res.status_code == 200:
        print("Success")
        data = get_data(res)
        write_to_csv(__path__, group, data)
        subprocess.Popen(['start', __path__], shell=True)
    else:
        print("Can't connect")
