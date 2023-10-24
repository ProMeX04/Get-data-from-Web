import requests
import subprocess
from bs4 import BeautifulSoup
#  Danh sách các buổi
list_lesson = (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14,
               15, 16, 17, 18, 19, 20, 21, 23, 24, 25, 26)
#  Số lượng bài tập
num_exercise = (29, 20, 27, 0, 7, 17, 0, 6, 26, 13, 5, 10,
                11, 0, 20, 5, 6, 0, 0, 0, 0, 0, 0, 0, 258)

session_id = 'dfpgfnnkqsknb6ur16441oa6yq90hr32	'

#  Nhóm người cần check
Group = ["phanhhoccode", "nguyenvanchinh28", "anhkhoi16", "Nguyenduytan"]


def request(session_id) -> None:
    session = requests.Session()
    session.cookies.set('sessionid', session_id)
    response = session.get("https://laptrinh24h.vn/contest/cpp33/ranking/")
    return response

def GetData(response) -> list:
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
                                if (anchor.text[:2] == '10'):
                                    User[index_row][bt[1]] += 1
                                    cnt += 1
            if (index_row > -1):
                User[index_row]["Total"] = cnt
    return User

def WriteToCSV(path, Group, User) -> None:
    with open(path, "w", encoding="utf-8") as OutFile:
        OutFile.write("Rank,ID,")
        for i in list_lesson:
            OutFile.write(str(i)+",")
        OutFile.write("Total,")
        OutFile.write("\n")
        for x in User:
            if x["ID"] in Group:
                for index, j in enumerate(x):
                    if (x[j] != 0):
                        OutFile.write(" "+str(x[j]))
                        if index > 1:
                            OutFile.write('/'+str(num_exercise[index-2]))
                    OutFile.write(',')
                OutFile.write("\n")


if __name__ == "__main__":
    res = request(session_id)
    if res.status_code == 200:
        print("Success")
        data = GetData(res)
        path = "BXH.csv"
        WriteToCSV(path, Group, data)
        file_path = 'D:\DSA\Python\BXH.csv'
        subprocess.Popen(['start', file_path], shell=True)
    else:
        print("Can't connect")
