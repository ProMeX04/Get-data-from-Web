""" main """
import subprocess
import requests
import pandas as pd
import source
# import openpyxl


def request(url):
    """Get response html web"""
    session = requests.Session()
    session.cookies['sessionid'] = source.SESSION_ID
    response = session.get(url)
    return response


def get_data_frame(response):
    """_summary_
    """
    df = pd.read_html(response.text)[0]
    # Read Table from web
    df = df.drop(columns=["Organization"])
    df = df.drop(columns=["Points"])

    # Delete column Organization
    for i in range(1, 259):
        head = str(i) + "  10"
        df[head] = df[head].astype(str).str[0:2]

    # print(df)
    new_df = pd.DataFrame()
    so_hoc_vien = df.shape[0]

    # print(df.head().columns)
    for i in range(so_hoc_vien):
        start_lesson = 2
        row = [i+1, df.iloc[i, 1].split()[0]]

        total = 0
        for lesson in source.LESSON:
            accepted = (
                df.iloc[i, start_lesson: start_lesson + lesson] == "10").sum()
            total += accepted
            if accepted:
                row.append(f"{accepted}/{lesson}")
            else :
                row.append("")
            start_lesson += lesson
        row.append(f"{total}/{source.TOTAL}")
        new_df = pd.concat(
            [new_df, pd.DataFrame([row])], ignore_index=True)

    new_df = new_df.rename(columns={0: "Rank", 1: "Name", 27: "Total"})
    return new_df


def dataframe_to_excel(df):
    """write and open excel file"""
    df.to_excel(source.EXCEL, index=False)
    subprocess.Popen(["start", source.EXCEL], shell=True)


if __name__ == "__main__":
    res = request(source.URL)
    if res.status_code == 200:
        print("Success")
        dataframe = get_data_frame(res)
        dataframe_to_excel(dataframe)
    else:
        print("Can't connect")
