import requests
import json
import base64
from openpyxl import load_workbook
import csv
from datetime import datetime


url = 'http://forms.product.in.ua/index.php/admin/remotecontrol'
login = 'admin'
password = 'ghjuhfvb'




# def jsonRequest(url, payload):
# 	headers = {'content-type': 'application/json'}
# 	r = requests.post(url, data=json.dumps(payload), headers = headers)
# 	return json.loads(r.text)

def json_request(j_url, j_payload):
    headers = {'content-type': 'application/json'}
    r = requests.post(j_url, data=json.dumps(j_payload), headers=headers)
    return r.json()


def get_session_key(m_url):
    payload = {'method': 'get_session_key', 'params': (login, password), 'id': 1}
    j = json_request(m_url, payload)
    return j["result"]


def release_session_key(m_url, key):
    payload = {"method": "release_session_key", "params": key, "id": 1}
    j = json_request(m_url, payload)
    return j["result"]


def export_responses(m_url, key, sid):
    payload = {"method": "export_responses", "params": (key, sid, "csv", "ru", "full"), "id": 1}
    j = json_request(m_url, payload)
    return base64.b64decode(j["result"])


def list_questions(m_url, key, sid):
    payload = {"method": "list_questions", "params": (key, sid), "id": 1}
    j = json_request(m_url, payload)
    return j["result"]

def date_format(datetime_string):
    return datetime.strptime(datetime_string, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')

def to_xlsx(row_data):
    wb = load_workbook(filename='template2.xlsx')
    ws = wb.active
    ws['E2'] = date_format(row_data['submitdate'])
    ws['A2'] = row_data['Q00']
    ws['A6'] = row_data['Q01']
    ws['A9'] = date_format(row_data['Q02'])
    ws['B9'] = row_data['Q06']
    if row_data['Q07'] == 'Y':
        ws['D9'] = row_data['Q06']
    else:
        ws['D9'] = row_data['Q09']
    ws['A14'] = row_data['Q10']
    if (row_data['Q13']+row_data['Q14']) == "A1":
        ws['B14'] = "женат/замужем"
    else:
        ws['B14'] = "холост/не замужем"

    ws['D14'] = row_data['Q17']


    # datetime_submit = datetime.strptime(row_data['submitdate'], '%Y-%m-%d %H:%M:%S')
    # print(datetime_submit.strftime('%d/%m/%Y'))


    #ws.append([1, 2, 3])
    wb.save("sample.xlsx")

#
# session_key = get_session_key(url)
#
#
#
# questions_list = list_questions(url, session_key, "563799")
#
# print(questions_list)
#
#
# for question in questions_list:
#     qid = question["id"]
#     title = question["title"]
#     question_text = question["question"]
#     print(title + "    " + question_text)
#
# my_csv = export_responses(url, session_key, "563799").decode("utf-8")
# print(my_csv)
# f = open("export.csv", "w")
# f.write(my_csv)
# f.close()
#
# release_session_key(url, session_key)
#
# lines = my_csv.splitlines(True)
# keys = lines[0]


reader = csv.DictReader(open("export.csv"))
for row in reader:
    if row['id'] == '56':
        print(row['submitdate'] + "  " + row['Q00'])
        to_xlsx(row)


