import pandas as pd
import openpyxl
import requests
import json
import xlrd
import xlsxwriter
import time
# 1. 读取 Excel 文件中的问题
def read_questions_from_excel(file_path, sheet_name, question_column):
    """
    从 Excel 文件中读取问题
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param question_column: 问题所在的列名
    :return: 问题列表
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    questions = df[question_column].tolist()
    return questions

#  2. 调用 Kimi API 获取答案
def get_answers_from_kimi(questions, api_key):
    """
    调用 Kimi API 获取答案
    :param questions: 问题列表
    :param api_key: Kimi API Key
    :return: 答案列表
    """
    url = "https://api.moonshot.cn/v1/chat/completions"  # 替换为实际的 Kimi API URL
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    answers = []
    for question in questions:
        payload = {
            "model": "moonshot-v1-8k",  # 替换为实际的模型名称
            "messages": [{"role": "user", "content": question}],
            "temperature": 0
        }
        response = requests.post(url, json=payload, headers=headers)
        if response.status_code == 200:
            data = response.json()
            answer = data.get("choices", [{}])[0].get("message", {}).get("content", "")
            answers.append(answer)
        else:
            answers.append("Error: " + str(response.status_code))
    return answers

# 3. 将答案写回 Excel 文件
def write_answers_to_excel(file_path, sheet_name, answers, answer_column):
    """
    将答案写回 Excel 文件
    :param file_path: Excel 文件路径
    :param sheet_name: 工作表名称
    :param answers: 答案列表
    :param answer_column: 答案要写入的列名
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df[answer_column] = answers
    df.to_excel(file_path, sheet_name=sheet_name, index=False)

# 配置参数
excel_file = "测试 kimi 调用.xlsx"  # Excel 文件路径
sheet_name = "Sheet1"  # 工作表名称
question_column = "问题"  # 问题所在的列名
answer_column = "答案"  # 答案要写入的列名
api_key = "sk-KrQrYWB2ew9pmZt6LqUGIvs2z4ksFLJfa59MXFm1ynsba1kr"  # 替换为你的 Kimi API Key

# 读取问题
questions = read_questions_from_excel(excel_file, sheet_name, question_column)
print("读取到的问题：", questions)

# 获取答案
answers = get_answers_from_kimi(questions, api_key)
print("获取到的答案：", answers)

# 写回答案
write_answers_to_excel(excel_file, sheet_name, answers, answer_column)
print("答案已写回 Excel 文件。")


