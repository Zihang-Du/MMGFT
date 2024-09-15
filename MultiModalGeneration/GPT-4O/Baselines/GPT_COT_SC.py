import base64
import os

import requests
import pandas as pd
from openai import OpenAI
import re
import ast



def data_extraction(file_path):
  # Load the Excel file
  df = pd.read_excel(file_path)
  # Extract the desired columns，提取第
  image_id = df.iloc[:, 0].tolist()
  Target = df.iloc[:, 2].tolist()
  Source = df.iloc[:, 3].tolist()
  Text = df.iloc[:, 1].tolist()

  # Return the four lists
  return image_id, Target, Source, Text

# 编码图像
def encode_image(image_path):
  with open(image_path, "rb") as image_file:
    return base64.b64encode(image_file.read()).decode('utf-8')

# 生成答案
def answer_generation():
  # 模型回答的所有答案，告诉他上下文
  payload = {
    "model": "gpt-4o",
    "messages": [
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": base_prompt
          },
          {
            "type": "image_url",
            "image_url": {
              "url": f"data:image/jpeg;base64,{base64_image}"
            }
          }
        ]
      },
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": "This is an advertising metaphor. "
                    "Please select the most likely source domain from the image and text based on the image and text as well as the target domain."
          }
        ]
      },
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": "For each source domain, you can only choose the one you think is most likely."
                    "Let's think about it step by step"
          }
        ]
      },
    ],
    "max_tokens": 300,
    "temperature": 0.6  # 设置温度参数为 0，生成尽可能一致的输出
  }
  response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
  # 获取响应的 JSON 数据
  response_data = response.json()
  print(f'response_data:{response_data}')
  # 提取 content 字段
  answer = response_data['choices'][0]['message']['content']
  print(f'answer: {answer}')
  return  answer


# 利用自一致性提取最可能的答案
def anwer_Extraction(answer1,answer2,answer3):
  payload = {
    "model": "gpt-4o",
    "messages": [
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": "[" + str(answer1) + "];[" + str(answer2) + "];[" + str(answer3) + "]."
          }
        ]
      },
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": "Here are three outputs of the same question, each answer enclosed in '[]', separated by "  +  "';'"    +".Using self-consistency, we retain the source domain. In the event of a tie, we employ a random selection method to determine the final answer."
          }
        ]
      },
    ],
    "max_tokens": 300,
    "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
  }
  response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
  # 获取响应的 JSON 数据
  response_data = response.json()
  # print(f'response_data:{response_data}')
  # 提取 content 字段
  final_answer = response_data['choices'][0]['message']['content']
  print(f'SC answer: {final_answer}')
  return final_answer




def final_anwer_Extraction(SC_answer):
  prompt = "Your answer is returned as a dictionary: {'Source': 'Source'}"
  # 将单引号替换为双引号
  prompt = prompt.replace("'", '"')
  payload = {
    "model": "gpt-4o",
    "messages": [
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": SC_answer
          }
        ]
      },
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": "This is the answer obtained through self-consistency. Please extract the accurate source domain from this answer"
          }
        ]
      },
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": prompt+"Only dictionaries are allowed to be returned, nothing else is allowed."
          }
        ]
      },
    ],
    "max_tokens": 200,
    "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
  }
  response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
  # 获取响应的 JSON 数据
  response_data = response.json()
  # print(f'response_data:{response_data}')
  # 提取 content 字段
  final_answer = response_data['choices'][0]['message']['content']
  print(f'final answer: {final_answer}')
  return final_answer

# OpenAI API Key
api_key = "YOUR-API-KEY"
# file_path = "D:/PycharmProject/SourceGeneration/data/ads_metaphor.xlsx"
excel_file_path = "D:/PycharmProject/SourceGeneration/data/MET-Meme_New/MET-Meme_New.xlsx"
# 2405条数据
# images_id, images_url, Target, Source, Text = data_extraction(file_path)
images_id, Target, Source, Text = data_extraction(excel_file_path)


count = 0
Sources = []
Grounds = []
Paraphrases = []
# print(len(pic_id))
file_path = 'D:/PycharmProject/SourceGeneration/Output/GPT_COT_SC_result_Meme.xlsx'

# 检查文件是否存在，如果不存在，创建并写入列名
if not os.path.exists(file_path):
    data = {'image_id': [], 'Source_generation': []}
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

# 如果文件存在，读取已有的数据并找到最后一个处理过的ID
last_processed_id = None
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    last_processed_id = df['image_id'].iloc[-1] if not df.empty else None
# 开始循环，从上一个ID之后开始处理
start_id = last_processed_id + 1 if last_processed_id is not None else 0
for i in range(start_id, len(images_id)):
  # if i == 0 or i ==1 or i ==2 or i ==3 or i ==4 or i ==5 :
  #   continue
  if count == 1:
    break
  # 检查ID是否已经在Excel文件中
  if i in df['image_id'].values:
    print(f"ID {i} 已经处理过，跳过...")
    continue
  image_path = "D:/PycharmProject/SourceGeneration/data/MET-Meme_New/" + str(i) + ".jpg"
  print(f'image_path: {image_path}')
  base_prompt = "The text in the image is " + str(Text[i]) + ", and the target domain is " + str(Target[i])

  print(f'new_base_prompt: {base_prompt}')
  base64_image = encode_image(image_path)
  headers = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {api_key}"
  }
  # 模型回答的所有答案，告诉他上下文,问三次
  answer1 = answer_generation()
  answer2 = answer_generation()
  answer3 = answer_generation()

  # 利用自一致性对答案进行提取
  SC_answer = anwer_Extraction(answer1,answer2,answer3)

  # 从自一致性抽取里面的答案抽取最终答案
  final_answer = final_anwer_Extraction(SC_answer)
  # 将字符串转换为字典
  final_answer = final_answer.replace('\n','').strip('```').strip('python').strip('json')
  try:
    final_answer = ast.literal_eval(final_answer)
    source = final_answer['Source']
    Ground = final_answer['Ground']
  except:
    print(f'error:转换为字典失败')
    source = final_answer
    Ground = ''


  # 创建一个DataFrame来存储当前迭代的数据
  data = {'image_id': [i], 'Source_generation': [source],'Ground_generation': [Ground]
          }
  df = pd.DataFrame(data)
  # 追加数据到Excel文件中
  with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
    df.to_excel(writer, startrow=writer.sheets['Sheet1'].max_row, index=False, header=False)
