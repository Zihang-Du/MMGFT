import base64
import os

import requests
import pandas as pd
from openai import OpenAI
import re
import ast


Manual_COT_CHINEESE = ['从图片和文本中提取所有实体及其特征。',
              '去除与目标域相同的实体。',
              '从实体中选取相关的候选源域和属性',
              '构建并根据相似度排序三元组',
              '选择最合适的源域,喻底,以及生成隐喻解释']


Manual_COT = ["Extract all entities and their features from the images and text."
              "Remove entities that are the same as the target domain."
              "To select relevant candidate source domains and attributes from entities."
              "Construct and rank the triples based on similarity."
              "Choose the most suitable source domain and attributes, and generate the metaphor explanation."
              ]

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
def answer_generation(Manual_COT):
  # 模型回答的所有答案，告诉他上下文
  Manual_COT_str = '\n'.join(Manual_COT)
  payload = {
    "model": "gpt-4o",
    "messages": [
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": "This is a multimodal advertising metaphor." + base_prompt
          }
        ]
      },
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": Manual_COT_str +"This is the set of steps I have designed for you. Please execute each step and return the results."
          },
          {
            "type": "image_url",
            "image_url": {
              "url": f"data:image/jpeg;base64,{base64_image}"
            }
          }
        ]
      }
    ],
    "max_tokens": 1000,
    "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
  }
  response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
  # 获取响应的 JSON 数据
  response_data = response.json()
  print(f'response_data:{response_data}')
  # 提取 content 字段
  answer = response_data['choices'][0]['message']['content']
  print(f'answer: {answer}')
  return  answer


# 提取精确答案
def anwer_Extraction(answer):
  payload = {
    "model": "gpt-4o",
    "messages": [
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": answer
          }
        ]
      },
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": "Please extract the source domain according to the above answer and generate a dictionary. The format is: {'Source': source}"
          }
        ]
      },
      {
        "role": "assistant",
        "content": [
          {
            "type": "text",
            "text": "Requires only dictionaries to be returned and the source domain should be the one you think is most likely."
          }
        ]
      },
    ],
    "max_tokens": 100,
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

Sources = []
Grounds = []
Paraphrases = []
# print(len(pic_id))
count = 0
file_path = 'D:/PycharmProject/SourceGeneration/Output/GPT_Manual_COT_result_Meme.xlsx'
# 检查文件是否存在，如果不存在，创建并写入列名
if not os.path.exists(file_path):
    data = {'image_id': [], 'Source_generation': [],'Ground_generation': [],'Explanation_generation': []
            }
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
  # 模型回答的所有答案，告诉他上下文
  answer = answer_generation(Manual_COT)
  # 从荣誉的答案中提取精确答案转换为字典形式
  final_answer = anwer_Extraction(answer)
  # 将字符串转换为字典
  final_answer = final_answer.replace('\n','').strip('```').strip('python')
  # final_answer = final_answer.replace(' ','')
  # 使用 ast.literal_eval 将字符串转换为字典
  try:
    final_answer = ast.literal_eval(final_answer)
    source = final_answer['Source']
    Ground = final_answer['Ground']
    Paraphrase = final_answer['Paraphrase']
  except:
    print(f'error:转换为字典失败')
    source = final_answer
    print(f'source: {source}')
    Ground = ''
    Paraphrase = ''
  count = count + 1
  # 创建一个DataFrame来存储当前迭代的数据
  data = {'image_id': [i], 'Source_generation': [source],'Ground_generation': [Ground],'Explanation_generation': [Paraphrase]
          }
  df = pd.DataFrame(data)
  # 追加数据到Excel文件中
  with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
    df.to_excel(writer, startrow=writer.sheets['Sheet1'].max_row, index=False, header=False)
# data = {'Source_generation': Sources, 'Ground_generation': Grounds, 'Paraphrase_generation': Paraphrases}
#
# # 使用字典创建一个DataFrame
# df = pd.DataFrame(data)
#
# # 将DataFrame写入Excel文件
# df.to_excel('D:/PycharmProject/SourceGeneration/Output/GPT_Manual_COT_result.xlsx', index=False)