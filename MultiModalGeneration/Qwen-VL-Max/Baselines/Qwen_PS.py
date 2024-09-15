import base64
import os

import dashscope
import requests
import pandas as pd
from openai import OpenAI
import re
import ast



def data_extraction(excel_file_path):
  # Load the Excel file
  df = pd.read_excel(excel_file_path)
  # Extract the desired columns，提取第
  image_id = df.iloc[:, 0].tolist()
  Target = df.iloc[:, 2].tolist()
  Source = df.iloc[:, 3].tolist()
  Text = df.iloc[:, 1].tolist()
  # Return the four lists
  return image_id, Target, Source, Text


# 生成答案
def answer_generation():
  # 模型回答的所有答案，告诉他上下文

  message = base_prompt
  message_content = (message + "This is an advertising metaphor.Please select the most likely source domain from the image and text based on the image and text as well as the target domain." +
                     "For each source domain, you can only choose the one you think is most likely.")
  messages = [
      {
          "role": "user",
          "content": [
              {"image":image_path},
              {"text": message_content +  "Let’s first understand the problem and devise a plan to solve the problem."
                    "Then, let’s carry out the plan and solve the problem step by step."}
          ]
      }
  ]
  try:
      response = dashscope.MultiModalConversation.call(model='qwen-vl-max',
                                                       messages=messages)
  except:
      response = None
  print('进入回答')
  try:
      answer = response["output"]["choices"][0]["message"]["content"][0]["text"]
  except:
      answer = "NULL"
  return answer




# 提取精确答案
def anwer_Extraction(answer):
  prompt = "Please extract the source domain according to the above answer and generate a dictionary. The format is: {'Source': 'source'}"
  prompt = prompt.replace("'", '"')
  message_content = (
              answer + prompt + "Requires only dictionaries to be returned, and the source domain should be the one you think is most likely.")
  messages = [
      {
          "role": "user",
          "content": [
              {"image": image_path},
              {"text": message_content}
          ]
      }
  ]
  try:
      response = dashscope.MultiModalConversation.call(model='qwen-vl-max',
                                                       messages=messages)
  except:
      response = None
  print('进入回答')
  try:
      answer = response["output"]["choices"][0]["message"]["content"][0]["text"]

  except:
      answer = "NULL"

  return answer

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
file_path = 'D:/PycharmProject/SourceGeneration/Output/Qwen/Qwen_PS_Meme.xlsx'

# 检查文件是否存在，如果不存在，创建并写入列名
if not os.path.exists(file_path):
    data = {'image_id': [], 'Source_generation': [],
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
  # 模型回答的所有答案，告诉他上下文
  answer = answer_generation()
  # 从荣誉的答案中提取精确答案转换为字典形式
  final_answer = anwer_Extraction(answer)
  # 将字符串转换为字典
  final_answer = final_answer.replace('\n','').strip('```').strip('python').strip('json')
  try:
    final_answer = ast.literal_eval(final_answer)
    source = final_answer['Source']


  except:
    print(f'error:转换为字典失败')
    source = final_answer

  count = count + 1
  # 创建一个DataFrame来存储当前迭代的数据
  data = {'image_id': [i], 'Source_generation': [source],
          }
  df = pd.DataFrame(data)
  # 追加数据到Excel文件中
  with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
    df.to_excel(writer, startrow=writer.sheets['Sheet1'].max_row, index=False, header=False)
