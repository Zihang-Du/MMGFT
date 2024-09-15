import base64
import os

import dashscope
import requests
import pandas as pd
from openai import OpenAI
import re
import ast
# import dashscope



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
def answer_generation(Manual_COT_str):
  # 模型回答的所有答案，告诉他上下文

  message_content = "This is a multimodal advertising metaphor." + base_prompt +"." + Manual_COT_str +".And this is the set of steps I have designed for you. Please execute each step and return the results." + "Let's think about it step by step"
  print(f'message content: {message_content}')
  messages = [
      {
          "role": "user",
          "content": [
              {"image":image_path},
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
      # print(f'对话后的答案是{answer}')
      # print(f'answer:{answer}')
  except:
      answer = "NULL"
  # answer = response[output.choices[x]][0]['message']['content']
  return answer




# 提取精确答案
def anwer_Extraction(answer):
  prompt = "Please extract the source domain according to the above answer and generate a dictionary. The format is: {'Source': source}"
  prompt = prompt.replace("'", '"')
  message_content = (
              answer + prompt + "Requires only dictionaries to be returned and the source domain should be the one you think is most likely.")
  print(f'message content2: {message_content}')
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
      # print(f'对话后的答案是{answer}')
      # print(f'answer:{answer}')
  except:
      answer = "NULL"
  # answer = response[output.choices[x]][0]['message']['content']
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
Manual_COT_CHINEESE = ['从图片和文本中提取所有实体及其特征。',
              '去除与目标域相同的实体。',
              '从提取后的实体中选取相关的候选源域和特征',
              '构建并根据相似度排序三元组',
              '选择最合适的源域']

Manual_COT = ["Extract all entities and their features from the images and text."
              "Remove entities that are the same as the target domain."
              "Select relevant candidate source domains and features from the extracted entities",
              "Construct and rank the triples based on similarity."
              "Choose the most suitable source domain."
              ]

file_path = 'D:/PycharmProject/SourceGeneration/Output/Qwen/Qwen_Manual_COT_Meme.xlsx'

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
  Manual_COT_str = str(Manual_COT)
  answer = answer_generation(Manual_COT_str)
  # 从答案中提取精确答案转换为字典形式
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
