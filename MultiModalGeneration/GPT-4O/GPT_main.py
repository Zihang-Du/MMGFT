import base64
import json
import os

import requests
import pandas as pd
from openai import OpenAI
import re


# 从excel中提取数据：图片id,目标域，源域，text
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

# 得到模型回答
def answer_get(base_prompt, base64_image, prompt_1, prompt_2_1, prompt_2_2, prompt_3):
    message_content = (
            base_prompt +
            "I have several questions. Please answer them one by one in a detailed manner.\n\n" +
            "Question 1: " + prompt_1 + "\n\n" +
            "Question 2: " + prompt_2_1 + "\n\n" +
            "Question 3: " + prompt_2_2 + "\n\n" +
            "Question 4: " + prompt_3 + "\n\n" +
            "This is a chain of thought process. Let's think step by step."
    )
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
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
        "max_tokens": 1200,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # print(f'response_data:{response_data}')
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'answer{i}: {answer}')
    return answer


# 从所有答案里抽取实体
def Entity_Extraction(answer_all):
    # 创建请求
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": answer_all + (
                            "Here are the answers to several questions provided by the previous model. Based on these answers:"
                            "1. List all entities mentioned in the answers from each question.Do not ignore any objects that appear(Even if it's \"resemble entity\" )."
                            "2. Count the total number of times each entity appears in the four answers (record it once if it appears once)."
                            "3. For entities that describe multiple objects(or entities with brackets like this should separate the entities outside the brackets from those inside the brackets), split them into basic entities that describe a single object. For example, Tongue (cactus-like):should be split into Tongue and cactus."
                            "4. Do not split entities that describe the composition of the target domain (e.g., \"many countries and cities' names\" forming a flag should not be split)."
                        )
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "Do not extract brand names.Only provide the final candidate pool containing the entities and their final number of repetitions"
                    },
                ]
            }
        ],
        "max_tokens": 2500,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'first_Candidate_pool: {answer}')
    return answer


# 利用自一致性重复提取实体
def SC_Entity_Extraction(first_Candidate_pool, first_Candidate_pool_2, first_Candidate_pool_3):
    message = "[" + str(first_Candidate_pool) + "];" + "[" + str(first_Candidate_pool_2) + "];" + "[" + str(
        first_Candidate_pool_3) + "];"
    message_content = "Here are three outputs of the same question, each answer enclosed in '[]', separated by " + ";" + ". Use self-consistency to retain entities that appear in at least two of the answers. Only provide the final entities and their occurrence count (if the occurrence count differs, use the highest count among the answers)."
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            }
        ],
        "max_tokens": 400,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'SC_Entity_Extraction: {answer}')
    return answer


# 对第一个候选池子进行筛选，得到筛选后的池子
def Screening_pool(first_Candidate_pool, image):
    message_content = first_Candidate_pool + (
                "Here is the final candidate pool. Please process the candidate pool according to the following rules: "
                "1) Remove any entities that share identical words or semantically equivalent terms with the target domain : " + str(Target[i]) + "; "
                "2)Remove entities that is a brand name:"
                "3) Remove non-physical entities or abstract concepts. Only Keep tangible objects or all entities that appear in the image.   "
                "4) Merge highly similar entities (add their repetition counts after merging). "
                "Only strict adherence to these rules is allowed, without any additional operations.")
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{image}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 1200,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Second_Candidate_pool: {answer}')
    return answer


# 利用自一致性重复提取实体
def SC_Screening_pool(second_Candidate_pool, second_Candidate_pool_2, second_Candidate_pool_3):
    message = "[" + str(second_Candidate_pool) + "];" + "[" + str(second_Candidate_pool_2) + "];" + "[" + str(
        second_Candidate_pool_3) + "];"
    message_content = "Here are three outputs of the same question, each answer enclosed in '[]', separated by " + ";" + ". Keep entities that appear in the answers two or three times. "
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            }
        ],
        "max_tokens": 400,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'SC_Entity_Extraction: {answer}')
    return answer


# 对筛选后的池子进行打分排序
def score_candidate_pool_function(second_Candidate_pool, image, base_prompt):

    message = (
        "Please score the object based on the image and text provided, following these rules:"
        "\n\n2 points: The object in the diagram has a direct metaphorical or symbolic meaning with the target domain, or their positions completely overlap or form a logical combination."
        "\n\n1 point: The object has a specific indirect association or supportive role with the target domain within the context of the image and text."
        "\n\n0 points: The object has no relevant relation to the target domain in the image and text, or even conflicts with or causes confusion regarding the target domain."
        "\n\nUse the following formula to calculate the final score: (number of repetitions + score) / 6 to get the final score. "
    )
    message_content = base_prompt + message

    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": second_Candidate_pool
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{image}"
                        }
                    }
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "Ranked in descending order based on scores."
                    },
                ]
            }
        ],
        "max_tokens": 400,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    response_data = response.json()
    answer = response_data['choices'][0]['message']['content']
    print(f'Score_answer: {answer}')
    return answer


# 根据得分选择候选源域并进行属性选择
def Attribute_Selection(Score_Candidate_pool):
    message_content = (
                "This is the pool after scoring. "
                "Select the top three candidate source domains from it, ranked by score (if there are only one or two objects in the pool, then select all of them as the candidate source domains) "
                "For each selected candidate source domain, perform attribute/feature selection using external dictionaries. The results should meet the following requirements:The selected attributes should be properties of the candidate source domain."
                "The selected attributes must be relevant to the text: " + str(Text[i]) + "The best case is to be consistent with the attribute words used in the text to describe the characteristics of the target domain..")
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": Score_Candidate_pool
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "For each candidate source domain, select the top three attributes/features that you consider most relevant."
                    },
                ]
            }
        ],
        "max_tokens": 600,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Attribute_Selection: {answer}')
    return answer


# 从答案里选择喻底


# 形成三元组
def Triples_Generation(Attribute_Selection_result):
    message_content = Attribute_Selection_result + "These are the selected candidate source domains and their corresponding attributes."
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Form triples [(Candidate Source Domain), (Attributes), (Text: " + str(Text[
                                                                                                           i]) + ")]," + "The text do not change. Each candidate source field generates three triples."
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "Only the final triples are given"
                    },
                ]
            }
        ],
        "max_tokens": 600,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Triples: {answer}')
    return answer


# 对三元组进行排序
def Source_Computer(Triples):
    message = Triples
    message_content = "These are all the triples generated by all candidate source domains.Please rank these triples based on their cosine similarity."
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "Only provide the final ranking results."
                    },
                ]
            }
        ],
        "max_tokens": 600,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Source_Computer: {answer}')
    return answer


# 根据相似度排序进行源域选择
def Source_Selection(Source_Computer):
    message = Source_Computer
    message_content = "According to the above answers, Choose the source domain and ground that holds the top position in the ranking., and output only the source domain and attribute, and do not output others"
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "Return in dictionary format like this: {\"Source\": \"source\",\"Ground\":\"attribute\"}。Only give the dictionary content, nothing else"
                    },
                ]
            }
        ],
        "max_tokens": 60,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'source_answer:{answer}')
    return answer


# 根据相似度排序进行源域选择
def Source_Selection_Two(Source_Computer):
    message = Source_Computer
    # 根据之前的回答，选择排名最高的源域及其相关的两个最可能的喻底。然后，仅输出与该源域及其属性相关的信息，排除任何其他细节。
    message_content = "According to the previous answers, select the highest-ranking source domain along with the two most probable grounds associated with it from the ranking. Then, output only the information pertaining to this source domain and its attributes, excluding any other details. "
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": "Return in dictionary format like this: {\"Source\": \"source\",\"Ground1\":\"attribute1\",\"Ground2\":\"attribute2\"}。Only give the dictionary content, nothing else"
                    },
                ]
            }
        ],
        "max_tokens": 80,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'source_answer:{answer}')
    return answer


# 三元组加上目标域形成四元组生成解释
def Paraphrase_Generation(Source_Selection_result):
    message = Source_Selection_result
    message_content = "Create a quadruplet using the chosen source domain and attribute, coupled with the target domain and text: [Target: { " + \
                      Target[i] + "}, Source Domain: {Source Domain}, Ground: {Ground}, Text: {" + Text[i] + "}"
    message_generation = "This advertisement contains a metaphor. Based on the provided quadruple, interpret this metaphor and provide a paraphrase. The final response you furnish should be concise, not exceeding 50 words"
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_generation
                    },
                ]
            },
        ],
        "max_tokens": 100,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Paraphrase: {answer}')
    return answer


# 三元组加上目标域形成四元组生成解释
def Paraphrase_Generation_Two(Source_Selection_result):
    message = Source_Selection_result
    message_content = "Create a quadruplet using the chosen source domain and attribute, coupled with the target domain and text: [Target: { " + \
                      Target[i] + "}, Source Domain: {Source Domain}, Ground: ({Ground1,Ground2}), Text: {" + Text[
                          i] + "}"
    message_generation = "This advertisement contains a metaphor. Based on the provided quadruple, interpret this metaphor and provide a paraphrase. The final response you furnish should be concise, not exceeding 50 words"
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message
                    },
                ]
            },
            {
                "role": "assistant",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_generation
                    },
                ]
            },
        ],
        "max_tokens": 100,
        "temperature": 0  # 设置温度参数为 0，生成尽可能一致的输出
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Paraphrase: {answer}')
    return answer


#  根据答案精准提取源域
def Source_Extraction(Source_Selection_result):
    message = Source_Selection_result
    message_content = message + "Select the source domain from the answer, only give the source domain words"
    payload = {
        "model": "gpt-4o",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": message_content
                    },
                ]
            }
        ],
        "max_tokens": 400
    }
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    # 获取响应的 JSON 数据
    response_data = response.json()
    # 提取 content 字段
    answer = response_data['choices'][0]['message']['content']
    print(f'Source: {answer}')
    return answer


# OpenAI API Key
api_key = "YOUR-API-KEY"

# file_path = "D:/PycharmProject/SourceGeneration/data/ads_metaphor.xlsx"
excel_file_path = "D:/PycharmProject/SourceGeneration/data/MET-Meme_New/MET-Meme_New.xlsx"
images_id, Target, Source, Text = data_extraction(excel_file_path)
# prompt设计

prompt_1 = "What entities are present in the image and the text respectively? Please answer regarding objects."
prompt_2_1 = "In the image and text, what elements or features appear somewhat unusual or extraordinary"
prompt_2_2 =  "Based on answer 2, extract all the physical entities that appear and record the number of occurrences. Do not ignore any object that appears (even if it is a \"similar entity\")."
prompt_3 = "Does the target domain form an object or resemble an object based on the image ? If yes, only answer the object; if not, answer with no."




first_Candidate_pools = []
second_Candidate_pools = []
score_results = []
Attribute_Selection_results = []
Triples_all = []
Source_Computer_results = []
Sources = []
Grounds = []
Grounds1 = []
Grounds2 = []
Paraphrases = []
Paraphrases_2 = []
count = 0
# 构建图片地址

# 文件路径
file_path = 'D:/PycharmProject/SourceGeneration/Output/GPT_main_result_Meme.xlsx'

# 检查文件是否存在，如果不存在，创建并写入列名
if not os.path.exists(file_path):
    data = {'image_id': [], 'Source_generation': [],'Ground_generation': [],'Explanation_generation': [],'answers_all' : [],
            'first_Candidate_pools' : [] ,'second_Candidate_pools' : [], 'score_results' : [],'Attribute_Selection_results' :[],
            'Triples_all' :[], 'Source_Computer_results' :[]
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

    print(f'base_prompt: {base_prompt}')
    base64_image = encode_image(image_path)
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    # 模型回答的所有答案，告诉他上下文
    answer_all = answer_get(base_prompt, base64_image, prompt_1, prompt_2_1, prompt_2_2, prompt_3)

    # 根据答案提取所有实体
    first_Candidate_pool = Entity_Extraction(answer_all)
    # 利用规则筛选
    second_Candidate_pool = Screening_pool(first_Candidate_pool, base64_image)
   # 只要Final candidate pool后半部分的精简答案
    try:
        pattern = re.compile(r'Processed', re.IGNORECASE)
        second_Candidate_pool = "The final  " + re.split(pattern, second_Candidate_pool, maxsplit=1)[1].strip()
        print(f'提取后的second_Candidate_pool：{second_Candidate_pool}')
    except:
        print("不包含final candidate pool")

    # 对筛选后的池子进行打分排序
    score_result = score_candidate_pool_function(second_Candidate_pool, base64_image, base_prompt)

    # 按照得分从高到低选择 然后进行属性选择
    Attribute_Selection_result = Attribute_Selection(score_result)

    # 三元组生成
    Triples = Triples_Generation(Attribute_Selection_result)

    #  计算三元组相似度排序
    Source_Computer_result = Source_Computer(Triples)

    #  根据答案提取源域，以及喻底
    Source_result = Source_Selection(Source_Computer_result)
    try:
        Source_result_dic = json.loads(Source_result)
    except :
        print(f'第{i}次转化为字典失败')
        source = Source_result

    # 生成解释
    Paraphrase = Paraphrase_Generation(Source_result)
    # 根据两个喻底生成的解释
    #Paraphrase_2 = Paraphrase_Generation_Two(Source_result_two)
    # Source_result = Source_Extraction(Source_Computer_result)
    # Sources.append(Source_result)
    try:
        source = Source_result_dic['Source']
        Ground = Source_result_dic['Ground']
    except:
        print(f'source: {source}')
    # 创建一个DataFrame来存储当前迭代的数据
    data = {'image_id': [i], 'Source_generation': [source],'Ground_generation': [Ground],'Explanation_generation': [Paraphrase],
            'answers_all': [answer_all],
            'first_Candidate_pools': [first_Candidate_pool], 'second_Candidate_pools': [second_Candidate_pool], 'score_results': [score_result],
            'Attribute_Selection_results': [Attribute_Selection_result],
            'Triples_all': [Triples], 'Source_Computer_results': [Source_Computer_result]
            }
    df = pd.DataFrame(data)

    # 追加数据到Excel文件中
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
        df.to_excel(writer, startrow=writer.sheets['Sheet1'].max_row, index=False, header=False)
    count = count + 1

