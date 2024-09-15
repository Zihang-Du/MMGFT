import base64
import json
import os

import dashscope
import pandas as pd
import requests


# 从excel中提取数据：图片id,目标域，源域，text
def data_extraction(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path)

    # # Extract the desired columns，提取第
    # image_id = df.iloc[:, 0].tolist()
    # image_url = df.iloc[:, 1].tolist()
    # Target = df.iloc[:, 3].tolist()
    # Source = df.iloc[:, 4].tolist()
    # Text = df.iloc[:, 6].tolist()

    # Extract the desired columns，提取第
    image_id = df.iloc[:, 0].tolist()
    Text = df.iloc[:, 1].tolist()
    Target = df.iloc[:, 2].tolist()
    Source = df.iloc[:, 3].tolist()
    # Return the four lists
    return image_id, Target, Source, Text



# 得到模型回答
def answer_get(base_prompt, image_path, prompt_1, prompt_2_1, prompt_2_2, prompt_3):
    messages = [
        {
            "role": "user",
            "content": [
                {"image": image_path},
                {"text": (
                        base_prompt +
                        "I have several questions. \n\n" +
                        "Question 1: " + prompt_1 + "\n\n" +
                        "Question 2: " + prompt_2_1 + "\n\n" +
                        "Question 3: " + prompt_2_2 + "\n\n" +
                        "Question 4: " + prompt_3 + "\n\n" +
                        "This is a chain of thought process. Let's think step by step."
                )}
            ]
        }
    ]
    try:
        response = dashscope.MultiModalConversation.call(model='qwen-vl-max',
                                                         messages=messages)
    except:
        response = None
    print('进入回答')
    print(response)
    try:
        answer = response["output"]["choices"][0]["message"]["content"][0]["text"]
        # print(f'对话后的答案是{answer}')
        # print(f'answer:{answer}')
    except:
        answer = "NULL"
    # answer = response[output.choices[x]][0]['message']['content']
    return answer


# 从所有答案里抽取实体
def Entity_Extraction(answer_all):
    messages = [
        {
            "role": "user",
            "content": [
                {"text": answer_all + (
                    "Here are the answers to several questions provided by the previous model. Based on these answers:"
                    "1. List all entities mentioned in the answers from each question.Do not ignore any objects that appear(Even if it's \"resemble entity\" )."
                    "2. Count the total number of times each entity appears in the four answers (record it once if it appears once)."
                ) + "Do not extract brand names.Only provide the final candidate pool containing the entities and their final number of repetitions"}
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
    answer = response_data["output"]["choices"][0]["message"]["content"][0]["text"]
    print(f'SC_Entity_Extraction: {answer}')
    return answer


# 对第一个候选池子进行筛选，得到筛选后的池子
def Screening_pool(first_Candidate_pool, image_path):
    messages = [
        {
            "role": "user",
            "content": [
                {"image": image_path},
                {"text": first_Candidate_pool + (
                        "Here is the final candidate pool. Please process the candidate pool according to the following rules: "
                        "1) Remove any entities that share identical words or semantically equivalent terms with the target domain : " + str(Target[i]) + "; "
                                 "2)Remove entities that is a brand name:"
                                 "3) Remove non-physical entities or abstract concepts. Only Keep tangible objects or all entities that appear in the image.   "
                                 "4) Merge highly similar entities (add their repetition counts after merging). " + "Let's think about it step by step"
                                 #"Only strict adherence to these rules is allowed, without any additional operations."
                    )
                 }
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
        print(f'对话后的答案是{answer}')
        # print(f'answer:{answer}')
    except:
        answer = "NULL"
    # answer = response[output.choices[x]][0]['message']['content']
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
    answer = response_data["output"]["choices"][0]["message"]["content"][0]["text"]
    print(f'SC_Entity_Extraction: {answer}')
    return answer


# 对筛选后的池子进行打分排序
def score_candidate_pool_function(second_Candidate_pool, image_path, base_prompt):
    #  "\n\n2 分：图示中的对象与目标领域具有直接的隐喻或象征意义，或者它们的位置完全重叠或形成逻辑上的组合。"
    #         "\n\n1 分：对象在图片和文本的上下文中与目标领域具有特定的间接关联或支持作用。"
    #         "\n\n0 分：对象与目标领域在图片和文本中没有相关关系，甚至与目标领域产生冲突或混淆。"

    message = (
        "Please rate the given object based on the image and context information provided, following these rules:"
        "\n\n2 points: The object in the diagram has a direct metaphorical or symbolic meaning with the target domain, or their positions completely overlap or form a logical combination."
        "\n\n1 point: The object has a specific indirect association or supportive role with the target domain within the context of the image and text."
        "\n\n0 points: The object has no relevant relation to the target domain in the image and text, or even conflicts with or causes confusion regarding the target domain."
        "\n\nUse the following formula to calculate the final score: (number of repetitions + score) / 6 to get the final score. "
    )
    messages = [
        {
            "role": "user",
            "content": [
                {"image": image_path},
                {"text": base_prompt + second_Candidate_pool + message + "Ranked in descending order based on scores."}
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


# 根据得分选择候选源域并进行属性选择
def Attribute_Selection(Score_Candidate_pool):
    #  message_content = "Select up to three candidates ranked by score from highest to lowest (if fewer than three, select as many as available) to serve as source domains. Combine these source domains with an external dictionary to select attributes/features for the target domain: " + str(Target[i])   +   " (Requirement 1: Provide attributes that are characteristics of the target domain; Requirement 2: Attributes must be related to the selected source domains). "

    message_content = (
            "This is the pool after scoring. "
            # "If an entity's score is greater than 1, only select objects from the pool with scores greater than 1. "
            #  "If not:"
            "Select the top three candidate source domains from the scored pool (if there is only one or two objects in the pool, select all of them as candidate source domains). "
            "Then, for each selected candidate source domain, choose up to three most relevant attributes or features from an external dictionary. The selected attributes should meet the following requirements: they should be characteristics of the candidate source domain and be relevant to the following text: " + str(Text[i] )+
            ".The best case is to be consistent with the attribute words used in the text to describe the characteristics of the target domain..")

    messages = [
        {
            "role": "user",
            "content": [

                {
                    "text": Score_Candidate_pool + message_content + "For each candidate source domain, select the top three attributes/features that you consider most relevant."}
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


# 从答案里选择喻底


# 形成三元组并对三元组进行余弦相似度排序
def Triples_Generation(Attribute_Selection_result):
    message_content = Attribute_Selection_result + "These are the selected candidate source domains and their corresponding attributes."

    messages = [
        {
            "role": "user",
            "content": [
                {
                    "text": message_content + "According to these form triples in the format [(Candidate Source Domain), (Attributes), (Text: " + str(
                        Text[i]) + ")]." +
                            "Each candidate source field generates three triples(Note that each triple contains only one attribute). Sort all triples by cosine similarity (you don’t need to calculate the exact value, just give the sorting result)." +
                            "Provide the triples in descending order of similarity. The text should remain unchanged."
                }
            ]
        }
    ]

    # 添加调试输出
    print('Messages:', messages)

    try:
        response = dashscope.MultiModalConversation.call(model='qwen-vl-max', messages=messages)

        # 添加调试输出
        print('Response:', response)

    except Exception as e:
        print(f"Error calling model: {e}")
        response = None

    print('进入回答')

    try:
        answer = response["output"]["choices"][0]["message"]["content"][0]["text"]
        # print(f'对话后的答案是{answer}')
        # print(f'answer:{answer}')
    except Exception as e:
        print(f"Error parsing response: {e}")
        answer = "NULL"

    return answer

# 对三元组进行排序
def Source_Computer(Triples):
    message = Triples
    # message_content = "These are all the triples generated by all candidate source domains. Follow these steps: 1. Calculate the cosine similarity scores obtained by pairwise combinations of elements within each of the three quadruples for each candidate source domain(Each quadruple has six scores.Each quadruple score is in the form: [score1, score2, score3, score4, score5, score6].). 2. Calculate the average of the 18 scores obtained from the three quadruples of each candidate source domain. (18 = 3 * 6, where 3 represents each candidate source domain having three quadruples, and 6 represents the six pairwise combinations within each quadruple). 3. Rank each candidate source domain based on the scores given."
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
    answer = response_data["output"]["choices"][0]["message"]["content"][0]["text"]
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
    print(f'response_data:{response_data}')
    # 提取 content 字段
    answer = response_data["choices"][0]["message"]["content"]
    print(f'source_answer:{answer}')
    return answer


# 根据相似度排序进行源域选择
def Source_Selection_Two(Source_Computer):
    message = Source_Computer
    # 根据之前的回答，选择排名最高的源域及其相关的两个最可能的喻底。然后，仅输出与该源域及其属性相关的信息，排除任何其他细节。
    message_content = "According to the previous answers, select the highest-ranking source domain along with the two most probable grounds associated with it from the ranking. Then, output only the information pertaining to this source domain and its attributes, excluding any other details. "
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "text": message + message_content + "Return in dictionary format like this: {\"Source\": \"source\",\"Ground1\":\"attribute1\",\"Ground2\":\"attribute2\"}。Only give the dictionary content, nothing else"}
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
    answer = response_data["output"]["choices"][0]["message"]["content"][0]["text"]
    print(f'Paraphrase: {answer}')
    return answer


# 三元组加上目标域形成四元组生成解释
def Paraphrase_Generation_Two(Source_Selection_result):
    message_content = "Create a quadruplet using the chosen source domain and attribute, coupled with the target domain and text: [Target: { " + \
                      Target[i] + "}, Source Domain: {Source Domain}, Ground: ({Ground1,Ground2}), Text: {" + Text[
                          i] + "}"
    message_generation = "This advertisement contains a metaphor. Based on the provided quadruple, interpret this metaphor and provide a paraphrase. The final response you furnish should be concise, not exceeding 100 words"
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "text": Source_Selection_result + message_content + message_generation}
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


#  根据答案精准提取源域
def Source_Extraction(Source_Selection_result):
    message = Source_Selection_result
    message_content = message + "Select the source domain from the answer, only give the source domain words"
    messages = [
        {
            "role": "user",
            "content": [
                {
                    "text": message_content}
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




# Qwen API Key
api_key = "Your-API-KEY"

# file_path = "D:/PycharmProject/SourceGeneration/data/ads_metaphor.xlsx"
file_path = "D:/PycharmProject/SourceGeneration/data/MET-Meme_New/MET-Meme_New.xlsx"
images_id, Target, Source, Text = data_extraction(file_path)
# prompt设计

prompt_1 = "What entities are present in the image and the text respectively? Please answer regarding objects."
prompt_2_1 = "In the image and text, what elements or features appear somewhat unusual or extraordinary"
prompt_2_2 = "Based on answer 2, extract all the physical entities that appear and record the number of occurrences. Do not ignore any object that appears (even if it is a \"similar entity\")."
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
# print(len(pic_id))
count = 0
# 构建图片地址

# 文件路径
file_path = 'D:/PycharmProject/SourceGeneration/Output/Qwen/Qwen_main_result_Meme.xlsx'

# 检查文件是否存在，如果不存在，创建并写入列名
if not os.path.exists(file_path):
    data = {'image_id': [], 'Source_generation': [], 'Source_AHR': [],  'answers_all': [],
            'first_Candidate_pools': [], 'second_Candidate_pools': [], 'score_results': [],
            'Attribute_Selection_results': [],
            'Triples_all': [], 'Source_Computer_results': []
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
    # if i == 0 or i ==1 or i ==2 or i ==3 or i ==4 or i ==5 :
    #   continue
    # if count == 2:
    #   break
    # 检查ID是否已经在Excel文件中
    if i in df['image_id'].values:
        print(f"ID {i} 已经处理过，跳过...")
        continue
    image_path = "D:/PycharmProject/SourceGeneration/data/MET-Meme_New/" + str(i) + ".jpg"
    print(f'image_path: {image_path}')
    base_prompt = "The text in the image is " + str(Text[i]) + ", and the target domain is " + str(Target[i])
    print(f'base_prompt: {base_prompt}')
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    # 模型回答的所有答案，告诉他上下文
    answer_all = answer_get(base_prompt, image_path, prompt_1, prompt_2_1, prompt_2_2, prompt_3)
    print(f'answer_all: {answer_all}')

    # 根据答案提取所有实体
    first_Candidate_pool = Entity_Extraction(answer_all)
    print(f'first_Candidate_pool: {first_Candidate_pool}')

    # 利用规则筛选
    second_Candidate_pool = Screening_pool(first_Candidate_pool, image_path)
    print(f'second_Candidate_pool: {second_Candidate_pool}')

    # 对筛选后的池子进行打分排序
    score_result = score_candidate_pool_function(second_Candidate_pool, image_path, base_prompt)
    print(f'score_result: {score_result}')

    # 按照得分从高到低选择 然后进行属性选择
    Attribute_Selection_result = Attribute_Selection(score_result)
    print(f'Attribute_Selection_result: {Attribute_Selection_result}')

    # 三元组生成并计算三元组相似度排序
    Triples = Triples_Generation(Attribute_Selection_result)
    print(f'Triples: {Triples}')

     #  根据答案提取源域，以及喻底
    Source_result = Source_Selection(Triples)
    Source_result_dic = json.loads(Source_result)

    try:
        Source_result_dic = json.loads(Source_result)
        source = Source_result_dic['Source']
        ground = Source_result_dic['Ground']
    except:
        source = Source_result

    # 生成解释
    Paraphrase = Paraphrase_Generation(Source_result)
    # 当你收集完所有需要的数据后，创建一个DataFrame
    # 创建一个DataFrame来存储当前迭代的数据
    data = {'image_id': [i], 'Source_generation': [source], 'Source_AHR' : [''],
            'answers_all': [answer_all],
            'first_Candidate_pools': [first_Candidate_pool], 'second_Candidate_pools': [second_Candidate_pool],
            'score_results': [score_result],
            'Attribute_Selection_results': [Attribute_Selection_result],
            'Triples_all': [Triples], 'Source_Computer_results': ['xx']
            }
    df = pd.DataFrame(data)

    # 追加数据到Excel文件中
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
        df.to_excel(writer, startrow=writer.sheets['Sheet1'].max_row, index=False, header=False)
    count = count + 1

