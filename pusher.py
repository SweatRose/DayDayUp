import os
import glob
import requests
from datetime import datetime
from datetime import datetime
from docx import Document
import random

# # 获取最新的 Word 文档
# def get_latest_word_document():
#     list_of_files = glob.glob('.archive/*.docx')  # 获取当前目录下所有的 .docx 文件
#     print(list_of_files)
#     latest_file = max(list_of_files, key=os.path.getctime)  # 获取最新的文件
#     return latest_file
# 获取最新的 Word 文档
# def get_latest_word_document(directory):
#     list_of_files = glob.glob(os.path.join(directory, '*.docx'))  # 获取指定目录下所有的 .docx 文件
#     if list_of_files:
#         latest_file = max(list_of_files, key=os.path.getctime)  # 获取最新的文件
#         return latest_file
#     else:
#         return None
def get_latest_word_document():
    # 假设你想在当前目录及其子目录中随机的 一篇Word 文档
    files = glob.glob('**/*.docx', recursive=True)
    if files:
        return random.choice(files)
    else:
        return None


def read_docx_content(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def push_to_enterprise_wechat(docx_file_path, agent_id, corp_id, corp_secret, media_id):
    # 上传 Word 文档到企业微信
    # 这里需要使用企业微信提供的接口进行文章的上传和推送
    # 需要使用到企业微信的开发者权限，包括 agent_id、corp_id 和 corp_secret
    # 以下代码仅作为示例，实际使用时需要将以下代码与企业微信的接口进行对接

    # 上传临时素材并获取media_id
    def upload_temp_media(file_path):
        upload_url = "https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token={}&type=image".format(get_access_token(corp_id, corp_secret))
        files = {'media': open(file_path, 'rb')}
        response = requests.post(upload_url, files=files)
        media_id = response.json().get('media_id')
        return media_id

    #获取content内容 来自最新的word内容
    content = read_docx_content(docx_file_path)
    dt = datetime.now()
    date = dt.strftime('%Y-%m-%d')
    time = dt.strftime('%H:%M:%S')
    file_name = os.path.basename(docx_file_path).split('.')[0]
    push_url = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={}".format(get_access_token(corp_id, corp_secret))
    #修复封面图片三天过去，即临时素材
    thumb_media_id = upload_temp_media('temp.jpg')  # 上传temp.jpg并获取media_id
    data = {
        'touser': '@all',
        'agentid': agent_id,
        'msgtype': 'mpnews',
        'mpnews': {
            'articles': [{
                # 'title': f'每日一练 - {file_name}- {date}',
                'title': f'每日一练 - {file_name}',
                #写不动了  直接写死一张图片(三天一更新 暂时没想到好办法)
                # 'thumb_media_id': '3Dmx7S_K6YtJ8qIGKE9MrxxyR5k8PsabPR7ZIf4j1FP4_5A0-TLabKkZ5sGmvz38L',  # 使用上传文件后得到的 media_id 作为缩略图
                'thumb_media_id': thumb_media_id,  # 使用上传文件后得到的 media_id 作为缩略图
                'author': 'Tonyz',
                'digest': f'{date}-{time} - 每日一练推送',
                'content': content
                }
            ]
        }
    }
    response = requests.post(push_url, json=data)
    if response.json()['errcode'] == 0:
        print("文章推送成功")
    else:
        print("文章推送失败")
        print(response.json())



# 获取 access_token
def get_access_token(corp_id, corp_secret):
    access_token_url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={}&corpsecret={}".format(corp_id, corp_secret)
    response = requests.get(access_token_url)
    return response.json()['access_token']

if __name__ == "__main__":
    agent_id = "1000002"  # 替换为企业微信应用的 agent_id
    corp_id = "ww60ea66699fe5c551"  # 替换为企业微信的 corp_id
    corp_secret = "AGAEKr4HsQFbBsb4YAyp8BaoLGVocRQ3f8x-h48pWGY"  # 替换为企业微信的 corp_secret
    # directory_path = "./archive"  # 替换为实际的目录路径
    latest_docx_file = get_latest_word_document()
    if latest_docx_file:
        upload_url = "https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token={}&type=file".format(get_access_token(corp_id, corp_secret))
        files = {'file': open(latest_docx_file, 'rb')}
        response = requests.post(upload_url, files=files)
        result = response.json()
        if result['errcode'] == 0:
            media_id = result['media_id']
            push_to_enterprise_wechat(latest_docx_file, agent_id, corp_id, corp_secret, media_id)
        else:
            print("文件上传失败")
    else:
        print("未找到任何 .docx 文件")
    # push_to_enterprise_wechat(latest_docx_file, agent_id, corp_id, corp_secret)