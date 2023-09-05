import sys
sys.path.append(r'c:\users\user\appdata\local\programs\python\python311\lib\site-packages')
import os
import openai
import pandas as pd
from docx import Document

def read_word_file(filepath):
    doc=Document(filepath)#找到文件
    filename=os.path.basename(filepath)
    print("處理 \"",filename,"\" 文檔")
    full_text=[]
    for paragraph in doc.paragraphs:#doc.garagraph 將word拆成段落再加入到full_text
        full_text.append(paragraph.text)
    return '\n'.join(full_text)     #組合成一個字串回傳(整個word的字)

def Openai_ask(text,filepath,start=0,step=2600,overlap=1300):


    #取得API Key
    with open(r"C:\Users\user\Desktop\Py問題生成\OpenaiAPIKey", 'r') as file:
        api_key = file.read()
    openai.api_key = api_key.strip()

    filename=os.path.basename(filepath)+".txt"
    file_path = os.path.join(r"C:\Users\user\Desktop\Py問題生成\Q", filename)
    with open(file_path, 'a'):
        pass
    if len(text)>2600:
        question_num=round(80/round(len(text)/2600))
    else:
        question_num=80
    
    segments=[]
    if(len(text)<overlap):
        overlap=0
    while start<len(text)-overlap:#還有文本則繼續
        end=start+step
        start=end-overlap#讓下一個迴圈從要overlap的地方開始

        messages=[
            {"role":"system","content":f"根據內容生成{question_num}個相關的問題，避免出現你、你的、您、您的、節目等字眼或是私人的問題或具有偏向性敏感性問題，問題不要重複，繁體中文輸出"},
            {"role": "user","content": text[start:end]},
        ]
        completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=messages
            )
        response = completion.choices[0].message.content
        # print(response)
        with open(file_path, 'a',encoding='utf-8') as f:
            f.write(response)
    overlap=1300


def main():
    folder_path = r"C:\Users\user\Desktop\Py問題生成\word"
    allfile = os.listdir(folder_path)
    for i in allfile:
        p = os.path.join(folder_path, i)
        a=read_word_file(p)
        Openai_ask(a,p)
        print(f"{i} 處理完成")
    

if __name__ == '__main__':
    main()