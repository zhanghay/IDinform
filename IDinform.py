# -*- coding: utf-8 -*-
import cv2
import json
import base64
import openpyxl

try :
    from urllib.error import HTTPError
    from urllib.request import Request,urlopen
except ImportError:
    from urllib2 import Request, urlopen, HTTPError
# 阿里云 身份证识别接口
REQUEST_URL = "http://dm-51.data.aliyun.com/rest/160601/ocr/ocr_idcard.json"  # 请求接口


# 将本地图片转成base64编码的字符串，或者直接返回远程图片
def get_img(img_file):
    if img_file.startswith("http"):
        return img_file
    else:
        with open(img_file, 'rb') as f:  # 以二进制读取本地图片
            data = f.read()
    try:
        encodestr = str(base64.b64encode(data),'utf-8')
    except TypeError:
        encodestr = base64.b64encode(data)

    return encodestr

# 发送请求，获取识别信息
def posturl(headers, body):
    try:
        params=json.dumps(body).encode(encoding='UTF8')
        req = Request(REQUEST_URL, params, headers)
        r = urlopen(req)
        html = r.read()
        # str 转成 json以获取key value
        return json.loads(html)
        #return html.decode("utf8")
    except HTTPError as e:
        print(e.code)
        print(e.read().decode("utf8"))

def parse(appcode, img_file, side='face'):
    config = {'side': side}  # face:表示正面；back表示反面

    # 请求体
    body = {"configure": config}
    img_info = get_img(img_file)
    body.update({'image':img_info})

    # 请求头
    headers = {
        'Authorization': 'APPCODE %s' % appcode,
        'Content-Type': 'application/json; charset=UTF-8'
    }

    html = posturl(headers, body)
    #print(html.keys())
    return(html)

def get_useful_inform(img_file):
    appcode = ''# 你的Appcode 阿里云购买
    hj=parse(appcode, img_file)
    ## 是否成功，姓名，身份证号，性别，生日，住址，通过key
    success,name,num,sex,birth,address=hj['success'],hj['name'],hj['num'],hj['sex'],hj['birth'],hj['address']
    if success==True:
        return name,num,sex,birth,address
    else :
        print('unsuccessful')

def writeFiles(image_file,sheet):
    name,num,sex,birth,address=get_useful_inform(image_file)
    print(num,name)
    #获取生日年月日
    year=eval(birth[:4])
    moth=eval(birth[4])*10+eval(birth[5])
    day=eval(birth[6])*10+eval(birth[7])
    #年龄
    age=2021-year
    if moth<1:
        age-=1
    if moth==1 and day<28:
        age-=1
    data=[name,address,sex,num,str(age)]
    sheet.append(data)
    workbook.save('ex.xlsx')
if __name__=="__main__":
    #摄像头 0号
    cap = cv2.VideoCapture(0,cv2.CAP_DSHOW) 
    num=0
    #要写入的表格
    workbook=openpyxl.load_workbook('ex.xlsx')
    #要写入的工作表
    sheet=workbook['Sheet1']
    while True:
        ret,frame = cap.read()
        cv2.imshow('frame',frame)
        c = cv2.waitKey(1)
        #按w拍摄，请求识别，写入表格
        if c==ord('w'):
            num+=1
            filename="frame_%s.jpg"%num
            cv2.imwrite(filename,frame)
            num = num+1
            writeFiles(filename,sheet)
        elif c == ord('q'):
            break
    cap.release()
    cv2.destroyAllWindows()
