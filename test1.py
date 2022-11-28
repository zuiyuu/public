import os
import time
import yagmail
from imbox import Imbox
from PIL import ImageGrab

class mailTools:
    username = ''
    password = ''
    receiver = ''
    imapAdd = 'imap.163.com'
    smtpAdd = 'smtp.163.com'
    
    def __init__(self,un,pw) -> None:
        if (un == '' or pw == ''): 
            username = 'test@163.com'
            password = '你的授权码'
            receiver = 'test@qq.com'
        else:
            username = un
            password = pw
            receiver = un
        # yagmail.register(username, password)

    def send_mail(self, sender, to, contents,subjectText):
        if (sender == ''):
            sender = self.username
        if (to == '') :
            to = self.receiver
        # # smtp = yagmail.SMTP(user=sender, host='smtp.163.com')
        # smtp = yagmail.SMTP(user=sender, host=self.smtpAdd)
        # # smtp.send(to, subject='Remote Control', contents=contents)
        # smtp.send(to, subject=subjectText, contents=contents)
        
    def read_mail(self, username, password):
        # with Imbox('imap.163.com', username, password, ssl=True) as box:
        with Imbox(self.imapAdd, username, password, ssl=True) as box:
            all_msg = box.messages(unread=True)
            for uid, message in all_msg:
                # 如果是手机端发来的远程控制邮件
                if message.subject == 'Remote Control':
                    # 标记为已读
                    box.mark_seen(uid)
                    return message.body['plain'][0]
                
    def shutdown():
        os.system('shutdown -s -t 0')
        
 import win32api     #需要事先安装该模块
import pyautogui    #只用该模块的截图功能，用pillow也可以哦
import time         
import datetime
import os
import imgHashCheck
import colGetMail
from configparser import ConfigParser

def creatDBInifile(filePath):
    # host     = 127.0.0.1
    # user     = root
    # password = 123456
    # port     = 3306
    # database = mysql
    if (os.path.exists(filePath)) :
        return 
    else :
        with open(filePath,"w") as f:
            f.write("[monitor]\n")
            # Mail用户
            f.write("user             = test@163.com\n") 
            # Mail密码
            f.write("password         = 12345zaqwsxedcrfvgtbyhnjmkmiiiiiiii\n") 
            # 邻近均值哈希算法相似度
            f.write("fSimilarity      = 0.8\n")
            # 邻近三直方图算法相似度
            f.write("fThreeSimilarity = 0.56\n") 
            # 参照物均值哈希算法相似度
            f.write("similarity      = 0.5\n")
            # 参照物图算法相似度
            f.write("threeSimilarity = 0.4\n") 
            # 参照物File名
            f.write("rootfileName = 1.png,2.png,3.png\n") 
            
def readDbInfo (filePatn) :
    cfg = ConfigParser()
    cfg.read(filePatn)
    return cfg.items("monitor")
    # def mainFaction(flag, sql) :
    
def getInitInfo():
    #得到当前脚本的执行目录
    currentDirectory  =  os.getcwd()
    #查看是否已经存在备份目录，如果有则删除，没有则新建目录
    backUpDirectory  =  "%s\\%s"  %( currentDirectory, "similarity.ini")
    creatDBInifile(backUpDirectory)
    dbinfo = readDbInfo(backUpDirectory)
    return dbinfo

 
dbinfo =  getInitInfo()
# Mail用户
username = dbinfo[0][1]
# Mail密码
password = dbinfo[1][1]
# 邻近均值哈希算法相似度
fSimilarity = float(dbinfo[2][1])
# 邻近三直方图算法相似度
fThreeSimilarity = float(dbinfo[3][1])
# 参照物均值哈希算法相似度
similarity = float(dbinfo[4][1])
# 参照物图算法相似度
threeSimilarity = float(dbinfo[5][1])
# 参照物File名
rootfileName = str(dbinfo[6][1])

#得到当前脚本的执行目录
currentDirectory  =  os.getcwd()
filedPath = "%s\\%s"  %( currentDirectory, "temp\\img")
if not os.path.exists(filedPath):#首先需要建立一个文件夹用于保存截图
    os.makedirs(filedPath)
    # os.mkdir(filedPath)
rootFiledPath = "%s\\%s"  %( currentDirectory, "temp")
if not os.path.exists(rootFiledPath):#首先需要建立一个文件夹用于保存截图
    os.mkdir(rootFiledPath)
    
oldFilePath = ''
referenceFilePath = rootFiledPath + '\\'
referenceFiledPathList = []
for rootFile in rootfileName.split(','):
    referenceFiledPathList.append(referenceFilePath + rootFile)
    
colGetMail = colGetMail.mailTools(username,password)


# 根据允许表示页面List检查当前画面相似度，只要有一个画面满足相似度方位返回对应值
def checkImgs(oldFilePath,tempFilePath,flag):
    if (flag == '1') :
        checkRet1,checkRet2 = imgHashCheck.imgCheck(oldFilePath,tempFilePath)
        if (checkRet1 >= fSimilarity and checkRet2 >= fThreeSimilarity):
            return True
    elif (flag == '2'):
        for referencePath in referenceFiledPathList:
            checkRet1,checkRet2 = imgHashCheck.imgCheck(referencePath,tempFilePath)
            if (checkRet1 >= similarity and checkRet2 >= threeSimilarity):
                return True
    return False
       
start=time.time()
while True: #无线循环
    num = win32api.GetAsyncKeyState(0x01)#监测键盘某一按键状态，你想具体监测那个键可以差asii码表，我这个是鼠标左键
    num1 = win32api.GetAsyncKeyState(0x0d)#监测键盘某一按键状态，你想具体监测那个键可以差asii码表，我这个是回车
    if num or num1:  #如果为真则截图，也可以加上空格键 用or连接就可以
        strDatatime = datetime.datetime.now().strftime('%Y%m%d%H_%M%S')
        tempFilePath = filedPath + '\\{}.png'.format(strDatatime)
        pyautogui.screenshot(tempFilePath)#文件名用日期加时间格式
        
        if (oldFilePath == '') :
            oldFilePath = tempFilePath
            continue
        else:
            # checkRetfirst1,checkRetfirst2 = imgHashCheck.imgCheck(oldFilePath,tempFilePath)
            checkFlag = checkImgs(oldFilePath,tempFilePath,'1')
            oldFilePath = tempFilePath
            if (checkFlag == False):
                print('时间：' + strDatatime + '。First画面变更了，可以发出警告了。')
                print('变更画面时点图片：' + tempFilePath)
                colGetMail.send_mail('','',tempFilePath,'操作变化请确认是否有问题')
                continue
                
        checkFlag = checkImgs('',tempFilePath,'2')
        if (checkFlag == False):
            print('时间：' + strDatatime + '。画面变更了，可以发出警告了。')
            print('变更画面时点图片：' + tempFilePath)
            colGetMail.send_mail('','',tempFilePath,'操作变化请确认是否有问题')
            
        end=time.time()
        usedatetime = divmod(round((end-start),2),60)
        if (int(usedatetime[0]) > 30):
            print("保存图片文件清除！！！")
            os.remove(filedPath)
            
    time.sleep(0.5)#这个必须有否则cpu占用率会很高，休眠时间也可以自己更改

    # -*- coding: utf-8 -*-
import cv2
import numpy as np

# 均值哈希算法
def aHash(img,shape=(10,10)):
    # 缩放为10*10
    img = cv2.resize(img, shape)
    # 转换为灰度图
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # s为像素和初值为0，hash_str为hash值初值为''
    s = 0
    hash_str = ''
    # 遍历累加求像素和
    for i in range(shape[0]):
        for j in range(shape[1]):
            s = s + gray[i, j]
    # 求平均灰度
    avg = s / 100
    # 灰度大于平均值为1相反为0生成图片的hash值
    for i in range(shape[0]):
        for j in range(shape[1]):
            if gray[i, j] > avg:
                hash_str = hash_str + '1'
            else:
                hash_str = hash_str + '0'
    return hash_str

# 差值感知算法
def dHash(img,shape=(10,10)):
    # 缩放10*11
    img = cv2.resize(img, (shape[0]+1, shape[1]))
    # 转换灰度图
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    hash_str = ''
    # 每行前一个像素大于后一个像素为1，相反为0，生成哈希
    for i in range(shape[0]):
        for j in range(shape[1]):
            if gray[i, j] > gray[i, j + 1]:
                hash_str = hash_str + '1'
            else:
                hash_str = hash_str + '0'
    return hash_str


# 感知哈希算法(pHash)
def pHash(img,shape=(10,10)):
    # 缩放32*32
    img = cv2.resize(img, (32, 32))  # , interpolation=cv2.INTER_CUBIC

    # 转换为灰度图
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 将灰度图转为浮点型，再进行dct变换
    dct = cv2.dct(np.float32(gray))
    # opencv实现的掩码操作
    dct_roi = dct[0:10, 0:10]

    hash = []
    avreage = np.mean(dct_roi)
    for i in range(dct_roi.shape[0]):
        for j in range(dct_roi.shape[1]):
            if dct_roi[i, j] > avreage:
                hash.append(1)
            else:
                hash.append(0)
    return hash


# 通过得到RGB每个通道的直方图来计算相似度
def classify_hist_with_split(image1, image2, size=(256, 256)):
    # 将图像resize后，分离为RGB三个通道，再计算每个通道的相似值
    image1 = cv2.resize(image1, size)
    image2 = cv2.resize(image2, size)
    sub_image1 = cv2.split(image1)
    sub_image2 = cv2.split(image2)
    sub_data = 0
    for im1, im2 in zip(sub_image1, sub_image2):
        sub_data += calculate(im1, im2)
    sub_data = sub_data / 3
    return sub_data


# 计算单通道的直方图的相似值
def calculate(image1, image2):
    hist1 = cv2.calcHist([image1], [0], None, [256], [0.0, 255.0])
    hist2 = cv2.calcHist([image2], [0], None, [256], [0.0, 255.0])
    # 计算直方图的重合度
    degree = 0
    for i in range(len(hist1)):
        if hist1[i] != hist2[i]:
            degree = degree + (1 - abs(hist1[i] - hist2[i]) / max(hist1[i], hist2[i]))
        else:
            degree = degree + 1
    degree = degree / len(hist1)
    return degree



# Hash值对比
def cmpHash(hash1, hash2,shape=(10,10)):
    n = 0
    # hash长度不同则返回-1代表传参出错
    if len(hash1)!=len(hash2):
        return -1
    # 遍历判断
    for i in range(len(hash1)):
        # 相等则n计数+1，n最终为相似度
        if hash1[i] == hash2[i]:
            n = n + 1
    return n/(shape[0]*shape[1])

# 均值哈希算法相似度/三直方图算法相似度检查
def imgCheck(imgFile1, imgFile2):
    img1 = cv2.imread(imgFile1)  
    # img2 = cv2.imread('D:\\work\\python\\temp\\2.png')
    img2 = cv2.imread(imgFile2)
    
    hash1 = aHash(img1)
    hash2 = aHash(img2)
    n1 = cmpHash(hash1, hash2)
    print('均值哈希算法相似度：', n1)
    
    hash1 = dHash(img1)
    hash2 = dHash(img2)
    n3 = cmpHash(hash1, hash2)
    print('差值哈希算法相似度：', n3)
    
    n2 = classify_hist_with_split(img1, img2)
    if (isinstance(n2,list)):
        print('三直方图算法相似度：', n2[0])
        return n1, n2[0]
    else:
        print('三直方图算法相似度：', n2)
        return n1, n2
    # return n1, n2[0]

def main():
    img1 = cv2.imread('D:\\work\\python\\temp\\4.png')  
    img2 = cv2.imread('D:\\work\\python\\temp\\7.png')
    # img2 = cv2.imread('D:\\work\\python\\temp\\2022112813_2148.png')

    hash1 = aHash(img1)
    hash2 = aHash(img2)
    n = cmpHash(hash1, hash2)
    print('均值哈希算法相似度：', n)

    hash1 = dHash(img1)
    hash2 = dHash(img2)
    n = cmpHash(hash1, hash2)
    print('差值哈希算法相似度：', n)

    hash1 = pHash(img1)
    hash2 = pHash(img2)
    n = cmpHash(hash1, hash2)
    print('感知哈希算法相似度：', n)

    n = classify_hist_with_split(img1, img2)
    print('三直方图算法相似度：', n[0])

    n = calculate(img1, img2)
    print('单通道的直方图算法相似度：', n[0])

if __name__=="__main__":
    main()

    
    
    
