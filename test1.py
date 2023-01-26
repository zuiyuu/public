
    #原始数据读取
    def loadData(self):
        fp = open(self.dataFile, 'rb')

        self.date_l = []            #存放日期
        self.ki_l = []              #存放期数
        self.alldate = []           #存放全部中奖号码
        self.num_l = []             #存放中奖号码
        #从文件种读取所需的数据
        while True:
            line = fp.readline().decode()
            if not line:
                break
            data = line.split(";")
            if len(data) < 2:
                continue
            num = re.findall(r'\d+',data[2])        #从(2019-07-21;2019084;04,08,14,18,20,27,03)提取出[04,08,14,18,20,27,03]
            # self.date_l.append(int(''.join(re.findall("\d",data[0]))))
            num_list = [ int(x) for x in re.findall("\d",data[0]) ]
            self.date_l.append(np.mean(num_list))
            self.ki_l.append(data[1])
            self.alldate.append(num)
            num = num[self.n]                       #从[04,08,14,18,20,27,03]中选择第n个
            self.num_l.append(int(num))
            
    # 构造满足LSTM的训练数据
    def buildTrainDataSet(self):
        
        if (commonPara == 1):
            self.num_l.reverse()
            self.meanNum = np.mean(self.num_l)      #平均值
            self.stdNum = np.std(self.num_l)        #标准差
            self.Data = (self.num_l - self.meanNum) / self.stdNum  # 标准化
        
        if (commonPara == 2):
            self.date_l.reverse()
            self.date_l_meanNum = np.mean(self.date_l)      #期数平均值
            self.date_l_stdNum = np.std(self.date_l)        #期数标准差
            
            self.num_l.reverse()
            self.meanNum = np.mean(self.num_l)      #平均值
            self.stdNum = np.std(self.num_l)        #标准差
            self.Data = abs(self.num_l - (self.meanNum + self.date_l_meanNum)) / (self.stdNum + self.date_l_stdNum)  # 标准化

        self.Data = self.Data[:, np.newaxis]  # 增加维度
        for i in range(len(self.Data)-self.timeStep-1):
            x = self.Data[i:i+self.timeStep]
            y = self.Data[i+1:i+self.timeStep+1]

            self.train_x.append(x)
            self.train_y.append(y)
            
            
 
# 参数1：构造训练标准参数 参数2：学习率
def setTrainParam(flag=1,lrflag=0):
    retFlag = 1
    retLrFlag = 0.0006
    # flag == 1 标准参数设定 buildTrainDataSet方法内体现
    if (flag == 1):
        retFlag = 1
    # flag == 2 标准参数追加期数数组平均值 buildTrainDataSet方法内体现
    elif (flag == 2) :
        retFlag = 2
    
    if (lrflag == 0):
        retLrFlag = 0.0006 #学习率标准设定 0.0006
    elif  (lrflag == 1) :
        retLrFlag = 0.0003 #学习率低标准设定 0.0003
    elif  (lrflag == 2) :
        retLrFlag = 0.0004 #学习率低标准设定 0.0004
    elif  (lrflag == 3) :
        retLrFlag = 0.0005 #学习率高标准设定 0.0005
    elif  (lrflag == 4) :
        retLrFlag = 0.0007 #学习率高标准设定 0.0007
    elif  (lrflag == 5) :
        retLrFlag = 0.0008 #学习率高标准设定 0.0008
    elif  (lrflag == 6) :
        retLrFlag = 0.0009 #学习率高标准设定 0.0009
    
    return retFlag,retLrFlag
if __name__ == '__main__':
    currentDirectory  =  os.getcwd()
    fileName = 'DCdemo-master\DataPreparation\DCnumber.txt'
    path  =  "%s\\%s"  %( currentDirectory, fileName)
    timeDate = time.strftime('%Y%m%d%H%M', time.localtime())
    testRootPath = "%s\\%s"  %( currentDirectory, 'DCdemo-master\DCModel_')
    getNumber = []
    
    #0表示训练，1表示预测
    typeinfo = 1
    for x in range(7):
        print('第'+str(x)+'次训练和预测。')
        # 参数1：构造训练标准参数 参数2：标准学习率
        commonPara,commonLr = setTrainParam(1,x)
        # 通过7次的训练和预测查看7次的结果
        for y in range(2):
            type = y
            if (typeinfo == 1 and 0 == y) :
                continue
            # type = 1             #0表示训练，1表示预测
            preNumber = []      #存放预测出来的值
            for n in range(7):    #0-5表示从1到6个红球，6表示篮球
                if type == 0:
                    predictor = DCPredictor(path,n)
                    predictor.loadData()
                    # 构建训练数据
                    predictor.buildTrainDataSet()
                    testPath = "DCdemo-master\\%s\\DCModel_"  %(timeDate+str(x))
                    testRootPath = "%s\\%s"  %( currentDirectory, testPath)
                    # 模型训练 test
                    predictor.trainLstm()
                else:
                    timeDate = '202301251614'
                    trainRootPath = "%s\\%s\\%s"  %( currentDirectory, 'DCdemo-master', timeDate+str(x))
                    fileList = getFiledList(trainRootPath)
                    onePreNumber = []
                    for fileRootPath in fileList:
                        if ('DCModel_' in fileRootPath) :
                            predictor = DCPredictor(path,n)
                            predictor.loadData()
                            # 构建训练数据
                            predictor.buildTrainDataSet()
                            testRootPath = fileRootPath
                            # 预测－预测前需要先完成模型训练
                            number = predictor.prediction()
                            onePreNumber.append(number)
                    preNumber.append({n:onePreNumber})
            # print(preNumber)
            # print(preNumber)
            if type == 1:
                getNumber.append(preNumber)
    print('全部训练模型的预测结果：')
    print(getNumber)
    getNum0 = 0
    getNum1 = 0
    getNum2 = 0
    getNum3 = 0
    getNum4 = 0
    getNum5 = 0
    getNum6 = 0
    getNumAll = []
    for lastNum in getNumber:
        getNumOne = []
        getNum0 = [ int(x) for x in lastNum[0][0] ]
        getNum1 = [ int(x) for x in lastNum[1][1] ]
        getNum2 = [ int(x) for x in lastNum[2][2] ]
        getNum3 = [ int(x) for x in lastNum[3][3] ]
        getNum4 = [ int(x) for x in lastNum[4][4] ]
        getNum5 = [ int(x) for x in lastNum[5][5] ]
        getNum6 = [ int(x) for x in lastNum[6][6] ]
        
        # 平均值
        getNumOne.append(int(np.mean(getNum0)))      #平均值
        getNumOne.append(int(np.mean(getNum1)))      #平均值
        getNumOne.append(int(np.mean(getNum2)))      #平均值
        getNumOne.append(int(np.mean(getNum3)))      #平均值
        getNumOne.append(int(np.mean(getNum4)))      #平均值
        getNumOne.append(int(np.mean(getNum5)))      #平均值
        getNumOne.append(int(np.mean(getNum6)))      #平均值
        getNumAll.append(getNumOne)
    print(getNumAll)
