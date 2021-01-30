

from selenium import webdriver
import time
import openpyxl
import datetime
import random

logFile = open("./log",'w',encoding='utf-8')

def login():
    driver.find_element_by_xpath('//*[@class="tabloginway"]/div[2]/p').click()

    elm=driver.find_element_by_xpath('//*[@id="box"]/div[2]/div/div[2]/div[2]/div[1]/input')
    elm.click()
    elm.send_keys(account)

    elm=driver.find_element_by_xpath('//*[@id="box"]/div[2]/div/div[2]/div[3]/div[1]/div[1]/input')
    elm.click()
    elm.send_keys(password)

    driver.find_element_by_xpath('//*[@id="box"]/div[2]/div/div[2]/div[3]/div[4]/button').click()

    try:
        driver.find_elements_by_class_name("userInfo_name")
        logFile.write("登录成功\n")
        return
    except:
        exit(0)

def read_ids():
    ids ={}
    with open(ids_path,'r',encoding='utf-8') as inputFile:
        for i in inputFile.readlines():
            line =i.strip().split()
            ids[line[1]]=line[0]
    print(ids)
    return ids

def write():
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "时间"+beginData.strftime('%Y%m%d%H%M%S')+" "+endData.strftime('%Y%m%d%H%M%S')
        sheet['A1'] = "id"
        sheet['B1'] = "社区名称"
        sheet['C1'] = "标题"
        sheet['D1'] = "回复量"
        sheet['E1'] = "浏览量"
        sheet['F1'] = "发表时间"
        for i in total:
            link ='=HYPERLINK("'+i[2]+'","'+i[1]+'")'
            sheet.append([i[0], id2name[i[0]], link, i[3], i[4], i[5]])

        workbook.save("绩效数据"+str(random.randint(0,1000))+".xlsx")
    except:
        logFile.write("写入报表失败，查看是否同名文件未关闭\n")





def parse(id='255527',num =8):
    result = []
    try:
        url ="https://www.oneplusbbs.com/home.php?mod=space&uid="+id+"&do=thread&view=me&from=space&type=thread"
        driver.get(url)
        if "id.oneplus.com" in driver.current_url:
            logFile.write("需要登录\n")
            login()
            while "id.oneplus.com" in driver.current_url:
                time.sleep(10)

        logFile.write("\n开始解析\n")

        table = driver.find_element_by_xpath('//*[@id="delform"]/table')
        trlist=table.find_elements_by_tag_name('tr')
        for row in trlist[1:num+1]:
            th = row.find_elements_by_tag_name('th')[0].find_element_by_xpath("a[1]")
            nums = row.find_elements_by_tag_name('td')[2].text.split('/')
            title=th.text
            titleUrl=th.get_attribute("href")
            replyNum = nums[0]
            readNum =nums[1]
            logFile.write(id + " " + title + " " + titleUrl + " " + replyNum + " " + readNum + "\n")
            result.append([id,title,titleUrl,replyNum,readNum])
        return result
    except:
        return result


def isInTime(url):
    driver.get(url)
    try:
        em = driver.find_element_by_xpath('//*[@class="authi"]/em')
        if em.find_element_by_css_selector("span"):
            createTime =em.find_element_by_css_selector("span").get_attribute("title").strip()
        else:
            createTime =em.text.replace("发表于","").strip()
        createData = datetime.datetime.strptime(createTime, '%Y-%m-%d %H:%M:%S')
        if createData >= beginData and createData <= endData:
            return True,createTime
        else:
            return False,None
    except:
        return False, None

def readFile():
    file = open("input.txt",'r',encoding='utf-8')
    account =file.readline().strip()
    password =file.readline().strip()
    beginTime =file.readline().strip()
    endTime =file.readline().strip()
    file.close()
    return account,password,beginTime,endTime


if __name__ == '__main__':
    # 获取driver
    path = './chromedriver'
    ids_path = "ids.txt"
    account, password, beginTime, endTime = readFile()
    beginData = datetime.datetime.strptime(beginTime.strip(), '%Y-%m-%d %H:%M:%S')
    endData = datetime.datetime.strptime(endTime.strip(), '%Y-%m-%d %H:%M:%S')

    driver = webdriver.Chrome(executable_path=path)

    id2name = read_ids()

    total = []
    for i in id2name.keys():
        for j in parse(str(i)):
            isTime, createTime = isInTime(j[2])
            if isTime==False:
                logFile.write("不在时间范围内\n")
                break
            logFile.write(createTime+"\n")
            j.append(createTime)
            total.append(j)
    write()
    logFile.close()
    driver.quit()
