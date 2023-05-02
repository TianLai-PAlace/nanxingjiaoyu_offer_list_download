# -*- codeing = utf-8 -*-
from bs4 import BeautifulSoup  # 网页解析，获取数据
import re  # 正则表达式，进行文字匹配`
import urllib.request, urllib.error  # 制定URL，获取网页数据
import xlwt  # 进行excel操作

findLink = re.compile(r'<a href="(.*?)">') #创建正则表达式对象，表示规则（字符串的模式）
find_text_span = re.compile(r'<span>(.*)</span>')
find_text_b = re.compile(r'<b>(.*?)</b>', re.S)

dict_searchid = {
    "sg": "1941",
    "hk":"1940",
    "uk":"1982",
    "us":"1970",
    "Europe":"1972",
    "中外合办":"2042",
}

#config
now_use_searchid = "uk" #选择要爬取的地区
pages =20 #选择要爬取的页数


def main():
    baseurl = "https://www.nanxingjiaoyu.com/e/search/result/index.php?page="  #要爬取的网页链接
    # 1.爬取网页
    linklist = getlink(baseurl,pages)  ##爬取所有offer链接
    print(linklist)
    datalist = getdata(linklist)

    savepath = "offer结果_"+now_use_searchid+".xls"    #当前目录新建XLS，存储进去
    #dbpath = "movie.db"              #当前目录新建数据库，存储进去
    # 3.保存数据
    saveData(datalist,savepath)      #2种存储方式可以只选择一种
    # saveData2DB(datalist,dbpath)



# 爬取链接
def getlink(baseurl,pages):
    linklist = []  #用来存储爬取的所有链接
    for i in range(0, pages):  # 调用获取页面信息的函数，10页
        url = baseurl + str(i) + "&searchid="+dict_searchid[now_use_searchid]
        html = askURL(url)  # 保存获取到的网页源码
        # 2.逐一解析数据
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all('div', class_="of-area"):  # 查找符合要求的字符串
            item = str(item)
            for link in re.findall(findLink, item):
                if not link.startswith("/case"):
                    continue
                link="https://www.nanxingjiaoyu.com"+link.split('"')[0]
                linklist.append(link)
    return linklist

# 爬取内容
def getdata(linklist):
    datalist = []  #用来存储爬取的所有链接
    for url in linklist: 
        html = askURL(url)  # 保存获取到的网页源码
        # 2.逐一解析数据
        soup = BeautifulSoup(html, "html.parser")
        list = []
        for item in soup.find_all('div', class_="of-xsbj")[0]:  # 查找符合要求的字符串
            item = str(item)
            for item in re.findall(find_text_b, item):
                list.append(item)
        try:
            for item in soup.find_all('div', class_="of-xsbj")[1]:  # 查找符合要求的字符串
                item = str(item)
                for item in re.findall(find_text_span, item):
                    if item.startswith("录取结果"):
                        continue
                    # print(item)
                    item = item.split(":")[1]
                    list.append(item)
        except Exception as e:
            print(e)
            for i in range(4):
                list.append("null")
        #第九列，转换4分制GPA
        if "/" in list[2]:
            x1=[float(s) for s in re.findall(r'-?\d+\.?\d*', list[2])][0]
            x2=[float(s) for s in re.findall(r'-?\d+\.?\d*', list[2])][1]
            x_min = min([x1,x2])
            x_max = max([x1,x2])
            x1 = x_min
            x2 = x_max
            list.append(str(x1))
        else:
            try:
                x1 =  [float(s) for s in re.findall(r'-?\d+\.?\d*', list[2])][0]
            except Exception as e:
                print(e)
                x1 = -1
            if x1>=5:
                x = (x1-60)/10 +1.5
                if x > 4: x=4
                list.append(str(x))
            else:
                list.append(str(x1))

        datalist.append(list)
        print(list)
    return datalist


# 得到指定一个URL的网页内容
def askURL(url):
    head = {  # 模拟浏览器头部信息，向服务器发送消息
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0; Win64; x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 80.0.3987.122  Safari / 537.36"
    }
    # 用户代理，表示告诉服务器，我们是什么类型的机器、浏览器（本质上是告诉浏览器，我们可以接收什么水平的文件内容）

    request = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


# 保存数据到表格
def saveData(datalist,savepath):
    print("save.......")
    book = xlwt.Workbook(encoding="utf-8",style_compression=0) #创建workbook对象
    sheet = book.add_sheet('结果', cell_overwrite_ok=True) #创建工作表
    col = ("学校","专业","GPA","语言","姓名","录取学校","录取专业","入学时间","4分换算gpa")
    for i in range(0,len(col)):
        sheet.write(0,i,col[i])  #列名
    for i in range(0,250):
        # print("第%d条" %(i+1))       #输出语句，用来测试
        try:
            data = datalist[i]
            for j in range(0,len(col)):
                sheet.write(i+1,j,data[j])  #数据
        except:
            continue
    book.save(savepath) #保存


if __name__ == "__main__":  # 当程序执行时
    # 调用函数
     main()

     print("爬取完毕！")


