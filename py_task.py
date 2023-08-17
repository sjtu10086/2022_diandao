import urllib.request
import urllib.error
import xlwt
import re
import tkinter
import matplotlib.pyplot as plt


def main():  # 主函数
    dic = {}
    cou = []
    data = []
    baseurl = "https://news.ifeng.com/c/special/85mhVvWS5i4"
    html = askurl(baseurl)
    dic = getData(html, dic, cou, data)
    output(cou, data)
    bubble_sort(cou, data)
    savedata(cou, data)
    plot(data)
    tk(dic)


def getData(html, dic, cou, data):  # 正则表达式获取数据存入字典与列表
    i = 0
    regex1 = r"\"country\":\"(.*?)\""
    country = re.findall(
        regex1,
        html,
    )
    regex2 = r"\"per_hundred\":(.*?)}"
    per_hundred = re.findall(regex2, html)
    while (i < len(country)):
        dic[country[i]] = per_hundred[i + 4]
        cou.insert(i, country[i])
        data.insert(i, float(per_hundred[i + 4]))
        i = i + 1
    # 源码提取发现"巴勒斯坦"显示为"\n巴勒斯坦"，手动修改
    cou[102] = "巴勒斯坦"
    del dic[r"\n巴勒斯坦"]
    dic["巴勒斯坦"] = data[102]
    return dic


def askurl(url):  # 获取网页源代码存入html
    head = {
        "User-agent":
        "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"
    }
    req = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(req)
        html = response.read().decode("utf-8")
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


def tk(dic):  # 构建tkinter交互窗口
    root = tkinter.Tk()
    root.geometry("600x300")
    root.title("新冠疫苗接种率")
    label = tkinter.Label(root, text="国家", font=("宋体", "25"), fg="red", bg="white")
    label.pack()
    entry = tkinter.Entry(root, font=("宋体", "25"), fg="red")
    entry.pack()

    def search():
        text.delete("1.0", "end")
        var = entry.get()
        if var in dic.keys():
            text.insert("end", dic[var])
            text.insert("end", "剂每百人")
        else:
            text.insert("end", "Error!")
        return var

    button = tkinter.Button(root,
                            text="查询",
                            font=("宋体", "25"),
                            fg="blue",
                            command=search)
    button.pack()
    text = tkinter.Text(root, font=("宋体", "25"), fg="blue")
    text.pack()
    root.mainloop()


def output(cou, data):  # 输出MAX,MIN,AVE(不知道为啥我自己在压缩包里面打开运行中文会出现乱码)
    i = 0
    max = data[i]
    min = data[i]
    sum = 0.0
    ma = 0
    mi = 0
    while (i < len(cou)):
        if (data[i] > max):
            ma = i
            max = data[i]
        if (data[i] < min):
            mi = i
            min = data[i]
        sum = sum + data[i]
        i = i + 1
    print("MAX:", cou[ma], ":", data[ma], "剂每百人")
    print("MIN:", cou[mi], ":", data[mi], "剂每百人")
    print("AVE:", sum / len(cou), "剂每百人")


def bubble_sort(cou, data):  # 冒泡排序
    n = len(cou)
    for j in range(0, n - 1):
        for i in range(0, n - 1 - j):
            if data[i] < data[i + 1]:
                data[i], data[i + 1] = data[i + 1], data[i]
                cou[i], cou[i + 1] = cou[i + 1], cou[i]


def savedata(cou, data):  # 存入excel
    i = 0
    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet('sheet1')
    worksheet.write(0, 0, "国家")
    worksheet.write(0, 1, "接种率/每百人")
    while (i < len(cou)):
        worksheet.write(i + 1, 0, cou[i])
        worksheet.write(i + 1, 1, data[i])
        i = i + 1
    workbook.save("ymjzl.xls")


def plot(data):  # 图像处理
    num = []
    i = 0
    while (i < len(data)):
        num.insert(i, (i + 1))
        i = i + 1
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    plt.xlabel("接种率排名位次")
    plt.ylabel("疫苗接种剂数/百人")
    plt.bar(num, data)
    plt.title('各国疫苗接种率')
    plt.show()


if __name__ == "__main__":
    main()
