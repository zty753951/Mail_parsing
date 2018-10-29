# -*- coding: utf-8 -*-  
import random
import matplotlib.pyplot as plt
import matplotlib as mpl
from matplotlib.ticker import FuncFormatter, MaxNLocator
from matplotlib.font_manager import FontProperties


mpl.rcParams['font.sans-serif'] = [u'SimHei']
mpl.rcParams['axes.unicode_minus'] = False



# 画折线图
def Broken_line(lineName,x_value,xlabel,y,ylabel):
    x = range(len(x_value))
    y=y
    fig = plt.figure(figsize=(90, 70))


    ax = fig.add_subplot(5,5,1)  # 建立对象     分割3行1列 放第一个

    def format_fn(tick_val, tick_pos):
        if int(tick_val) in x:
            return x_value[int(tick_val)]
        else:
            return ''

    ax.xaxis.set_major_formatter(FuncFormatter(format_fn))
    ax.xaxis.set_major_locator(MaxNLocator(integer=True))

    # 可指定计算机内的任意字体，size为字体大小
    font1 = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=11)
    plt.title(lineName, fontproperties=font1,size=15)
    plt.xlabel(xlabel, fontproperties=font1)
    plt.ylabel(ylabel, fontproperties=font1)

    plt.plot(x, y, 'o-')  # 画图 , label=u"线条"
    ax.legend(prop=font1, loc="upper right")
    # plt.show()
    # plt.savefig("temp.png")


    # 多个图

    ax = fig.add_subplot(5, 5, 2)        #分割3行1列 放第三个

   # 可指定计算机内的任意字体，size为字体大小
    font1 = FontProperties(fname=r"c:\windows\fonts\simsun.ttc", size=11)
    plt.title(lineName, fontproperties=font1,size=15)
    plt.xlabel(xlabel, fontproperties=font1)
    plt.ylabel(ylabel, fontproperties=font1)

    plt.plot(x, y, 'o-')  # 画图 , label=u"线条"
    ax.legend(prop=font1, loc="upper right")
    plt.show()
    # plt.savefig("temp.png")


a=Broken_line("信审日期-进件量",['3/28','3/29','3/30','4/2','4/3','4/4','4/8'],"日期",[1463,2284,2181,3079,3039,3036,3079],"数量")









