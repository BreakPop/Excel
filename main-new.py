import os
<<<<<<< HEAD
=======
import xlwt
>>>>>>> 更新-操作DataFrame进行运算，并把不同大小的Dataframe导入Excel
import pandas  as pd
from pandas import DataFrame as df

def file_name():
    file_name = []
    for files in os.listdir("./"):
<<<<<<< HEAD
        if os.path.splitext(files)[1] == '.xlsx':
=======
        if os.path.splitext(files)[1] == '.xls':
>>>>>>> 更新-操作DataFrame进行运算，并把不同大小的Dataframe导入Excel
            file_name.append(files)
    return file_name   
             

def over(filename):
    for i in filename:
        data = pd.read_excel(i)
        #columname = data.columns.values 查看表头
<<<<<<< HEAD
        data = data.drop([ '点击时间','结算时间', '商品图片', '商品标题', '商品单价', '淘宝订单编号', '淘宝子订单号' , '结算金额'
=======
        data = data.drop(['结算时间', '商品图片', '商品标题', '商品单价', '淘宝订单编号', '淘宝子订单号' , '结算金额'
>>>>>>> 更新-操作DataFrame进行运算，并把不同大小的Dataframe导入Excel
        , '佣金金额' ,'补贴比率', '补贴金额', '补贴类型' ,'收入比率', '分成比率', '提成' ,'技术服务费率', '技术服务费', '付款预估收入',
        '结算预估收入', '产品类型', '推广者身份', '媒体ID' ,'媒体名称', '推广位ID', '推广位名称', '成交平台', '维权标签',
         '专项服务费率' ,'预估专项服务费' ,'结算专项服务费', '渠道关系ID', '会员运营ID', '定金付款时间', '定金淘宝付款时间',
         '定金付款金额'], axis=1)
        df(data).to_excel(i)
<<<<<<< HEAD
        
def main():
    list=file_name()
    print(list)
    over(list)
=======
def write():
    df = pd.read_excel("./东北二妞13.xlsx")
    df1 = df.groupby("订单状态").sum()[["付款金额","商品数量"]]
    df4 = pd.DataFrame({
        '退款率':[''],
        '退货率':['']
        })
    df4.index = ["已失效"]
    df4["退款率"]= df[df["订单状态"]=="已失效"]["付款金额"].sum()/df["付款金额"].sum()
    df4["退货率"]= df[df["订单状态"]=="已失效"]["商品数量"].sum()/df["商品数量"].sum()
    df4["退款率"] = df4["退款率"].apply(lambda x: format(x,'.2%'))
    df4["退货率"] = df4["退货率"].apply(lambda x: format(x,'.2%'))
    #df3 = pd.concat([df1,df4])
    df3=pd.merge(df1,df4,left_index=True,right_index=True,how="outer")
    df3=df3.fillna(value="")
    #pd.concat([df,df3],axis=1,sort=False).to_excel("东北二妞13.xlsx")
    writer = pd.ExcelWriter("东北二妞13.xlsx")
    df.to_excel(writer,"东北二妞13.xlsx")
    df3.to_excel(writer,"东北二妞13.xlsx",startcol=16)
    writer.save()
    print(df1.index)
    #print(df2)
    #print(df3)
    print("----------------")
    print(df4)
    print("----------------")
    #df(df3).to_excel("若谷大大.xls")
        
def main():
    #list=file_name()
    #print(list)
    #over(list)
    write()
>>>>>>> 更新-操作DataFrame进行运算，并把不同大小的Dataframe导入Excel
    print("工作完成")
    
if __name__ == '__main__':
    main()
