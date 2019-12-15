import os
import pandas  as pd
from pandas import DataFrame as df

def file_name():
    file_name = []
    for files in os.listdir("./"):
        if os.path.splitext(files)[1] == '.xlsx':
            file_name.append(files)
    return file_name   
             

def over(filename):
    for i in filename:
        data = pd.read_excel(i)
        #columname = data.columns.values 查看表头
        data = data.drop([ '点击时间','结算时间', '商品图片', '商品标题', '商品单价', '淘宝订单编号', '淘宝子订单号' , '结算金额'
        , '佣金金额' ,'补贴比率', '补贴金额', '补贴类型' ,'收入比率', '分成比率', '提成' ,'技术服务费率', '技术服务费', '付款预估收入',
        '结算预估收入', '产品类型', '推广者身份', '媒体ID' ,'媒体名称', '推广位ID', '推广位名称', '成交平台', '维权标签',
         '专项服务费率' ,'预估专项服务费' ,'结算专项服务费', '渠道关系ID', '会员运营ID', '定金付款时间', '定金淘宝付款时间',
         '定金付款金额'], axis=1)
        df(data).to_excel(i)
        
def main():
    list=file_name()
    print(list)
    over(list)
    print("工作完成")
    
if __name__ == '__main__':
    main()
