import hashlib
# import datetime
seral = input("input:请输入待注册机子的机器码：")
# 获取当前时间
# now = datetime.datetime.now()

# 获取年份和月份
# year = now.year
# month = now.month
# seral = seral + "Chenxuan" + str(year) + str(month)
seral = seral + "Chenxuan" + "202212"
print(hashlib.sha256(seral.encode('utf-8')).hexdigest())