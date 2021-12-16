from datetime import datetime, date

# 计算时间差的分钟数
# 同一天的时间差
time_1 = '2021-08-10 16:44:35'
time_2 = '2021-08-10 17:15:31'

time_1_struct = datetime.strptime(time_1, "%Y-%m-%d %H:%M:%S")
time_2_struct = datetime.strptime(time_2, "%Y-%m-%d %H:%M:%S")
seconds = (time_2_struct - time_1_struct).seconds
print('同一天的秒数为：')
print(seconds)
