import pandas as pd
import sklearn
import numpy as np
import sklearn
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt


df = pd.read_excel('TM2.xlsx')
print(df.head(5).to_string())
print(df.shape)



#checks= ['Pass']
#checks = ['Fail']
#df = df[df['Result'].isin (checks)]



reg = LinearRegression()


X = [[
    "جمع کل هزینه ها",
    "هزینه مطالبات مشکوک الوصول",
    "سود پرداختی به ریال", # Direct interest paid on funding
    "کارمزد پرداختی",
    "هزینه کارکنان",
    "هزینه اداری و عمومی",
    "زیان معاملات ارزي",
    "سپرده های بلند مدت ریالی", # Cost of long-term deposits
    "مطالبات معوق",
    "مطالبات مشکوک الوصول",
    "تسهیلات سرنرسیده"
]]
le = LabelEncoder()
df['ABSGainLoss'] = le.fit_transform(df['ABSGainLoss'])

y = df['GainLoss']
#print (type (y))

#y= df['GainLoss']
#y= df['Risk']

reg.fit(X,y)

print(reg.coef_)
print(reg.intercept_)
print(reg.score(X,y))

# plt.xlabel(['DaramadGhMo', 'DramdAmaliati', 'DaramadTamin', 'HazineTamin', 'HazineGheyrAmal', 'HazineAmaliat'])
# plt.ylabel('y')
# plt.scatter(X['DaramadTamin'],y,color='Red',marker='+')
# plt.show()
# plt.scatter(X['DaramadGhMo'],y,color='Red',marker='+')
# plt.show()