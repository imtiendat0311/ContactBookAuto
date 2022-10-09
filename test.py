from tokenize import String
import pandas as pd # import dữ liệu từ excel
from datetime import date # đổi giờ thời gian dateTime()
import time # ngủ =))

excel= pd.read_excel(r"C:\Users\khanh\Desktop\ContactBookAuto-main\lam hop dong.xlsx",sheet_name='24.9',header = None,dtype='object')# header = None tức là k bỏ dòng đầu ß
#name address dob idnumber idissued id issued place,  bank ac , bank name , bank address , swiftcode ,email
c=[0,1,2,3,4,5,6,7,8,9,10] 
for i in excel.values:
     #value = excel.values[i]
    b=[] # tạo list chứa dữ liệu cần dùng cho form 
    for j in c:
        if j == 5: # nếu j=10 => nơi cấp id 
             b[4]+=" at "+i[j] # tại b[4] đang có ngày cấp => thêm at và nơi cấp vào string
        else:
            if isinstance(i[j],str):
                i[j]=i[j].strip()
            if isinstance(i[j],date): # nếu đây là dữ liệu ngày tháng 
                i[j]=i[j].strftime("%d/%m/%Y") # đổi ra đúng format ngày tháng năm 
            b.append(i[j]) # add dữ liệu vào list b
    print(b)
    print("\n")
     # hiển thị dữ liệu list b ra màn hình ( Debug purpose )