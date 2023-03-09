import win32com.client
 
# Check Connectivity
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("This does not connect normally with PLUS. ")
    exit()
 

objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, 'A005930')   # Item Code - Samsung Electronics
objStockMst.BlockRequest()
 
# Current Value Connection and Connection Error Handling
rqStatus = objStockMst.GetDibStatus()
rqRet = objStockMst.GetDibMsg1()
print("Connection status", rqStatus, rqRet)
if rqStatus != 0:
    exit()
 
# Current Query Information
offer = objStockMst.GetHeaderValue(16)  #The selling price
 
import requests
 
def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )
    print(response)
 
myToken = "xoxb-123*****"
 
post_message("xoxb-3637107969666-3622518528199-o5uRo57xlJ2YLeHYJJVOXAa6","#stock","Current price of Samsung Electronics:" + str(offer))

