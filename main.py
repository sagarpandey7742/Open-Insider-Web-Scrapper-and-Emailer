from util import *
import os
import time
from csv import writer
import traceback

dataList = []
if os.path.isfile(outputPath + "output.csv"):
    dataList = generateList([])
print(len(dataList))
# for ele in dataList:
#     print(ele)

url1 = "http://openinsider.com/screener?s=&o=&pl=&ph=&ll=&lh=&fd=1&fdr=&td=0&tdr=&fdlyl=&fdlyh=&daysago=&xp=1" \
       "&excludeDerivRelated=1&vl=" + str(
    value1) + "&vh=&ocl=&och=&sic1=-1&sicl=100&sich=9999&isofficer=1&iscob=1&isceo=1" \
              "&ispres=1&iscoo=1&iscfo=1&isgc=1&isvp=1&grp=0&nfl=&nfh=&nil=&nih=&nol=&noh" \
              "=&v2l=&v2h=&oc2l=&oc2h=&sortcol=0&cnt=100&page=1 "
print(url1)
url2 = "http://openinsider.com/screener?s=&o=&pl=&ph=&ll=&lh=&fd=1&fdr=&td=0&tdr=&fdlyl=&fdlyh=&daysago=&xp=1" \
       "&excludeDerivRelated=1&vl=" + str(
    value2) + "&vh=&ocl=&och=&sic1=-1&sicl=100&sich=9999&isdirector=1&istenpercent=1" \
              "&isother=1&grp=0&nfl=&nfh=&nil=&nih=&nol=&noh=&v2l=&v2h=&oc2l=&oc2h=&sortcol=0&cnt=100&page=1 "
print(url2)
driver = webdriver.Firefox(executable_path=driverPath, options=options)

url = "http://openinsider.com/screener?s=&o=&pl=&ph=&ll=&lh=&fd=3&fdr=&td=0&tdr=&fdlyl=&fdlyh=&daysago=&xp=1&excludeDerivRelated=1&vl=10&vh=&ocl=&och=&sic1=-1&sicl=100&sich=9999&isdirector=1&istenpercent=1&isother=1&grp=0&nfl=&nfh=&nil=&nih=&nol=&noh=&v2l=&v2h=&oc2l=&oc2h=&sortcol=0&cnt=100&page=1"
urls = [url1, url2]

xRow = []
fillingDateRow = []
tradeDateRow = []
tickerRow = []
companyNameRow = []
insiderRow = []
titleRow = []
tradeTypeRow = []
priceRow = []
qtyRow = []
ownedRow = []
detaOwnRow = []
valueRow = []
d1Row = []
w1Row = []
m1Row = []
m6Row = []

for url in enumerate(urls):
    rowFound = False
    try:
        driver.get(url[1])
        time.sleep(sleep)
        subjectParam = url[0]
        table = driver.find_element_by_xpath('//*[@id="tablewrapper"]/table')
        rows = table.find_elements_by_tag_name("tr")[1:]
        print("elements:",len(rows))
        for row in rows:
            print(len(dataList))
            x, fillingDate, tradeDate, ticker, companyName, insider, title, tradeType, price, qty, owned, detaOwn, value, d1, w1, m1, m6 = getElements(
                row)
            if dataList:
                rowFound = checkInCsv(subjectParam, dataList, x, fillingDate, tradeDate, ticker, companyName, insider,
                                      title, tradeType,
                                      price, qty, owned, detaOwn, value, d1, w1, m1, m6)
            else:
                dataList.append(
                    [fillingDate, tradeDate, ticker, insider, title, tradeType, price, qty, owned, detaOwn, value, x])


    except Exception as E:
        print("Exiting Gracefully: ", E)
        traceback.print_exc()


print(len(dataList))
xRow = []
fillingDateRow = []
tradeDateRow = []
tickerRow = []
companyNameRow = []
insiderRow = []
titleRow = []
tradeTypeRow = []
priceRow = []
qtyRow = []
ownedRow = []
detaOwnRow = []
valueRow = []
d1Row = []
w1Row = []
m1Row = []
m6Row = []
for i in range(0, len(dataList)):
    fillingDateRow.append(dataList[i][0])
    tradeDateRow.append(dataList[i][1])
    tickerRow.append(dataList[i][2])
    insiderRow.append(dataList[i][3])
    titleRow.append(dataList[i][4])
    tradeTypeRow.append(dataList[i][5])
    priceRow.append(dataList[i][6])
    qtyRow.append(dataList[i][7])
    ownedRow.append(dataList[i][8])
    detaOwnRow.append(dataList[i][9])
    valueRow.append(dataList[i][10])
    xRow.append(dataList[i][11])

name_of_file = "output"
completeName = os.path.join(outputPath + name_of_file + ".csv")
data = pd.DataFrame({
    "Filling Date": fillingDateRow,
    "Trade Date": tradeDateRow,
    "Ticker": tickerRow,
    "Insider Name": insiderRow,
    "Title": titleRow,
    "Trade Type": tradeTypeRow,
    "Price": priceRow,
    "Qty": qtyRow,
    "Owned": ownedRow,
    "Own Inc": detaOwnRow,
    "Value": valueRow,
    "X": xRow,
})

data.to_csv(completeName, index=False, header=False)
print("File written")

driver.close()
