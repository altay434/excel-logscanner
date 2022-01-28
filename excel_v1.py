import openpyxl as open

folder = open.load_workbook("raporDeneme.xlsx")
ws = folder.active


def extractData(baslangic: str,bitis: str) -> tuple: # This function extract datas between the start cell and finish cell.
    return ws[baslangic:bitis]

def parsHour(tupleList: tuple) -> list:             # In my excel file, I have to pars my data like that. You can customize it for your usage.
    hourList = []
    for dataTuple in tupleList:
        data = str(dataTuple[0].value).split(" ")
        data = data[1]
        data = data.split(":")
        data = data[0]
        hourList.append(data)
    
    return hourList

def parsMinute(tupleList: tuple) -> list:         # Same like parsHour
    minuteList = []
    for dataTuple in tupleList:
        data = str(dataTuple[0].value).split(":")
        data = data[1]
        minuteList.append(data)
    
    return minuteList


def minuteCheck(minuteList: list,index1: int,index2: int) -> bool: # This function compares two minute values from 2 data.
    if (minuteList[index1] == minuteList[index2]):
        return True
            
            

def hourCheck(hourList: list,index1: int,index2: int) -> bool: # same like minuteCheck
    if (hourList[index1] == hourList[index2]):
        return True
                
    
    

def nameIndexGetter(nameList: tuple) -> list:     # This function gives the indexes of the same names. I put same names to the list. Also this list in nameIndexes.
    nameIndexes = []
    indexFirst = 0
    for dataFirst in nameList:
        indexSecond = 0
        for dataSecond in nameList:
            if (dataFirst[0].value == dataSecond[0].value) and (indexFirst != indexSecond):
                nameIndexes.append([indexFirst,indexSecond])
            indexSecond += 1    
        indexFirst +=1  
    return nameIndexes
        
    
def check(nameList: tuple,hourList: list,minuteList: list): # And finally here i use all of the check methods. I use checked list because i dont want to print same names again and again.
    checkedList = []
    nameIndexes = nameIndexGetter(nameList)
    for data in nameIndexes:
        if hourCheck(hourList,data[0],data[1]):
            if minuteCheck(minuteList,data[0],data[1]):
                if checkedList.count(nameList[data[0]][0].value) == 0:
                    checkedList.append(nameList[data[0]][0].value)
    for data in checkedList:
        print(f"The student named '{data}' used his student card multiple times. ")
    
def printTuple(tupleList: tuple): #Quick method for printing tuples.
    for dataTuple in tupleList:
        print(dataTuple[0].value)

def printList(normalList: list): #Quick method for printing tuples.
    for dataList in normalList:
        print(dataList)


def main():
    
    nameList = extractData("I14","I2060")
    dateList = extractData("D14","D2060")
    hourList = parsHour(dateList)
    minuteList = parsMinute(dateList)
    check(nameList,hourList,minuteList)
    input("\n You can press any button for quit.")

    
    

main()




