import DataAnalysis
import Main

menu = ("1)Add New full data list\n"
        "2)Add New uncompleted data\n"
        "3)Analysis\n"
        "4)Show graphic")


answer = input(menu)
if(answer=="1"):
    Main.addNewData(1,True)
elif(answer=="2"):
    Main.addNewData(1,False)
elif(answer=="3"):
    DataAnalysis.doAnalysis()
elif(answer=="4"):
    DataAnalysis.show()


