Dim objuft


Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.launch
objuft.open("C:\Users\sfjbs\Documents\UFT One\DATA_Driven Framework\Driver\Data Driver")
objuft.Test.Run
objuft.Test.close
objuft.quit
set objuft=nothing