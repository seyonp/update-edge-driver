Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'C:\Users\MSDemo\Desktop\Edge WebDriver\edge-webdriver-update - Copy.xlsm'!Module1.UpdateEdgeWebDriver"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing