' Define and openExcel as an object, then locate the sheet number
Set excelSheet = CreateObject("Excel.Application")
appPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OpenText Functional Testing\Sample Applications\Flight GUI.lnk"
Set dataSheet = excelSheet.Workbooks.Open("C:\rui's\FlightGUI_DataSheet3.0.xlsx")
SheetName = "Sheet1"
' Retrieve the agentName and password to login
agentName = dataSheet.Sheets(SheetName).Cells(2, 1).Value
password = dataSheet.Sheets(SheetName).Cells(2, 2).Value
' Using a loop to go through every row and read their data
For DataRow = 2 To 5
' Read data from specified points
    fromCity = dataSheet.Sheets(SheetName).Cells(DataRow, 3).Value
    toCity = dataSheet.Sheets(SheetName).Cells(DataRow, 4).Value
    flightClass = dataSheet.Sheets(SheetName).Cells(DataRow, 5).Value
    numOfTickets = dataSheet.Sheets(SheetName).Cells(DataRow, 6).Value - 1
    datePicker = dataSheet.Sheets(SheetName).Cells(DataRow, 7).Value
    userName = dataSheet.Sheets(SheetName).Cells(DataRow, 8).Value
' Pre-Login screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\BeforeLogin"&DataRow&".png", True
' Log into the FlightGUI
WpfWindow("OpenText MyFlight Sample").WpfEdit("agentName").Set agentName
WpfWindow("OpenText MyFlight Sample").WpfEdit("password").SetSecure password
' Post-Login screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\AfterLogin"&DataRow&".png", True
WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Click
' Pre-Flight selection screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\BeforeFlightSelection"&DataRow&".png", True
' Performing following actions in FlightGUI    
WpfWindow("OpenText MyFlight Sample").WpfComboBox("fromCity").Select fromCity
WpfWindow("OpenText MyFlight Sample").WpfComboBox("toCity").Select toCity
WpfWindow("OpenText MyFlight Sample").WpfComboBox("Class").Select flightClass
WpfWindow("OpenText MyFlight Sample").WpfCalendar("datePicker").SetDate datePicker
WpfWindow("OpenText MyFlight Sample").WpfComboBox("numOfTickets").Select numOfTickets
' Post-Flight selection screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\AfterFlightSelection"&DataRow&".png", True
' Pre-Ticket selection screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\BeforeTicketSelection"&DataRow&".png", True
' Find the ticket and select the first one to buy 
WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Click
WpfWindow("OpenText MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 1,0
WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Click
' Post-Ticket selection screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\AfterTicketSelection"&DataRow&".png", True
' Pre-Order purchase screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\BeforeTicketSelection"&DataRow&".png", True
' Set the user name and finalize the ticket purchase
WpfWindow("OpenText MyFlight Sample").WpfEdit("passengerName").Set userName
WpfWindow("OpenText MyFlight Sample").WpfButton("ORDER").Click
wait 3
' Post-Order purchase screenshot
WpfWindow("OpenText MyFlight Sample").CaptureBitmap "C:\rui's\UFT_Screenshot\BeforeTicketSelection"&DataRow&".png", True

' Close and reopen the program
WpfWindow("OpenText MyFlight Sample").Close
wait 2
SystemUtil.Run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OpenText Functional Testing\Sample Applications\Flight GUI.lnk"
wait 3
Next
' Exit the program
WpfWindow("OpenText MyFlight Sample").Close
' Path to GUITest1: "C:\Users\User\Documents\Functional Testing\GUITest1" 