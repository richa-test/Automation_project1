 RepositoriesCollection.Add "C:\Users\Admin\Documents\Unified Functional Testing\Automation_proj1\Repositories\desktopFlightGUI_OR.tsr"
 Dim flightDetailsofPlacedOrder
 Dim passengerName
dtfileName="C:\Users\Admin\Documents\Unified Functional Testing\Automation_proj1\TestData\DesktopflightGUI_bookflight.xlsx"
sheetname="searchFlights"
DataTable.AddSheet sheetname
DataTable.Import dtfileName
DataTable.ImportSheet dtfileName,sheetname,sheetname
rowCounts = DataTable.GetSheet(sheetname).GetRowCount

'For i = 1 To rowCounts Step 1
For i = 1 To 2 Step 1
 loginMyFlight()
 searchFlight i,sheetname
 selectFlight()
 name=Trim(CStr(i))
 passengerName="simi " & name
 Set flightDetailsofPlacedOrder = CreateObject("Scripting.Dictionary")
 Set flightDetailsofPlacedOrder=flightDetailsPlaceOrder(passengerName)
 searchOrderCheck(flightDetailsofPlacedOrder)
 closeMyFlightAppln()	
 Next




Function loginMyFlight
SystemUtil.Run "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe"
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set "john"
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "5e789d0a23932bdfda5a"
print "****************loginMyFlight***********************"
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click	
End Function

Function searchFlight(rownmbr,sheetname)
DataTable.LocalSheet.SetCurrentRow(rownmbr)
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select DataTable.GetSheet(sheetname).GetParameter("fromCity").ValueByRow(rownmbr)
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select DataTable.GetSheet(sheetname).GetParameter("toCity").ValueByRow(rownmbr)
WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("datePicker").SetDate  DataTable.GetSheet(sheetname).GetParameter("datePicker").ValueByRow(rownmbr)
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select DataTable.GetSheet(sheetname).GetParameter("class").ValueByRow(rownmbr)
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select DataTable.GetSheet(sheetname).GetParameter("numOfTickets").ValueByRow(rownmbr)
print "***************search flights*********************"
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click	
End Function

Function selectFlight
availableFlightsCount=WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").RowCount
print "**********************select flight***********************"
print "no of flights are " & availableFlightsCount & " selecting first flight"
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectRow(0)
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click
End Function



Function flightDetailsPlaceOrder(passengerName)
flightnmbr=WpfWindow("Micro Focus MyFlight Sample").WpfObject("devname:=flightNumber").getROProperty("text")
print "********************flight details placing order***********************"
print "flightnmbr="& flightnmbr
totalprice=WpfWindow("Micro Focus MyFlight Sample").WpfObject("devname:=totalPrice").getROProperty("text")
print "totalprice= " & totalprice
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set passengerName
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click
wait(5)
If WpfWindow("Micro Focus MyFlight Sample").WpfObject("devname:=orderCompleted").exist(2) Then
Order=WpfWindow("Micro Focus MyFlight Sample").WpfObject("devname:=orderCompleted").getROProperty("text")
print Order
temp=split(Order," ")
ordernumber=temp(1)
print "order number is " & ordernumber	
End If

Dim dict
Set dict = CreateObject("Scripting.Dictionary")
dict.Add "ordernumber",ordernumber
dict.Add "flightnmbr",flightnmbr
dict.Add "totalprice",totalprice
dict.Add "passengerName",passengerName
Set flightDetailsPlaceOrder=dict
End Function

Function searchOrderCheck(flightDetailsPlaceOrder)
print "***********************searchOrderCheck********************"
ordernumber=flightDetailsPlaceOrder.Item("ordernumber")
print "searching for order " & ordernumber
WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click
WpfWindow("Micro Focus MyFlight Sample").WpfTabStrip("WpfTabStrip").Select "SEARCH ORDER"
WpfWindow("Micro Focus MyFlight Sample").WpfRadioButton("byNumberRadio").Set
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("byNumberWatermark").Set ordernumber
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SEARCH").Click
flightnumber=WpfWindow("Micro Focus MyFlight Sample").WpfObject("devname:=flightNumber").getROProperty("text")
passengername=WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").getROProperty("text")
totalprice=WpfWindow("Micro Focus MyFlight Sample").WpfObject("devname:=totalPrice").getROProperty("text")


print "flight details on searching by order number" &"   " &" flight details at time of placing order"
print flightnumber & Space(47-len(flightnumber)) & flightDetailsPlaceOrder.Item("flightnmbr") 
print passengername & Space(47-len(passengername))  & flightDetailsPlaceOrder.Item("passengerName")
print totalprice & Space(47-len(totalprice)) & flightDetailsPlaceOrder.Item("totalprice")
If flightnumber=flightDetailsPlaceOrder.Item("flightnmbr")  and StrComp(passengername,flightDetailsPlaceOrder.Item("passengerName"))=0 and StrComp(totalprice,flightDetailsPlaceOrder.Item("totalprice"))=0 Then
 print "flight booking has correct details"
 else
 print "flight booking has incorrect details and flight details on search are" &"flightnumber =" &flightnumber &"passengername ="&passengername &"totalprice ="&totalprice
 
End If


End Function

Function closeMyFlightAppln
WpfWindow("Micro Focus MyFlight Sample").Close
End Function
 

