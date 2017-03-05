On Error Resume Next
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

Dim reportFolder, resultExcel, sourceExcel, targetExcel
reportFolder="D:\report"
resultExcel="D:\demo.xlsx"
sourceExcel="D:\Src.xlsx"
targetExcel="D:\Src.xlsx"

Set objFSCSS = CreateObject("Scripting.FileSystemObject")
Set objCSSFile = objFSCSS.CreateTextFile(reportFolder+"\style.css")



objCSSFile.WriteLine "* { 	margin: 0; 	padding: 0; }body { 	font: 14px/1.4 Georgia, Serif; }#page-wrap {	margin: 50px;}h1 {	text-align: center;	padding-bottom: 10px;}p 	margin: 20px 0; }	table { 		width: 100%; 		border-collapse: collapse; 	}	tr:nth-of-type(odd) { 		background: #eee; 	}	th { 		background: #333; 		color: white; 		font-weight: bold; 	}	td, th { 		padding: 6px; 		border: 1px solid #ccc; 		text-align: left; 	}.id {	width: 7%;}.name {	width: 20%;}"

Dim objExcel,ObjWorkbook,objsheet
Dim total, passed, failed
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(resultExcel)
set objsheet = objExcel.ActiveWorkbook.Worksheets(1)


total=objsheet.usedrange.rows.count-1
passed=0
failed=0
For i=2 to objsheet.usedrange.rows.count
	If(StrComp(objsheet.Cells(i,4),"Passed",1)=0) Then
		passed=passed+1
	Else
		failed=failed+1
	End If
Next

Set objFSHTML1 = CreateObject("Scripting.FileSystemObject")
Set objHTMLFile1 = objFSHTML1.CreateTextFile(reportFolder+"\index.html")
objHTMLFile1.WriteLine "<html>"
objHTMLFile1.WriteLine "  <head>"
objHTMLFile1.WriteLine "  	<title>WIMI1 VIEW</title>"
objHTMLFile1.WriteLine "	<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">"
objHTMLFile1.WriteLine "	<link rel=""stylesheet"" href=""style.css"">"
objHTMLFile1.WriteLine "    <script type=""text/javascript"" src=""https://www.gstatic.com/charts/loader.js""></script>"
objHTMLFile1.WriteLine "    <script type=""text/javascript"">"
objHTMLFile1.WriteLine "      google.charts.load('current', {'packages':['corechart']});"
objHTMLFile1.WriteLine "      google.charts.setOnLoadCallback(drawChart);"
objHTMLFile1.WriteLine "      function drawChart() {"
objHTMLFile1.WriteLine "        var data = google.visualization.arrayToDataTable(["
objHTMLFile1.WriteLine "          ['Status', 'Number of Test Cases'],"
objHTMLFile1.WriteLine "          ['Passed',     "+CStr(passed)+"],"
objHTMLFile1.WriteLine "          ['Failed',      "+CStr(failed)+"],"
objHTMLFile1.WriteLine "        ]);"
objHTMLFile1.WriteLine "        var options = {  "
objHTMLFile1.WriteLine "        };"
objHTMLFile1.WriteLine "        var chart = new google.visualization.PieChart(document.getElementById('piechart'));"
objHTMLFile1.WriteLine "        chart.draw(data, options);"
objHTMLFile1.WriteLine "      }"
objHTMLFile1.WriteLine "    </script>"
objHTMLFile1.WriteLine "  </head>"
objHTMLFile1.WriteLine "  <body>"
objHTMLFile1.WriteLine "  <center>	"
objHTMLFile1.WriteLine "  <div id=""page-wrap"">"
objHTMLFile1.WriteLine "	<h1 center>WIMI1 VIEW REPORT</h1>	"
objHTMLFile1.WriteLine "	<table>"
objHTMLFile1.WriteLine "		<tr>"
objHTMLFile1.WriteLine "			<th>Test Cases Executed</th>"
objHTMLFile1.WriteLine "			<th>Test Cases Passed</th>"
objHTMLFile1.WriteLine "			<th>Test Cases Failed</th>"
objHTMLFile1.WriteLine "		</tr>"
objHTMLFile1.WriteLine "		<tr>"
objHTMLFile1.WriteLine "			<td>"+CStr(total)+"</td>"
objHTMLFile1.WriteLine "			<td>"+CStr(passed)+"</td>"
objHTMLFile1.WriteLine "			<td>"+CStr(failed)+"</td>"
objHTMLFile1.WriteLine "		</tr>"
objHTMLFile1.WriteLine "	</table>"	
objHTMLFile1.WriteLine "	</div>"
objHTMLFile1.WriteLine "	</center>	"
objHTMLFile1.WriteLine "	<center><a href=""report.html"">View Detailed Report</a></center>"
objHTMLFile1.WriteLine "    <center><div id=""piechart"" style=""width: 900px; height: 500px;""></div></center>"
objHTMLFile1.WriteLine "  </body>"
objHTMLFile1.WriteLine "</html>"



Set objFSHTML = CreateObject("Scripting.FileSystemObject")
Set objHTMLFile = objFSHTML.CreateTextFile(reportFolder+"\report.html")
objHTMLFile.WriteLine "<!DOCTYPE html>"
objHTMLFile.WriteLine "<html>"
objHTMLFile.WriteLine "<head>"
objHTMLFile.WriteLine "	<meta charset='UTF-8'>"	
objHTMLFile.WriteLine "	<title>WIMI1 View Report</title>"
objHTMLFile.WriteLine "	<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">"	
objHTMLFile.WriteLine "	<link rel=""stylesheet"" href=""style.css"">"
objHTMLFile.WriteLine "</head>"
objHTMLFile.WriteLine "<body>"
objHTMLFile.WriteLine "	<div id=""page-wrap"">"
objHTMLFile.WriteLine "	<h1 center>WIMI1 VIEW REPORT</h1>"
objHTMLFile.WriteLine "	<table>"
objHTMLFile.WriteLine "		<tr>"
objHTMLFile.WriteLine "			<th class=""id"">Test Case Id</th>"
objHTMLFile.WriteLine "			<th class=""name"">Test Case Name</th>"
objHTMLFile.WriteLine "			<th class=""fun"">Function Name</th>"
objHTMLFile.WriteLine "			<th class=""status"">Status</th>"
objHTMLFile.WriteLine "			<th class=""eresult"">Expected Result</th>"
objHTMLFile.WriteLine "			<th class=""aresult"">Actual Result</th>"
objHTMLFile.WriteLine "			<th class=""cref"">Source Values</th>"
objHTMLFile.WriteLine "			<th class=""cref"">Target Values</th>"
objHTMLFile.WriteLine "		</tr>"

Set srcExcel = CreateObject("Excel.Application")
Set srcWorkbook = srcExcel.Workbooks.Open(sourceExcel)
Set srcsheet = srcExcel.ActiveWorkbook.Worksheets(1)

Set tarExcel = CreateObject("Excel.Application")
Set tarWorkbook = tarExcel.Workbooks.Open(targetExcel)
Set tarsheet = tarExcel.ActiveWorkbook.Worksheets(1)

For i=2 to objsheet.usedrange.rows.count
	objHTMLFile.WriteLine "		<tr>"
	For j=1 to objsheet.usedrange.columns.count-1
		objHTMLFile.WriteLine "			<td>"+CStr(objsheet.Cells(i,j).value)+"</td>"
	Next
	
	Dim value
	value=objSheet.Cells(i,7).value
	value=Replace(value,"$","")
	Dim cell,row,col
	Dim sourcevalue
	Dim targetvalue
	For k=1 to Len(value)-8
		If(StrComp(Mid(value,k,8),"Src Cell",1)=0)  Then
			cell=Mid(value,k+9,2)
			row=Asc(Mid(cell,1,1))-64
			col=Mid(cell,2,1)
			sourcevalue=sourcevalue+CStr(srcSheet.Cells(1,CInt(row)).value)+" : "+CStr(srcsheet.Cells(CInt(col),CInt(row)).value)+" ; "
		End If
	Next
	For l=1 to Len(value)-8
		If(StrComp(Mid(value,l,8),"Tar Cell",1)=0)  Then
			cell=Mid(value,l+11,2)
			row=Asc(Mid(cell,1,1))-64
			col=Mid(cell,2,1)
			targetvalue=targetvalue+CStr(tarSheet.Cells(1,CInt(row)).value)+" : "+CStr(tarsheet.Cells(CInt(col),CInt(row)).value)+" ; "
		End If
	Next
	objHTMLFile.WriteLine "			<td>"+sourcevalue+"</td>"
	objHTMLFile.WriteLine "			<td>"+targetvalue+"</td>"
	objHTMLFile.WriteLine "		</tr>"
	sourcevalue=""
	targetvalue=""
Next
objHTMLFile.WriteLine "	</table>"	
objHTMLFile.WriteLine "	</div>	"	
objHTMLFile.WriteLine "</body>"
objHTMLFile.WriteLine "</html>"

msgbox "Done"