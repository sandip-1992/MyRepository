Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20


Dim inputCsv, outputCsv, inputPKey, outputPKey, numberOfColumnsInput,numberOfColumnsOutput,reportFolder
inputCsv="D:\Source.csv"
outputCsv="D:\Target.csv"
reportFolder="D:\report"
inputPKey=2
outputPKey=1
numberOfColumnsInput=3
numberOfColumnsOutput=3

Set objFSCSS = CreateObject("Scripting.FileSystemObject")
Set objCSSFile = objFSCSS.CreateTextFile(reportFolder+"\style.css")



objCSSFile.WriteLine "* { 	margin: 0; 	padding: 0; }body { 	font: 14px/1.4 Georgia, Serif; }#page-wrap {	margin: 50px;}h1 {	text-align: center;	padding-bottom: 10px;}p 	margin: 20px 0; }	table { 		width: 100%; 		border-collapse: collapse; 	}	tr:nth-of-type(odd) { 		background: #eee; 	}	th { 		background: #333; 		color: white; 		font-weight: bold; 	}	td, th { 		padding: 6px; 		border: 1px solid #ccc; 		text-align: left; 	}.id {	width: 7%;}.name {	width: 20%;}"


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
objHTMLFile.WriteLine "			<th class=""cref"">Source Values</th>"
objHTMLFile.WriteLine "			<th class=""cref"">Target Values</th>"
objHTMLFile.WriteLine "		</tr>"



Set objFSOInput = CreateObject("Scripting.FileSystemObject")
Set inputFile = objFSOInput.OpenTextFile(inputCsv)
Dim rowInput, columnInput
Dim fieldsInput
inputFile.ReadAll
ReDim inputCsvArray(inputFile.Line-1,numberOfColumnsInput)
inputFile.close
Set inputFile = objFSOInput.OpenTextFile(inputCsv)
Do Until inputFile.AtEndOfStream  
    fieldsInput = Split(inputFile.Readline,",") 'store line in temp array  
    For columnInput = 0 To UBound(fieldsInput) 'iterate through the fields of the temp array  
        inputCsvArray(rowInput,columnInput) = fieldsInput(columnInput) 'store each field in the 2D array with the given coordinates  
    Next
    rowInput = rowInput + 1  'next line 
Loop

inputFile.close

Set objFSOOutput = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSOOutput.OpenTextFile(outputCsv)
Dim rowOutput, columnOutput
Dim fieldsOutput
outputFile.ReadAll
ReDim outputCsvArray(outputFile.Line-1,numberOfColumnsInput)
outputFile.close
Set outputFile = objFSOOutput.OpenTextFile(outputCsv)
Do Until outputFile.AtEndOfStream  
    fieldsOutput = Split(outputFile.Readline,",") 'store line in temp array  
    For columnOutput = 0 To UBound(fieldsOutput) 'iterate through the fields of the temp array  
        outputCsvArray(rowOutput,columnOutput) = fieldsOutput(columnOutput) 'store each field in the 2D array with the given coordinates  
    Next
    rowOutput = rowOutput + 1  'next line 
Loop

outputFile.close



Dim outputCol,tcStatus,inputValue,outputValue
Dim total,passed,failed

For i=0 to numberOfColumnsInput
	tcStatus="Passed"
	inputValue=""
	outputValue=""
	For j=0 to numberOfColumnsOutput
		If(StrComp(inputCsvArray(0,i),outputCsvArray(0,j),1)=0) Then
			outputCol=j
		End If
	Next
	objHTMLFile.WriteLine "		<tr>"
	objHTMLFile.WriteLine "			<td>"+CStr(i+1)+"</td>"
	objHTMLFile.WriteLine "			<td>"+CStr("TC_"+inputCsvArray(0,i))+"</td>"
	objHTMLFile.WriteLine "			<td>"+CStr(inputCsvArray(0,i))+"</td>"
	For k=0 to UBound(inputCsvArray,2)
		For l=0 to UBound(outputCsvArray,2)
			if(StrComp(inputCsvArray(k,inputPKey),outputCsvArray(l,outputPKey),1)=0) Then
				if(StrComp(inputCsvArray(k,i),outputCsvArray(l,outputCol),1)<>0) Then
					tcStatus="Failed"
					inputValue=inputValue+inputCsvArray(k,inputPKey)+" = "+inputCsvArray(k,i)+" ; "
					outputValue=outputValue+outputCsvArray(l,outputPKey)+" = "+outputCsvArray(l,outputCol)+" ; "
				End If
			End If
		Next
	Next
	objHTMLFile.WriteLine "			<td>"+tcStatus+"</td>"
	objHTMLFile.WriteLine "			<td>"+inputValue+"</td>"
	objHTMLFile.WriteLine "			<td>"+outputValue+"</td>"
	objHTMLFile.WriteLine "		</tr>"
	total=total+1
	If(StrComp(tcStatus,"Passed",1)=0) Then
		passed=passed+1
	End If
	If(StrComp(tcStatus,"Failed",1)=0) Then
		failed=failed+1
	End If
Next

objHTMLFile.WriteLine "	</table>"	
objHTMLFile.WriteLine "	</div>	"	
objHTMLFile.WriteLine "</body>"
objHTMLFile.WriteLine "</html>"


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



msgbox "Done"
