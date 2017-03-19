Dim inputCsv, outputCsv, inputPKey, outputPKey, numberOfColumnsInput,numberOfColumnsOutput,reportFolder,hraderMap,transMap,mandFields,optFields,dirColumns,transColumns
inputCsv="D:\data\Source.csv"
outputCsv="D:\data\Target.csv"
headerMap="D:\data\HeaderMap.xlsx"
transMap="D:\data\TransformationMap.xlsx"
mandFields="D:\data\MandatoryFields.csv"
optFields="D:\data\OptionalFields.csv"
dirColumns="D:\data\DirectColumns.csv"
transColumns="D:\data\TransformColumns.csv"
reportFolder="D:\report"
inputPKey=2
outputPKey=3
numberOfColumnsInput=6
numberOfColumnsOutput=8

'inputCsv="C:\Source.csv"
'outputCsv="C:\Target.csv"
'reportFolder="C:\report"
'inputPKey=2
'outputPKey=1
'numberOfColumnsInput=3
'numberOfColumnsOutput=3

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
objHTMLFile.WriteLine "			<th class=""name"">Test Case Name</th>"
objHTMLFile.WriteLine "			<th class=""status"">Status</th>"
objHTMLFile.WriteLine "			<th class=""cref"">Expected Result</th>"
objHTMLFile.WriteLine "			<th class=""cref"">Actual Result</th>"
objHTMLFile.WriteLine "		</tr>"



Set objFSOInput = CreateObject("Scripting.FileSystemObject")
Set inputFile = objFSOInput.OpenTextFile(inputCsv)
Dim rowInput, columnInput
Dim fieldsInput
inputFile.ReadAll
ReDim inputCsvArray(inputFile.Line,numberOfColumnsInput)
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
ReDim outputCsvArray(outputFile.Line,numberOfColumnsOutput)
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

'MsgBox inputCsvArray(2,14)
'MsgBox outputCsvArray(2,34)
Dim objexcel,headerMapWorkbook,headerMapSheet
Set objexcel = Createobject("Excel.Application")
Set headerMapWorkbook = objExcel.WorkBooks.Open(headerMap)
Set headerMapSheet = headerMapWorkbook.Worksheets("Sheet1")
ReDim inputHeaderMap(headerMapSheet.usedrange.rows.count)
ReDim outputHeaderMap(headerMapSheet.usedrange.rows.count)
For x=2 to headerMapSheet.usedrange.rows.count
	inputHeaderMap(x-2)=headerMapSheet.Cells(x,1).value
	outputHeaderMap(x-2)=headerMapSheet.Cells(x,2).value
Next

Dim prevOutputCol,tcStatus,inputValue,outputValue,tcName,expResult,actResult


'Mnadatory Fields Validation


tcName="TC_MandatoryFieldsCheck"

Set objFSOOutput = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSOOutput.OpenTextFile(mandFields)
outputFile.ReadAll
Dim mandFieldsArray
outputFile.close
Set outputFile = objFSOOutput.OpenTextFile(mandFields)
mandFieldsArray=Split(outputFile.Readline,",")
expResult="All mandatory fields should be updated in output file"
actResult=""
For x=0 to UBound(mandFieldsArray)
	founderFlag=false
	For y=0 to UBound(outputCsvArray,2)
		If(StrComp(mandFieldsArray(x),outputCsvArray(0,y),1)=0) Then
			founderFlag= true
		End If
	Next
	If(founderFlag=false) Then
		Dim inputIndex
		inputIndex=-1
		For z=0 to UBound(inputHeaderMap)
			If(StrComp(mandFieldsArray(x),inputHeaderMap(z),1)=0) Then
				inputIndex=z
			End If
		Next
		If(inputIndex<>-1) Then
			For z=0 to UBound(outputCsvArray,2)
				If(StrComp(outputHeaderMap(inputIndex),outputCsvArray(0,z),1)=0) Then
					founderFlag=true
				End If
			Next
		End If
	End If
	If(founderFlag=false) Then
		actResult=actResult+mandFieldsArray(x)+" , "
	End If
Next
If(StrComp(actResult,"",1)<>0) Then
	tcStatus="FAIL"
	actResult = actResult+" Not Found"
Else
	tcStatus="PASS"
	actResult="All mandatory fields are updated in output file"
End If
objHTMLFile.WriteLine "		<tr>"
objHTMLFile.WriteLine "			<td>"+tcName+"</td>"
	objHTMLFile.WriteLine "			<td>"+tcStatus+"</td>"
	objHTMLFile.WriteLine "			<td>"+expResult+"</td>"
	objHTMLFile.WriteLine "			<td>"+actResult+"</td>"
	objHTMLFile.WriteLine "		</tr>"
'MsgBox actResult



'Mnadatory Fields Validation

'Optional Fields Validation

tcName="TC_OptionalFieldsCheck"

Set objFSOOutput = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSOOutput.OpenTextFile(optFields)
outputFile.ReadAll
Dim optFieldsArray
outputFile.close
Set outputFile = objFSOOutput.OpenTextFile(optFields)
optFieldsArray=Split(outputFile.Readline,",")
expResult="All optional fields should be updated in output file"
actResult=""
For x=0 to UBound(optFieldsArray)
	founderFlag=false
	For y=0 to UBound(outputCsvArray,2)
		If(StrComp(optFieldsArray(x),outputCsvArray(0,y),1)=0) Then
			founderFlag= true
		End If
	Next
	If(founderFlag=false) Then
		Dim inputIndexOpt
		inputIndexOpt=-1
		For z=0 to UBound(inputHeaderMap)
			If(StrComp(optFieldsArray(x),inputHeaderMap(z),1)=0) Then
				inputIndexOpt=z
			End If
		Next
		If(inputIndexOpt<>-1) Then
			For z=0 to UBound(outputCsvArray,2)
				If(StrComp(outputHeaderMap(inputIndexOpt),outputCsvArray(0,z),1)=0) Then
					founderFlag=true
				End If
			Next
		End If
	End If
	If(founderFlag=false) Then
		actResult=actResult+optFieldsArray(x)+" , "
	End If
Next
If(StrComp(actResult,"",1)<>0) Then
	tcStatus="FAIL"
	actResult = actResult+" Not Found"
Else
	tcStatus="PASS"
	actResult="All optinal fields are updated in output file"
End If
objHTMLFile.WriteLine "		<tr>"
objHTMLFile.WriteLine "			<td>"+tcName+"</td>"
	objHTMLFile.WriteLine "			<td>"+tcStatus+"</td>"
	objHTMLFile.WriteLine "			<td>"+expResult+"</td>"
	objHTMLFile.WriteLine "			<td>"+actResult+"</td>"
	objHTMLFile.WriteLine "		</tr>"
'MsgBox actResult

'Optional Fields Validation


'Non-transformation Validation

tcName="TC_NonTransformableFields"

Set objFSOOutput = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSOOutput.OpenTextFile(dirColumns)
outputFile.ReadAll
Dim dirColumnArray
outputFile.close
Set outputFile = objFSOOutput.OpenTextFile(dirColumns)
dirColumnArray=Split(outputFile.Readline,",")
expResult="All columns should be updated in output file"
actResult=""
Dim inputCol,outputCol
For ivar=0 to UBound(dirColumnArray)
	For j=0 to numberOfColumnsInput
		If(StrComp(dirColumnArray(ivar),inputCsvArray(0,j),1)=0) Then
			inputCol=j
		End If
	Next
	Dim outputColValue
	For j=0 to UBound(inputHeaderMap)
		If(StrComp(dirColumnArray(ivar),inputHeaderMap(j),1)=0) Then
			outputColValue=outputHeaderMap(j)
		'	Exit For
		End If
	Next
	For j=0 to numberOfColumnsOutput
		If(StrComp(outputColValue,outputCsvArray(0,j),1)=0) Then
			outputCol=j
		End If
	Next
	
	If(StrComp(inputCsvArray(1,inputCol),outputCsvArray(1,outputCol),1)<>0) Then
		actResult=actResult+inputCsvArray(0,inputCol)+" = "+inputCsvArray(1,inputCol)+" , "+outputCsvArray(0,outputCol)+" = "+outputCsvArray(1,outputCol)+" ; "
	End If
Next

If(StrComp(actResult,"",1)=0) Then
	actResult="All non-transformable fields are updated in output file"
	tcStatus="PASS"
Else
	tcStatus="FAIL"
End If
objHTMLFile.WriteLine "		<tr>"
objHTMLFile.WriteLine "			<td>"+tcName+"</td>"
	objHTMLFile.WriteLine "			<td>"+tcStatus+"</td>"
	objHTMLFile.WriteLine "			<td>"+expResult+"</td>"
	objHTMLFile.WriteLine "			<td>"+actResult+"</td>"
	objHTMLFile.WriteLine "		</tr>"
'MsgBox actResult


'Non-transformation Validation

'Transformation Validation

tcName="TC_TransformableFields"

Set objFSOOutput = CreateObject("Scripting.FileSystemObject")
Set outputFile = objFSOOutput.OpenTextFile(transColumns)
outputFile.ReadAll
Dim transColumnArray
outputFile.close
Set outputFile = objFSOOutput.OpenTextFile(transColumns)
transColumnArray=Split(outputFile.Readline,",")


Dim objexcel1,transMapWorkbook,transMapSheet
Set objexcel1 = Createobject("Excel.Application")
Set transMapWorkbook = objExcel.WorkBooks.Open(transMap)
Set transMapSheet = transMapWorkbook.Worksheets("Sheet1")
ReDim columnHeadMap(transMapSheet.usedrange.rows.count)
ReDim inputValueMap(transMapSheet.usedrange.rows.count)
ReDim outputValueMap(transMapSheet.usedrange.rows.count)
For x=2 to transMapSheet.usedrange.rows.count
	columnHeadMap(x-2)=transMapSheet.Cells(x,1).value
	inputValueMap(x-2)=transMapSheet.Cells(x,2).value
	outputValueMap(x-2)=transMapSheet.Cells(x,3).value
Next
expResult="All transformable fields should be updated in output file"
actResult=""
For ivar=0 to UBound(transColumnArray)
	For j=0 to numberOfColumnsInput
		If(StrComp(transColumnArray(ivar),inputCsvArray(0,j),1)=0) Then
			inputCol=j
		End If
	Next
	For j=0 to UBound(inputHeaderMap)
		If(StrComp(transColumnArray(ivar),inputHeaderMap(j),1)=0) Then
			outputColValue=outputHeaderMap(j)
			Exit For
		End If
	Next
	For j=0 to numberOfColumnsOutput
		If(StrComp(outputColValue,outputCsvArray(0,j),1)=0) Then
			outputCol=j
		End If
	Next
	Dim outputTransValue
	For j=0 to UBound(columnHeadMap)
		If(StrComp(inputCsvArray(0,inputCol),columnHeadMap(j),1)=0) Then
			If (StrComp(inputCsvArray(1,inputCol),inputValueMap(j),1)=0) Then
				outputTransValue=outputValueMap(j)
				Exit For
			End If
		End If
	Next
	If(StrComp(outputTransValue,outputCsvArray(1,outputCol),1)<>0) Then
		actResult=actResult+"Expected "+outputCsvArray(0,outputCol)+" = "+outputTransValue+" , Actual "+outputCsvArray(0,outputCol)+" = "+outputCsvArray(1,outputCol)+" ; "
	End If
Next

If(StrComp(actResult,"",1)=0) Then
	actResult="All transformable fields are updated in output file"
	tcStatus="PASS"
Else
	tcStatus="FAIL"
End If
objHTMLFile.WriteLine "		<tr>"
objHTMLFile.WriteLine "			<td>"+tcName+"</td>"
	objHTMLFile.WriteLine "			<td>"+tcStatus+"</td>"
	objHTMLFile.WriteLine "			<td>"+expResult+"</td>"
	objHTMLFile.WriteLine "			<td>"+actResult+"</td>"
	objHTMLFile.WriteLine "		</tr>"
'MsgBox actResult


'Transformation Validation

objHTMLFile.WriteLine "	</table>"	
objHTMLFile.WriteLine "	</div>	"	
objHTMLFile.WriteLine "</body>"
objHTMLFile.WriteLine "</html>"

msgbox "Done"
