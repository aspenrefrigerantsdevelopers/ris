<!--#include virtual="/shared/global/server_includes/email.asp"-->

<%
'---- ExecuteOptionEnum Values ----
Const adAsyncExecute = &H00000010
Const adAsyncFetch = &H00000020
Const adAsyncFetchNonBlocking = &H00000040
Const adExecuteNoRecords = &H00000080

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- ADO Cursor Type Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- ADO Lock Type Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- ADO Command Type Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004

'---- DataTypeEnum Values ----
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

'---- FilterGroupEnum Values ----
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adFilterAffectedRecords = 2
Const adFilterFetchedRecords = 3
Const adFilterPredicate = 4

'---- Common database variables ----
Dim strSql

'---- Connection objects & their open flags ----
Dim conAS400, conOpenAS400 : conOpenAS400 = False
Dim conRefronShape, conOpenRefronShape : conOpenRefronShape = False
Dim conRefron, conOpenRefron : conOpenRefron = False
Dim conImaging, conOpenImaging : conOpenImaging = False
Dim conViewPoint, conOpenViewPoint : conOpenViewPoint = False
Dim conRefronNative, conOpenRefronNative : conOpenRefronNative = False

'---- Creates database connection & runs database check sub routine ----
Sub ConnectToDb(varDb)
	Select Case varDb
	    Case "AS400"
			If Not conOpenAS400 Then
			    Set conAS400 = Server.CreateObject("ADODB.Connection")
			    conAS400.ConnectionTimeout = 5
			    conAS400.CommandTimeout = 30
			    conAS400.ConnectionString = "DSN=AS400_W21; UID=LICWEB; PWD=K32JB49"
			    conAS400.Open
			    If conAS400.State = 1 Then conOpenAS400 = True
	        End If
		Case "RefronShape"
			If Not conOpenRefronShape Then
			    Set conRefronShape = Server.CreateObject("ADODB.Connection")
			    conRefronShape.ConnectionTimeout = 5
			    conRefronShape.CommandTimeout = 30
			    conRefronShape.ConnectionString = "Provider=MSDataShape;Data Provider=SQLOLEDB;Data Source=10.96.185.48;Initial Catalog=Refron;User ID=sa;Password=2Azx58cbQ;"
			    conRefronShape.Open
			    If conRefronShape.State = 1 Then conOpenRefronShape = True
	        End If
	    Case "Refron"
			If Not conOpenRefron Then
			    Set conRefron = Server.CreateObject("ADODB.Connection")
			    conRefron.ConnectionTimeout = 5
			    conRefron.CommandTimeout = 30
			    conRefron.ConnectionString = "Provider=SQLOLEDB.1;Data Source=10.96.185.48;Initial Catalog=Refron;User ID=Intranet;Password=Temp123;"
			    conRefron.Open
			    If conRefron.State = 1 Then conOpenRefron = True
	        End If
	    Case "Imaging"
			If Not conOpenImaging Then
			    Set conImaging = Server.CreateObject("ADODB.Connection")
			    conImaging.ConnectionTimeout = 5
			    conImaging.CommandTimeout = 30
			    conImaging.ConnectionString = "Provider=SQLOLEDB.1;Data Source=10.96.185.18;Initial Catalog=Imaging;User ID=sa;Password=openup;"
			    conImaging.Open
			    If conImaging.State = 1 Then conOpenImaging = True
	        End If
	    Case "ViewPoint"
			If Not conOpenViewPoint Then
			    Set conViewPoint = Server.CreateObject("ADODB.Connection")
			    conViewPoint.ConnectionTimeout = 5
			    conViewPoint.CommandTimeout = 30
			    conViewPoint.ConnectionString = "Provider=SQLOLEDB.1;Data Source=10.96.185.48;Initial Catalog=sgmsdb;User ID=viewpoint;Password=Refron3818Admin33;"
			    conViewPoint.Open
			    If conViewPoint.State = 1 Then conOpenViewPoint = True
	        End If
	    Case "RefronNative"
			If Not conOpenRefronNative Then
			    Set conRefronNative = Server.CreateObject("ADODB.Connection")
			    conRefronNative.ConnectionTimeout = 5
			    conRefronNative.CommandTimeout = 30
			    conRefronNative.ConnectionString = "Provider=SQLNCLI;Data Source=10.96.185.48;Initial Catalog=Refron;User ID=Intranet;Password=Temp123;"
			    conRefronNative.Open
			    If conRefronNative.State = 1 Then conOpenRefronNative = True
	        End If
	End Select
End Sub

Sub DeleteAllParameters(cmd)
	Dim i
	If cmd.Parameters.Count > 0 Then
		For i = 0 To (cmd.Parameters.Count - 1)
			cmd.Parameters.Delete 0
		Next
	End If
End Sub
%>