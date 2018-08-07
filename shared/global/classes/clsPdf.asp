<!--#include virtual="/shared/global/server_includes/http_upload.asp"-->
<!--#include virtual="/shared/global/server_includes/fax.asp"-->

<%
'Put in Header
'Dim Pdf
'Set Pdf = New clsPdf

'Use in Code
'Pdf.FileName = "Order 123456-0.pdf"
'Pdf.AddForm FormPath 
'Pdf.AddField FieldName, FieldValue
	
'Pdf.SaveAs SavePath 'To save a permanent copy to another folder
'Pdf.Display 'Binary writes the PDF to a browser
'Pdf.Download 'Same as display, but does not open inline
'Pdf.Print PrinterIP 'Uploads the PDF to a printer
'Pdf.Fax FromName, FromEmail, ToName, ToCompany, ToFaxNumber, ToPhoneNumber, BccName, BccEmail, Subject, BillBackCode, Company
'Pdf.Email FromName, FromEmail, ToName, ToEmail, CcName, CcEmail, BccName, BccEmail, Subject

'Put in Footer
'Pdf.End
'Set Pdf = Nothing

Class clsPdf

	Public FileName

    Private SeedFileName, TempFileName, TempPath, PDF, FSO, Doc, Page

    Private Sub Class_Initialize()
		Set PDF = Server.CreateObject("Persits.PDF")
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		SeedFileName = "TempPdf" & "_" & Session.SessionID & "_" & GetRandomNumber(1000000, 9999999) & ".pdf"
		TempFileName = SeedFileName
		FileName = SeedFileName
		TempPath = "C:\PdfTempStorage\"
		If Not FSO.FolderExists(TempPath) Then
			FSO.CreateFolder(TempPath)
		End If
		Set Doc = PDF.CreateDocument()
		Doc.ModDate = Now()
		Doc.Creator = "Refron Website Auto-Generated"
		Doc.Producer = "Persits AspPDF " & PDF.Version
    End Sub

	Private Sub Class_Terminate()
		Doc.Close
		Set Doc = Nothing
		Set PDF = Nothing
		FSO.DeleteFile TempPath & Replace(SeedFileName, ".pdf", "*.pdf"), True
		Set FSO = Nothing
	End Sub

	Public Sub AddForm(FormPath)
		If Doc.Pages.COunt > 0 Then
			Call Save()
			Doc.Close
			Set Doc = PDF.OpenDocument(TempPath & TempFileName)
		End If
		Doc.AppendDocument(PDF.OpenDocument(FormPath))
		Call Save()
		Doc.Close
		Set Doc = PDF.OpenDocument(TempPath & TempFileName)
		Doc.Form.RemoveXFA
	End Sub
	
	Public Sub AddBarcode(Data, Page, Parameters)
		Set Page = Doc.Pages(Page)
		Page.Canvas.DrawBarcode Data, Parameters
	End Sub

	Public Sub AddField(Name, Value, FontName)
		Dim Field, Font, Child
		Set Field = Doc.Form.FindField(Name)
		If Not Field Is Nothing Then
			If FontName = "" Then FontName = "Arial"
			Set Font = Doc.Fonts(FontName)
			If Field.Children.Count > 0 Then
				For Each Child In Field.Children
					If IsNull(Value) Then
						Child.SetFieldValue "", Font
					Else
						Child.SetFieldValue Value, Font
					End If
				Next
			Else
				If IsNull(Value) Then
					Field.SetFieldValue "", Font
				Else
					Field.SetFieldValue Value, Font
				End If
			End If
		End If
	End Sub

	Public Sub SaveAs(SavePath)
		Call Save()
		FSO.CopyFile TempPath & TempFileName, SavePath & FileName, True
	End Sub

	Public Sub Display()
		Doc.SaveHttp "filename=" & FileName
	End Sub

	Public Sub Download()
		Doc.SaveHttp "attachment; filename=" & FileName
	End Sub

	Public Sub [Print](PrinterIP)
		Dim DestURL
		DestURL = "http://" & PrinterIP & ":7627/hp/device/this.printservice?printThis"
		Call Save()
		UploadFile DestURL, TempPath & TempFileName
	End Sub

	Public Sub Fax(FromName, FromEmail, ToName, ToCompany, ToFaxNumber, ToPhoneNumber, BccName, BccEmail, Subject, BillBackCode, Company)
		Dim strToField : strToField = ""
		strToField = strToField & "{r=f}" 'resolution is fine
		strToField = strToField & "{rt=4}" 'retry x times - was 5
		strToField = strToField & "{ri=300}" 'retry interval (min) - was 3
		strToField = strToField & "{in=y}" 'send confirmation
		strToField = strToField & "{af=y}" 'attach fax to confirmation
		strToField = strToField & "{df=n}" 'delete fax after confirmation sent
		strToField = strToField & "{c=n}" 'do not use a cover page
		strToField = strToField & "{b=" & BillBackCode & "}" 'tag bill back code
		strToField = strToField & "{n=" & StripNonAlphaNumericCharacters(ToName) & "}"
		strToField = strToField & "{cn=" & StripNonAlphaNumericCharacters(ToCompany) & "}"
		strToField = strToField & "{ph=" & Left(FormatPhone(ToPhoneNumber, "PPP-PPP-PPPP"), 12) & "}"
		strToField = GetToAddress(strToField, FixFaxNumber(ToFaxNumber))
		Call Email(FromName, GetFromAddress(FromEmail), "", strToField, "", "", "", "", GetSubject(Subject))
	End Sub
	
	Public Sub Email(FromName, FromEmail, ToName, ToEmail, CcName, CcEmail, BccName, BccEmail, Subject)
		Dim Mail
		Set Mail = Server.CreateObject("Persits.MailSender")
		'Mail.Host = "localhost"
        Mail.Host = "mail.refron.com"
		Mail.From = FromEmail
		Mail.FromName = FromName
		Mail.AddAddress ToEmail, ToName
		If CcEmail <> "" Then
			Mail.AddCc CcEmail, CcName
		End If
		If BccEmail <> "" Then
			Mail.AddBcc BccEmail, BccName
		End If
		Mail.Subject = Subject
		Call SaveAs(TempPath)
		Mail.AddAttachment(TempPath & FileName)
		Mail.Send
		FSO.DeleteFile TempPath & FileName, True
		Set Mail = Nothing
	End Sub

	Private Function GetRandomNumber(ByVal LowerLimit, ByVal UpperLimit)
		Randomize()
		GetRandomNumber = Int((UpperLimit - LowerLimit + 1) * Rnd() + LowerLimit)
	End Function
	
	Private Sub Save()
		TempFileName = Doc.Save(TempPath & TempFileName, False)
	End Sub

End Class
%>