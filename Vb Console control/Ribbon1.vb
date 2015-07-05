﻿Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
	Public xmlPart As Office.CustomXMLPart
	Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

	End Sub

	Private Sub DouserEnable_Click(sender As Object, e As RibbonControlEventArgs) Handles DouserCtrlEnable.Click
		'if Douser Enable is not true disable all controls
		Dim status As Boolean
		Dim ctrl

		status = DouserCtrlEnable.Checked
		OpenTime.Enabled = False
		CloseTime.Enabled = False

		For Each ctrl In Douser_Info.Items
			ctrl.Enabled = status
		Next

		For Each ctrl In Console_Settings.Items
			ctrl.Enable = status
		Next
	End Sub

	Private Sub IP_Address_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles IP_Address.TextChanged
		Dim pres As PowerPoint.Presentation = Globals.ThisAddIn.Application.ActivePresentation

		Dim xml As Office.CustomXMLPart = pres.CustomXMLParts.SelectByID( _
			pres.CustomDocumentProperties.Item("douser_controls").Value _
			)
		Dim item As Office.CustomXMLNode = xml.SelectSingleNode("/douser_controls/console/ip")

		If System.Net.IPAddress.TryParse(IP_Address.Text, Nothing) Then
			item.Text = IP_Address.Text
		Else
			MsgBox("Please enter a valid IP address")
			IP_Address.Text = item.Text
		End If
	End Sub

	Private Sub Port_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Port.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("^\d")
		Dim strVal As String = Port.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer
		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal <= ThisAddIn.MAX_PORT_NO And intVal > 0 Then
				Globals.ThisAddIn.setPort(intVal)
			Else
				MsgBox("Please enter a valid port number")
			End If
		Else
			MsgBox("Please enter a valid port number")
		End If
	End Sub

	Private Sub Open_val_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Open_val.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Open_val.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		'define a string containing the acceptable range of values to be used with error messages
		Dim strValidRange As String = ThisAddIn.MIN_POSITION_VAL & " and " & ThisAddIn.MAX_POSITION_VAL

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal <= ThisAddIn.MAX_POSITION_VAL And intVal >= ThisAddIn.MIN_POSITION_VAL Then
				Globals.ThisAddIn.setOpenVal(intVal)
			Else
				MsgBox("Please enter a value between " & strValidRange)
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between " & strValidRange)
		End If
	End Sub

	Private Sub Closed_val_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Closed_val.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Closed_val.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		'define a string containing the acceptable range of values to be used with error messages
		Dim strValidRange As String = ThisAddIn.MIN_POSITION_VAL & " and " & ThisAddIn.MAX_POSITION_VAL

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal <= ThisAddIn.MAX_POSITION_VAL And intVal >= ThisAddIn.MIN_POSITION_VAL Then
				Globals.ThisAddIn.setOpenVal(intVal)
			Else
				MsgBox("Please enter a value between " & strValidRange)
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between " & strValidRange)
		End If
	End Sub


	Private Sub Douser_Sub_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Douser_Sub.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Douser_Sub.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal <= ThisAddIn.MAX_SUB And intVal >= 1 Then
				Globals.ThisAddIn.setOpenVal(intVal)
			Else
				MsgBox("Please enter a valid submaster")
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between 1 and 999")
		End If
	End Sub

	Private Sub Douser_Channel_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Douser_Channel.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Douser_Channel.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal >= 1 Then
				Globals.ThisAddIn.setDouserChannel(intVal)
			Else
				MsgBox("Please enter a valid Channel number")
			End If
		Else
			MsgBox("Non Numeric values are not accepted")
		End If
	End Sub

	Private Sub OpenTime_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles OpenTime.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("[0-9]{1,}")
		Dim strVal As String = OpenTime.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal >= 0 Then
				Globals.ThisAddIn.setOpenTime(intVal)
			Else
				MsgBox("please enter a valid time")
			End If
		Else
			MsgBox("Non Numeric values are not accepted.")
		End If
	End Sub

	Private Sub CloseTime_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles CloseTime.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = CloseTime.Text ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal >= 0 Then
				Globals.ThisAddIn.setCloseTime(intVal)
			Else
				MsgBox("please enter a valid time")
			End If
		Else
			MsgBox("Non Numeric values are not accepted.")
		End If
	End Sub


End Class


