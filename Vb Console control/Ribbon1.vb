Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

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
		' verifies that the input meets the valid form of an ip address and then sets the addIn ip address variable via its setter.
		If System.Net.IPAddress.TryParse(IP_Address.ToString, Nothing) Then
			Globals.ThisAddIn.setIp(IP_Address.ToString)
		Else
			MsgBox("Please enter a valid IP address")
		End If
	End Sub

	Private Sub Port_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Port.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("^\d")
		Dim strVal As String = Port.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer
		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal < ThisAddIn.MAX_PORT_NO & intVal > 0 Then
				Globals.ThisAddIn.setPort(intVal)
			Else
				MsgBox("Please enter a valid port#")
			End If
		Else
			MsgBox("Please enter a valid port#")
		End If
	End Sub

	Private Sub Open_val_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Open_val.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Open_val.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal < ThisAddIn.MAX_POSITION_VAL & intVal > ThisAddIn.MIN_POSITION_VAL Then
				Globals.ThisAddIn.setOpenVal(intVal)
			Else
				MsgBox("Please enter a value between 0 and 100")
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between 0 and 100")
		End If
	End Sub

	Private Sub Closed_val_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Closed_val.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Closed_val.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal < ThisAddIn.MAX_POSITION_VAL & intVal > ThisAddIn.MIN_POSITION_VAL Then
				Globals.ThisAddIn.setOpenVal(intVal)
			Else
				MsgBox("Please enter a value between 1 and 999")
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between 0 and 100")
		End If
	End Sub


	Private Sub Douser_Sub_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Douser_Sub.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Douser_Sub.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal < ThisAddIn.MAX_SUB & intVal > 1 Then
				Globals.ThisAddIn.setOpenVal(intVal)
			Else
				MsgBox("Please enter a valid Channel")
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between 1 and 999")
		End If
	End Sub

	Private Sub Douser_Address_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Douser_Address.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = Douser_Address.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal > 0 Then
				Globals.ThisAddIn.setDouserAddress(intVal)
			Else
				MsgBox("Please enter a value between 0 and 100")
			End If
		Else
			MsgBox("Non Numeric values are not accepted")
		End If
	End Sub

	Private Sub OpenTime_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles OpenTime.TextChanged
		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")
		Dim strVal As String = OpenTime.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal > 0 Then
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
		Dim strVal As String = CloseTime.ToString ' store the edit box value in a variable for easier changes
		Dim intVal As Integer

		If regEx.IsMatch(strVal) Then
			intVal = CInt(strVal)
			If intVal > 0 Then
				Globals.ThisAddIn.setCloseTime(intVal)
			Else
				MsgBox("please enter a valid time")
			End If
		Else
			MsgBox("Non Numeric values are not accepted.")
		End If
	End Sub
End Class


