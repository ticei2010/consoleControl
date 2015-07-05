Imports Microsoft.Office.Tools.Ribbon

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

	''' <summary>
	''' Returns the xml node specified by the extension
	''' </summary>
	''' <param name="extension">A string defining the path form "/douser_controls/" to the desired node </param>
	''' <returns>Office.CustomXMLNode</returns>
	''' <remarks></remarks>
	Private Function getNode(extension As String) As Office.CustomXMLNode
		'retrieve the active presentation object
		Dim pres As PowerPoint.Presentation = Globals.ThisAddIn.Application.ActivePresentation

		'load the douser control xml and select the port node
		Dim xml As Office.CustomXMLPart = pres.CustomXMLParts.SelectByID( _
			pres.CustomDocumentProperties.Item("douser_controls").Value)
		Return xml.SelectSingleNode("/douser_controls/" & extension)
	End Function

	Private Sub IP_Address_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles IP_Address.TextChanged
		Dim nodePath As String = "console/ip"

		'if a new valid ip address is provided update the xml otherwise
		'notify the user and reset to the last good value
		If System.Net.IPAddress.TryParse(IP_Address.Text, Nothing) Then
			getNode(nodePath).Text = IP_Address.Text
		Else
			MsgBox("Please enter a valid IP address")
			IP_Address.Text = getNode(nodePath).Text
		End If
	End Sub

	Private Sub Port_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Port.TextChanged
		Dim nodePath As String = "console/port"

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("^\d")
		'if a new valid port is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(Port.Text) Then
			If CInt(Port.Text) <= ThisAddIn.MAX_PORT_NO And CInt(Port.Text) > 0 Then
				getNode(nodePath).Text = Port.Text
			Else
				MsgBox("Please enter a valid port number")
				Port.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Please enter a valid port number")
			Port.Text = getNode(nodePath).Text
		End If
	End Sub

	Private Sub Open_val_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Open_val.TextChanged
		Dim nodePath As String = "douser/open_val"

		'define a string containing the acceptable range of values to be used with error messages
		Dim strValidRange As String = ThisAddIn.MIN_POSITION_VAL & " and " & ThisAddIn.MAX_POSITION_VAL

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")

		'if a new valid open val is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(Open_val.Text) Then
			If CInt(Open_val.Text) <= ThisAddIn.MAX_POSITION_VAL And CInt(Open_val.Text) >= ThisAddIn.MIN_POSITION_VAL Then
				getNode(nodePath).Text = Open_val.Text
			Else
				MsgBox("Please enter a value between " & strValidRange)
				Open_val.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between " & strValidRange)
			Open_val.Text = getNode(nodePath).Text
		End If
	End Sub

	Private Sub Closed_val_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Closed_val.TextChanged
		Dim nodePath As String = "douser/closed_val"

		'define a string containing the acceptable range of values to be used with error messages
		Dim strValidRange As String = ThisAddIn.MIN_POSITION_VAL & " and " & ThisAddIn.MAX_POSITION_VAL

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")

		'if a new valid closed val is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(Closed_val.Text) Then
			If CInt(Closed_val.Text) <= ThisAddIn.MAX_POSITION_VAL And CInt(Closed_val.Text) >= ThisAddIn.MIN_POSITION_VAL Then
				getNode(nodePath).Text = Closed_val.Text
			Else
				MsgBox("Please enter a value between " & strValidRange)
				Closed_val.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between " & strValidRange)
			Closed_val.Text = getNode(nodePath).Text
		End If
	End Sub


	Private Sub Douser_Sub_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Douser_Sub.TextChanged
		Dim nodePath As String = "douser/submaster"

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")

		'if a new valid submaster is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(Douser_Sub.Text) Then
			If CInt(Douser_Sub.Text) <= ThisAddIn.MAX_SUB And CInt(Douser_Sub.Text) >= 1 Then
				getNode(nodePath).Text = Douser_Sub.Text
			Else
				MsgBox("Please enter a valid submaster")
				Douser_Sub.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Non Numeric values are not accepted. Please enter a value between 1 and 999")
			Douser_Sub.Text = getNode(nodePath).Text
		End If
	End Sub

	Private Sub Douser_Channel_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles Douser_Channel.TextChanged
		Dim nodePath As String = "douser/channel"

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")

		'if a new valid channel is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(Douser_Channel.Text) Then
			If CInt(Douser_Channel.Text) >= 1 Then
				getNode(nodePath).Text = Douser_Channel.Text
			Else
				MsgBox("Please enter a valid Channel number")
				Douser_Channel.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Non Numeric values are not accepted")
			Douser_Channel.Text = getNode(nodePath).Text
		End If
	End Sub

	Private Sub OpenTime_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles OpenTime.TextChanged
		Dim nodePath As String = "show/open_time"

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("[0-9]{1,}")

		'if a new valid open time is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(OpenTime.Text) Then
			If CInt(OpenTime.Text) >= 0 Then
				getNode(nodePath).Text = OpenTime.Text
			Else
				MsgBox("please enter a valid time")
				OpenTime.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Non Numeric values are not accepted.")
			OpenTime.Text = getNode(nodePath).Text
		End If
	End Sub

	Private Sub CloseTime_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles CloseTime.TextChanged
		Dim nodePath As String = "show/close_time"

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("\d")

		'if a new valid close time is provided update the xml otherwise
		'notify the user and reset to the last good value
		If regEx.IsMatch(CloseTime.Text) Then
			If CInt(CloseTime.Text) >= 0 Then
				getNode(nodePath).Text = CloseTime.Text
			Else
				MsgBox("please enter a valid time")
				CloseTime.Text = getNode(nodePath).Text
			End If
		Else
			MsgBox("Non Numeric values are not accepted.")
			CloseTime.Text = getNode(nodePath).Text
		End If
	End Sub


End Class


