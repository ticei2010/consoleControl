Public Class ThisAddIn


	'https://msdn.microsoft.com/en-us/library/vstudio/bb608612(v=vs.100).aspx (storing values in file)
	'https://msdn.microsoft.com/en-us/library/bb960904.aspx creating powerpoint add-on


	'define the private variables
	Private ipAddress As String
	Private port As String


	Public Const MAX_SUB As Integer = 999
	Public Const MAX_POSITION_VAL As Integer = 100
	Public Const MIN_POSITION_VAL As Integer = 0
	Public Const MAX_PORT_NO As Integer = UInt16.MaxValue

	Public mute As Boolean

	Private Sub ThisAddIn_Startup() Handles Me.Startup

	End Sub

	Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

	End Sub



	''' <summary>
	''' Sends the input string as a Byte array in a UDP packet
	''' </summary>
	''' <param name="command">String to be sent in the UDP packet </param>
	Private Sub sendCommand(command As String)
		Dim udpClient As New System.Net.Sockets.UdpClient(ipAddress, CInt(port))

		Dim dgram As Byte() = System.Text.Encoding.ASCII.GetBytes(command)

		udpClient.Send(dgram, dgram.Length)


	End Sub
	Private Sub parseCmd(ByVal wn As PowerPoint.SlideShowWindow)
		Dim activeSlide As Integer = wn.View.CurrentShowPosition
		Dim sld As PowerPoint.Slide = wn.View.Slide
		Dim notes As String = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("<(open|close)>")
		Dim inst As String = regEx.Match(notes).Groups(1).ToString
		Dim cmd As String
		With Globals.Ribbons.Ribbon1
			ipAddress = .IP_Address.Text
			port = .Port.Text
			Dim user As String = .User.Text
			Dim openVal As Integer = .Open_val.Text
			Dim closedVal As String = .Closed_val.Text
			Dim chan_sub As String
			Dim openTime As String = .OpenTime.Text
			Dim closeTime As String = .CloseTime.Text

			If .Channel.Checked Then
				chan_sub = .Douser_Channel.Text
			Else
				chan_sub = .Douser_Sub.Text
			End If

			If inst <> Nothing Then
				If inst = "open" Then
					cmd = "$<u" & user & "> " & .getNode("douser/channel-sub").Text & " " & chan_sub & " @ " & openVal & " sneak " & openTime & "#"
				ElseIf inst = "close" Then
					cmd = "$<u" & user & "> " & .getNode("douser/channel-sub").Text & " " & chan_sub & " @ " & closedVal & " sneak " & closeTime & "#"
				Else
					cmd = ""
				End If
				sendCommand(cmd)
			End If
		End With
	End Sub

	Private Sub Application_AfterPresentationOpen(Pres As PowerPoint.Presentation) Handles Application.AfterPresentationOpen
		If Globals.Ribbons.Ribbon1.storedControls() Then
			Globals.Ribbons.Ribbon1.loadSettings()
		End If
	End Sub


	

	Private Sub Application_SlideShowBegin(Wn As PowerPoint.SlideShowWindow) Handles Application.SlideShowBegin
		If Globals.Ribbons.Ribbon1.mute.Checked Then
			MsgBox("Douser mute is enabled.")
		End If
	End Sub

	Private Sub Application_SlideShowNextSlide(Wn As PowerPoint.SlideShowWindow) Handles Application.SlideShowNextSlide
		If Not Globals.Ribbons.Ribbon1.mute.Checked And Globals.Ribbons.Ribbon1.DouserCtrlEnable.Checked Then
			parseCmd(Wn)
		End If
	End Sub
End Class
