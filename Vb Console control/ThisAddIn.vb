Public Class ThisAddIn
	'define the private variables
	Private ipAddress As String
	Private port As Integer
	Private user As Integer
	Private openVal As Integer
	Private closedVal As Integer
	Private douserChannel As String
	Private douserSub As Integer
	Private openTime As Integer
	Private closeTime As Integer


	Public Const MAX_SUB As Integer = 999
	Public Const MAX_POSITION_VAL As Integer = 100
	Public Const MIN_POSITION_VAL As Integer = 0
	Public Const MAX_PORT_NO As Integer = UInt16.MaxValue

	Public mute As Boolean

	'setter functions
	Public Sub setIp(consoleIpAddress As String)
		ipAddress = consoleIpAddress
	End Sub
	Public Sub setPort(consolePort As Integer)
		port = consolePort
	End Sub
	Public Sub setUser(userNo As Integer)
		user = userNo
	End Sub
	Public Sub setOpenVal(value As Integer)
		openVal = value
	End Sub
	Public Sub setClosedVal(value As Integer)
		closedVal = value
	End Sub
	Public Sub setDouserChannel(channel As String)
		douserChannel = channel
	End Sub
	Public Sub setDouserSub(submaster As Integer)
		douserSub = submaster
	End Sub
	Public Sub setOpenTime(time As Integer)
		openTime = time
	End Sub
	Public Sub setCloseTime(time As Integer)
		closeTime = time
	End Sub


	Private Sub ThisAddIn_Startup() Handles Me.Startup

	End Sub

	Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

	End Sub



	''' <summary>
	''' Sends the input string as a Byte array in a UDP packet
	''' </summary>
	''' <param name="command">String to be sent in the UDP packet </param>
	Private Sub sendCommand(command As String)
		Dim udpClient As New System.Net.Sockets.UdpClient(ipAddress, port)

		Dim dgram As Byte() = System.Text.Encoding.ASCII.GetBytes(command)

		udpClient.Send(dgram, dgram.Length)


	End Sub
	Private Sub parseCmd(ByVal wn As PowerPoint.SlideShowWindow)
		Dim activeSlide As Integer = wn.View.CurrentShowPosition
		Dim sld As PowerPoint.Slide = wn.View.Slide
		Dim notes As String = sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text

		Dim regEx As RegularExpressions.Regex = New RegularExpressions.Regex("<(open|close>")
		Dim inst As String = regEx.Match(notes).Groups(1).ToString
		Dim cmd As String

		If inst <> Nothing Then
			If inst = "open" Then
				cmd = "$Sub " & douserSub & openVal & " sneak " & openTime & "#"
			ElseIf inst = "close" Then
				cmd = "$Sub " & douserSub & closedVal & " sneak " & closeTime & "#"
			Else
				MsgBox("invalid test you dummy")
				cmd = ""
			End If
			sendCommand(cmd)
		End If
	End Sub

	Private Sub Application_SlideShowNextSlide(Wn As PowerPoint.SlideShowWindow) Handles Application.SlideShowNextSlide
		parseCmd(Wn)
	End Sub
End Class
