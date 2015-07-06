Public Class ThisAddIn


	'https://msdn.microsoft.com/en-us/library/vstudio/bb608612(v=vs.100).aspx (storing values in file)
	'https://msdn.microsoft.com/en-us/library/bb960904.aspx creating powerpoint add-on


	'define the private variables
	Private ipAddress As String
	Private port As String
	Private user As String
	Private openVal As Integer
	Private closedVal As String
	Private douserChannel As String
	Private douserSub As String
	Private openTime As String
	Private closeTime As String


	Public Const MAX_SUB As Integer = 999
	Public Const MAX_POSITION_VAL As Integer = 100
	Public Const MIN_POSITION_VAL As Integer = 0
	Public Const MAX_PORT_NO As Integer = UInt16.MaxValue

	Public mute As Boolean

	'setter functions


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

		ipAddress = Globals.Ribbons.Ribbon1.IP_Address.Text
		port = Globals.Ribbons.Ribbon1.Port.Text
		openVal = Globals.Ribbons.Ribbon1.Open_val.Text
		closedVal = Globals.Ribbons.Ribbon1.Closed_val.Text
		douserChannel = Globals.Ribbons.Ribbon1.Douser_Channel.Text
		douserSub = Globals.Ribbons.Ribbon1.Douser_Sub.Text
		openTime = Globals.Ribbons.Ribbon1.OpenTime.Text
		closeTime = Globals.Ribbons.Ribbon1.CloseTime.Text


		If inst <> Nothing Then
			If inst = "open" Then
				cmd = "$Sub " & douserSub & " @ " & openVal & " sneak " & openTime & "#"
			ElseIf inst = "close" Then
				cmd = "$Sub " & douserSub & " @ " & closedVal & " sneak " & closeTime & "#"
			Else
				cmd = ""
			End If
			sendCommand(cmd)
		End If
	End Sub

	Private Sub Application_AfterNewPresentation(Pres As PowerPoint.Presentation) Handles Application.AfterNewPresentation
	End Sub

	Private Sub Application_NewPresentation(Pres As PowerPoint.Presentation) Handles Application.NewPresentation

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
