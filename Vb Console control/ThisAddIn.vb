Public Class ThisAddIn
	'define the private variables
	Private ipAddress As String
	Private port As Integer
	Private user As Integer
	Private openVal As Integer
	Private closedVal As Integer
	Private douserAddress As String
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
	Public Sub setDouserAddress(address As String)
		douserAddress = address
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
		Dim udpClient As New System.Net.Sockets.UdpClient()

		Dim dgram As Byte() = System.Text.Encoding.ASCII.GetBytes(command)

		udpClient.Send(dgram, dgram.Length)


	End Sub
End Class
