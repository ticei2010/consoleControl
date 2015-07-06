Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.DouserCtrlTab = Me.Factory.CreateRibbonTab
		Me.ShowSettings = Me.Factory.CreateRibbonGroup
		Me.DouserCtrlEnable = Me.Factory.CreateRibbonToggleButton
		Me.mute = Me.Factory.CreateRibbonToggleButton
		Me.Separator2 = Me.Factory.CreateRibbonSeparator
		Me.Label1 = Me.Factory.CreateRibbonLabel
		Me.OpenTime = Me.Factory.CreateRibbonEditBox
		Me.CloseTime = Me.Factory.CreateRibbonEditBox
		Me.Douser_Info = Me.Factory.CreateRibbonGroup
		Me.Douser_Channel = Me.Factory.CreateRibbonEditBox
		Me.Douser_Sub = Me.Factory.CreateRibbonEditBox
		Me.Separator1 = Me.Factory.CreateRibbonSeparator
		Me.Open_val = Me.Factory.CreateRibbonEditBox
		Me.Closed_val = Me.Factory.CreateRibbonEditBox
		Me.Console_Settings = Me.Factory.CreateRibbonGroup
		Me.IP_Address = Me.Factory.CreateRibbonEditBox
		Me.Port = Me.Factory.CreateRibbonEditBox
		Me.User = Me.Factory.CreateRibbonEditBox
		Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
		Me.Channel = Me.Factory.CreateRibbonToggleButton
		Me.submaster = Me.Factory.CreateRibbonToggleButton
		Me.DouserCtrlTab.SuspendLayout()
		Me.ShowSettings.SuspendLayout()
		Me.Douser_Info.SuspendLayout()
		Me.Console_Settings.SuspendLayout()
		Me.ButtonGroup1.SuspendLayout()
		'
		'DouserCtrlTab
		'
		Me.DouserCtrlTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
		Me.DouserCtrlTab.Groups.Add(Me.ShowSettings)
		Me.DouserCtrlTab.Groups.Add(Me.Douser_Info)
		Me.DouserCtrlTab.Groups.Add(Me.Console_Settings)
		Me.DouserCtrlTab.Label = "Douser"
		Me.DouserCtrlTab.Name = "DouserCtrlTab"
		'
		'ShowSettings
		'
		Me.ShowSettings.Items.Add(Me.DouserCtrlEnable)
		Me.ShowSettings.Items.Add(Me.mute)
		Me.ShowSettings.Items.Add(Me.Separator2)
		Me.ShowSettings.Items.Add(Me.Label1)
		Me.ShowSettings.Items.Add(Me.OpenTime)
		Me.ShowSettings.Items.Add(Me.CloseTime)
		Me.ShowSettings.Label = "Show Settings"
		Me.ShowSettings.Name = "ShowSettings"
		'
		'DouserCtrlEnable
		'
		Me.DouserCtrlEnable.Label = "Enable Douser Control"
		Me.DouserCtrlEnable.Name = "DouserCtrlEnable"
		'
		'mute
		'
		Me.mute.Enabled = False
		Me.mute.Image = Global.Vb_Console_control.My.Resources.Resources.NO_pic
		Me.mute.Label = "Mute Douser"
		Me.mute.Name = "mute"
		Me.mute.ShowImage = True
		Me.mute.Tag = ""
		'
		'Separator2
		'
		Me.Separator2.Name = "Separator2"
		'
		'Label1
		'
		Me.Label1.Label = "Action Time"
		Me.Label1.Name = "Label1"
		'
		'OpenTime
		'
		Me.OpenTime.Enabled = False
		Me.OpenTime.Label = "Open"
		Me.OpenTime.Name = "OpenTime"
		Me.OpenTime.ScreenTip = "seconds"
		Me.OpenTime.Text = "3"
		'
		'CloseTime
		'
		Me.CloseTime.Enabled = False
		Me.CloseTime.Label = "Close"
		Me.CloseTime.Name = "CloseTime"
		Me.CloseTime.ScreenTip = "seconds"
		Me.CloseTime.Text = "5"
		'
		'Douser_Info
		'
		Me.Douser_Info.Items.Add(Me.ButtonGroup1)
		Me.Douser_Info.Items.Add(Me.Douser_Channel)
		Me.Douser_Info.Items.Add(Me.Douser_Sub)
		Me.Douser_Info.Items.Add(Me.Separator1)
		Me.Douser_Info.Items.Add(Me.Open_val)
		Me.Douser_Info.Items.Add(Me.Closed_val)
		Me.Douser_Info.Label = "Douser Settings"
		Me.Douser_Info.Name = "Douser_Info"
		'
		'Douser_Channel
		'
		Me.Douser_Channel.Enabled = False
		Me.Douser_Channel.Label = "Douser Channel   "
		Me.Douser_Channel.Name = "Douser_Channel"
		Me.Douser_Channel.Text = "150"
		'
		'Douser_Sub
		'
		Me.Douser_Sub.Enabled = False
		Me.Douser_Sub.Label = "Douser Submaster"
		Me.Douser_Sub.Name = "Douser_Sub"
		Me.Douser_Sub.SizeString = "100"
		Me.Douser_Sub.Text = "100"
		'
		'Separator1
		'
		Me.Separator1.Name = "Separator1"
		'
		'Open_val
		'
		Me.Open_val.Enabled = False
		Me.Open_val.Label = "Open Value  "
		Me.Open_val.Name = "Open_val"
		Me.Open_val.SizeString = "100"
		Me.Open_val.SuperTip = "The intensity at which the douser is open."
		Me.Open_val.Text = "0"
		'
		'Closed_val
		'
		Me.Closed_val.Enabled = False
		Me.Closed_val.Label = "Closed Value"
		Me.Closed_val.Name = "Closed_val"
		Me.Closed_val.SizeString = "100"
		Me.Closed_val.SuperTip = "The intensity at which the douser is closed."
		Me.Closed_val.Text = "85"
		'
		'Console_Settings
		'
		Me.Console_Settings.Items.Add(Me.IP_Address)
		Me.Console_Settings.Items.Add(Me.Port)
		Me.Console_Settings.Items.Add(Me.User)
		Me.Console_Settings.Label = "Console Settings"
		Me.Console_Settings.Name = "Console_Settings"
		'
		'IP_Address
		'
		Me.IP_Address.Enabled = False
		Me.IP_Address.Label = "Ip Address"
		Me.IP_Address.MaxLength = 15
		Me.IP_Address.Name = "IP_Address"
		Me.IP_Address.SizeString = "000.000.000.000"
		Me.IP_Address.SuperTip = "Your console's IP Address"
		Me.IP_Address.Text = "192.168.1.84"
		'
		'Port
		'
		Me.Port.Enabled = False
		Me.Port.Label = "Port          "
		Me.Port.Name = "Port"
		Me.Port.SizeString = "5000"
		Me.Port.SuperTip = "The recieve port set in your console's show control settings."
		Me.Port.Text = "5000"
		'
		'User
		'
		Me.User.Enabled = False
		Me.User.Label = "User #      "
		Me.User.Name = "User"
		Me.User.ScreenTip = "0-9"
		Me.User.SizeString = "0"
		Me.User.Text = "1"
		'
		'ButtonGroup1
		'
		Me.ButtonGroup1.Items.Add(Me.Channel)
		Me.ButtonGroup1.Items.Add(Me.submaster)
		Me.ButtonGroup1.Name = "ButtonGroup1"
		'
		'Channel
		'
		Me.Channel.Checked = True
		Me.Channel.Enabled = False
		Me.Channel.Label = " Channel "
		Me.Channel.Name = "Channel"
		'
		'submaster
		'
		Me.submaster.Enabled = False
		Me.submaster.Label = " Sub "
		Me.submaster.Name = "submaster"
		'
		'Ribbon1
		'
		Me.Name = "Ribbon1"
		Me.RibbonType = "Microsoft.PowerPoint.Presentation"
		Me.Tabs.Add(Me.DouserCtrlTab)
		Me.DouserCtrlTab.ResumeLayout(False)
		Me.DouserCtrlTab.PerformLayout()
		Me.ShowSettings.ResumeLayout(False)
		Me.ShowSettings.PerformLayout()
		Me.Douser_Info.ResumeLayout(False)
		Me.Douser_Info.PerformLayout()
		Me.Console_Settings.ResumeLayout(False)
		Me.Console_Settings.PerformLayout()
		Me.ButtonGroup1.ResumeLayout(False)
		Me.ButtonGroup1.PerformLayout()

	End Sub

	Friend WithEvents DouserCtrlTab As Microsoft.Office.Tools.Ribbon.RibbonTab
	Friend WithEvents Console_Settings As Microsoft.Office.Tools.Ribbon.RibbonGroup
	Friend WithEvents IP_Address As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents Port As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents Douser_Info As Microsoft.Office.Tools.Ribbon.RibbonGroup
	Friend WithEvents Douser_Channel As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents Open_val As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents Closed_val As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents User As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents Douser_Sub As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
	Friend WithEvents ShowSettings As Microsoft.Office.Tools.Ribbon.RibbonGroup
	Friend WithEvents DouserCtrlEnable As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
	Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
	Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
	Friend WithEvents OpenTime As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents CloseTime As Microsoft.Office.Tools.Ribbon.RibbonEditBox
	Friend WithEvents mute As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
	Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
	Friend WithEvents Channel As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
	Friend WithEvents submaster As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
