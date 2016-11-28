Partial Class ReplyRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabAutoReply = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnStandardReply = Me.Factory.CreateRibbonButton
        Me.btnNomReply = Me.Factory.CreateRibbonButton
        Me.btnGenericResponse = Me.Factory.CreateRibbonButton
        Me.btnNEMReply = Me.Factory.CreateRibbonButton
        Me.TabAutoReply.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabAutoReply
        '
        Me.TabAutoReply.Groups.Add(Me.Group1)
        Me.TabAutoReply.Groups.Add(Me.Group2)
        Me.TabAutoReply.Label = "Auto-reply"
        Me.TabAutoReply.Name = "TabAutoReply"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnStandardReply)
        Me.Group1.Items.Add(Me.btnNomReply)
        Me.Group1.Items.Add(Me.btnGenericResponse)
        Me.Group1.Label = "Honours"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnNEMReply)
        Me.Group2.Label = "NEM"
        Me.Group2.Name = "Group2"
        '
        'btnStandardReply
        '
        Me.btnStandardReply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnStandardReply.Label = "Standard"
        Me.btnStandardReply.Name = "btnStandardReply"
        Me.btnStandardReply.OfficeImageId = "DirectRepliesTo"
        Me.btnStandardReply.ShowImage = True
        '
        'btnNomReply
        '
        Me.btnNomReply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnNomReply.Label = "Nomination"
        Me.btnNomReply.Name = "btnNomReply"
        Me.btnNomReply.OfficeImageId = "DirectRepliesTo"
        Me.btnNomReply.ShowImage = True
        '
        'btnGenericResponse
        '
        Me.btnGenericResponse.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnGenericResponse.Label = "Generic"
        Me.btnGenericResponse.Name = "btnGenericResponse"
        Me.btnGenericResponse.OfficeImageId = "DirectRepliesTo"
        Me.btnGenericResponse.ShowImage = True
        '
        'btnNEMReply
        '
        Me.btnNEMReply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnNEMReply.Label = "NEM Nomination"
        Me.btnNEMReply.Name = "btnNEMReply"
        Me.btnNEMReply.OfficeImageId = "DirectRepliesTo"
        Me.btnNEMReply.ShowImage = True
        '
        'ReplyRibbon
        '
        Me.Name = "ReplyRibbon"
        Me.RibbonType = "Microsoft.Outlook.Mail.Read,Microsoft.Outlook.Mail.Compose"
        Me.Tabs.Add(Me.TabAutoReply)
        Me.TabAutoReply.ResumeLayout(False)
        Me.TabAutoReply.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabAutoReply As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnStandardReply As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnNomReply As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnNEMReply As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnGenericResponse As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ReplyRibbon() As ReplyRibbon
        Get
            Return Me.GetRibbon(Of ReplyRibbon)()
        End Get
    End Property
End Class
