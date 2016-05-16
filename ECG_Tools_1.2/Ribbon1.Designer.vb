Partial Class Ribbon1
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
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl6 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl7 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl8 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl9 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl10 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl11 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Tables = Me.Factory.CreateRibbonGroup
        Me.btnTBMargin = Me.Factory.CreateRibbonButton
        Me.btnLRMargin = Me.Factory.CreateRibbonButton
        Me.btnAutoRowHeight = Me.Factory.CreateRibbonButton
        Me.ebTBMargin = Me.Factory.CreateRibbonEditBox
        Me.ebLRMargin = Me.Factory.CreateRibbonEditBox
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ButtonGroup2 = Me.Factory.CreateRibbonButtonGroup
        Me.btnApplyTags = Me.Factory.CreateRibbonButton
        Me.btnRefreshTags = Me.Factory.CreateRibbonButton
        Me.ebTagName = Me.Factory.CreateRibbonEditBox
        Me.ebTagValue = Me.Factory.CreateRibbonEditBox
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
        Me.btnApplyNames = Me.Factory.CreateRibbonButton
        Me.btnRefreshNames = Me.Factory.CreateRibbonButton
        Me.ebShpName = Me.Factory.CreateRibbonEditBox
        Me.ebSldName = Me.Factory.CreateRibbonEditBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnSetMaster = Me.Factory.CreateRibbonButton
        Me.btnSetTop = Me.Factory.CreateRibbonButton
        Me.btnSetLeft = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.ddWingdings = Me.Factory.CreateRibbonDropDown
        Me.ddElements = Me.Factory.CreateRibbonDropDown
        Me.ButtonGroup3 = Me.Factory.CreateRibbonButtonGroup
        Me.btnInsertSource = Me.Factory.CreateRibbonButton
        Me.btnInsertPBHeader = Me.Factory.CreateRibbonButton
        Me.btn16ptTextBox = Me.Factory.CreateRibbonButton
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.ebSuggestions = Me.Factory.CreateRibbonEditBox
        Me.gText = Me.Factory.CreateRibbonGroup
        Me.btnHorizontalAnchor = Me.Factory.CreateRibbonButton
        Me.btnZeroMargin = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Tables.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.ButtonGroup2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.ButtonGroup1.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.ButtonGroup3.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.gText.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Tables)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.gText)
        Me.Tab1.KeyTip = "E"
        Me.Tab1.Label = "ECG_Tools1.1"
        Me.Tab1.Name = "Tab1"
        '
        'Tables
        '
        Me.Tables.Items.Add(Me.btnTBMargin)
        Me.Tables.Items.Add(Me.btnLRMargin)
        Me.Tables.Items.Add(Me.btnAutoRowHeight)
        Me.Tables.Items.Add(Me.ebTBMargin)
        Me.Tables.Items.Add(Me.ebLRMargin)
        Me.Tables.Label = "Tables"
        Me.Tables.Name = "Tables"
        '
        'btnTBMargin
        '
        Me.btnTBMargin.Label = "TB Margin"
        Me.btnTBMargin.Name = "btnTBMargin"
        Me.btnTBMargin.ScreenTip = "Set Top and Bottom Margins"
        '
        'btnLRMargin
        '
        Me.btnLRMargin.Label = "LR Margin"
        Me.btnLRMargin.Name = "btnLRMargin"
        Me.btnLRMargin.ScreenTip = "Set Left and Right Margins"
        '
        'btnAutoRowHeight
        '
        Me.btnAutoRowHeight.Label = "AutoRwHeight"
        Me.btnAutoRowHeight.Name = "btnAutoRowHeight"
        '
        'ebTBMargin
        '
        Me.ebTBMargin.Label = "ebTBMargin"
        Me.ebTBMargin.Name = "ebTBMargin"
        Me.ebTBMargin.ScreenTip = "Enter margin in hundreths of cm"
        Me.ebTBMargin.ShowLabel = False
        Me.ebTBMargin.SizeString = "000"
        Me.ebTBMargin.Text = Nothing
        '
        'ebLRMargin
        '
        Me.ebLRMargin.Label = "ebLRMargin"
        Me.ebLRMargin.Name = "ebLRMargin"
        Me.ebLRMargin.ScreenTip = "Enter margin in hundreths of cm"
        Me.ebLRMargin.ShowLabel = False
        Me.ebLRMargin.SizeString = "000"
        Me.ebLRMargin.Text = Nothing
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ButtonGroup2)
        Me.Group2.Items.Add(Me.ebTagName)
        Me.Group2.Items.Add(Me.ebTagValue)
        Me.Group2.Label = "Tags"
        Me.Group2.Name = "Group2"
        '
        'ButtonGroup2
        '
        Me.ButtonGroup2.Items.Add(Me.btnApplyTags)
        Me.ButtonGroup2.Items.Add(Me.btnRefreshTags)
        Me.ButtonGroup2.Name = "ButtonGroup2"
        '
        'btnApplyTags
        '
        Me.btnApplyTags.Label = "Apply Tags"
        Me.btnApplyTags.Name = "btnApplyTags"
        '
        'btnRefreshTags
        '
        Me.btnRefreshTags.Label = "Refresh"
        Me.btnRefreshTags.Name = "btnRefreshTags"
        '
        'ebTagName
        '
        Me.ebTagName.Label = "Tag Name"
        Me.ebTagName.Name = "ebTagName"
        Me.ebTagName.SizeString = "00000"
        Me.ebTagName.Text = Nothing
        '
        'ebTagValue
        '
        Me.ebTagValue.Label = "Tag Value"
        Me.ebTagValue.Name = "ebTagValue"
        Me.ebTagValue.SizeString = "00000"
        Me.ebTagValue.Text = Nothing
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ButtonGroup1)
        Me.Group3.Items.Add(Me.ebShpName)
        Me.Group3.Items.Add(Me.ebSldName)
        Me.Group3.Label = "Names"
        Me.Group3.Name = "Group3"
        '
        'ButtonGroup1
        '
        Me.ButtonGroup1.Items.Add(Me.btnApplyNames)
        Me.ButtonGroup1.Items.Add(Me.btnRefreshNames)
        Me.ButtonGroup1.Name = "ButtonGroup1"
        '
        'btnApplyNames
        '
        Me.btnApplyNames.Label = "Apply Names"
        Me.btnApplyNames.Name = "btnApplyNames"
        '
        'btnRefreshNames
        '
        Me.btnRefreshNames.Label = "Refresh"
        Me.btnRefreshNames.Name = "btnRefreshNames"
        '
        'ebShpName
        '
        Me.ebShpName.Label = "Shape Name"
        Me.ebShpName.Name = "ebShpName"
        Me.ebShpName.SizeString = "00000"
        Me.ebShpName.Text = Nothing
        '
        'ebSldName
        '
        Me.ebSldName.Label = "Slide Name"
        Me.ebSldName.Name = "ebSldName"
        Me.ebSldName.SizeString = "00000"
        Me.ebSldName.Text = Nothing
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnSetMaster)
        Me.Group4.Items.Add(Me.btnSetTop)
        Me.Group4.Items.Add(Me.btnSetLeft)
        Me.Group4.Label = "Shapes"
        Me.Group4.Name = "Group4"
        '
        'btnSetMaster
        '
        Me.btnSetMaster.Label = "Set Master"
        Me.btnSetMaster.Name = "btnSetMaster"
        '
        'btnSetTop
        '
        Me.btnSetTop.Label = "Set Top"
        Me.btnSetTop.Name = "btnSetTop"
        '
        'btnSetLeft
        '
        Me.btnSetLeft.Label = "Set Left"
        Me.btnSetLeft.Name = "btnSetLeft"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.ddWingdings)
        Me.Group5.Items.Add(Me.ddElements)
        Me.Group5.Items.Add(Me.ButtonGroup3)
        Me.Group5.Label = "Insert"
        Me.Group5.Name = "Group5"
        '
        'ddWingdings
        '
        RibbonDropDownItemImpl1.Label = "1"
        RibbonDropDownItemImpl2.Label = "2"
        RibbonDropDownItemImpl3.Label = "3"
        RibbonDropDownItemImpl4.Label = "4"
        RibbonDropDownItemImpl5.Label = "5"
        RibbonDropDownItemImpl6.Label = "6"
        RibbonDropDownItemImpl7.Label = "7"
        RibbonDropDownItemImpl8.Label = "8"
        RibbonDropDownItemImpl9.Label = "9"
        RibbonDropDownItemImpl10.Label = "10"
        RibbonDropDownItemImpl11.Label = "All"
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl1)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl2)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl3)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl4)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl5)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl6)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl7)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl8)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl9)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl10)
        Me.ddWingdings.Items.Add(RibbonDropDownItemImpl11)
        Me.ddWingdings.Label = "Numbers"
        Me.ddWingdings.Name = "ddWingdings"
        Me.ddWingdings.SizeString = "00"
        '
        'ddElements
        '
        Me.ddElements.Label = "Elements"
        Me.ddElements.Name = "ddElements"
        Me.ddElements.SizeString = "0000"
        '
        'ButtonGroup3
        '
        Me.ButtonGroup3.Items.Add(Me.btnInsertSource)
        Me.ButtonGroup3.Items.Add(Me.btnInsertPBHeader)
        Me.ButtonGroup3.Items.Add(Me.btn16ptTextBox)
        Me.ButtonGroup3.Name = "ButtonGroup3"
        '
        'btnInsertSource
        '
        Me.btnInsertSource.Label = "SC"
        Me.btnInsertSource.Name = "btnInsertSource"
        '
        'btnInsertPBHeader
        '
        Me.btnInsertPBHeader.Label = "PB"
        Me.btnInsertPBHeader.Name = "btnInsertPBHeader"
        '
        'btn16ptTextBox
        '
        Me.btn16ptTextBox.Label = "16"
        Me.btn16ptTextBox.Name = "btn16ptTextBox"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.ebSuggestions)
        Me.Group6.Label = "Other"
        Me.Group6.Name = "Group6"
        '
        'ebSuggestions
        '
        Me.ebSuggestions.Label = "Suggestions"
        Me.ebSuggestions.Name = "ebSuggestions"
        Me.ebSuggestions.SizeString = """00000000000"""
        Me.ebSuggestions.Text = Nothing
        '
        'gText
        '
        Me.gText.Items.Add(Me.btnHorizontalAnchor)
        Me.gText.Items.Add(Me.btnZeroMargin)
        Me.gText.Label = "Text"
        Me.gText.Name = "gText"
        '
        'btnHorizontalAnchor
        '
        Me.btnHorizontalAnchor.Label = "H.Anchor"
        Me.btnHorizontalAnchor.Name = "btnHorizontalAnchor"
        '
        'btnZeroMargin
        '
        Me.btnZeroMargin.Label = "0Marg"
        Me.btnZeroMargin.Name = "btnZeroMargin"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Tables.ResumeLayout(False)
        Me.Tables.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ButtonGroup2.ResumeLayout(False)
        Me.ButtonGroup2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.ButtonGroup1.ResumeLayout(False)
        Me.ButtonGroup1.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ButtonGroup3.ResumeLayout(False)
        Me.ButtonGroup3.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.gText.ResumeLayout(False)
        Me.gText.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ebSuggestions As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonGroup2 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents btnApplyTags As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRefreshTags As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ebTagName As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ebTagValue As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents btnApplyNames As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnRefreshNames As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ebShpName As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ebSldName As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Tables As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnTBMargin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLRMargin As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAutoRowHeight As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ebTBMargin As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ebLRMargin As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents ButtonGroup3 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents btn16ptTextBox As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnInsertSource As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnInsertPBHeader As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddElements As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents btnSetMaster As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSetTop As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSetLeft As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddWingdings As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents gText As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnHorizontalAnchor As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnZeroMargin As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
