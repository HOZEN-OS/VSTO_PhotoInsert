Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms クラス作成デザイナーのサポートに必要です
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'この呼び出しは、コンポーネント デザイナーで必要です。
        InitializeComponent()

    End Sub

    'Component は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
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

    'コンポーネント デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
    'コンポーネント デザイナーを使って変更できます。
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ChkHCenter = Me.Factory.CreateRibbonCheckBox
        Me.ChkVCenter = Me.Factory.CreateRibbonCheckBox
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.ChkRSel = Me.Factory.CreateRibbonCheckBox
        Me.ChkR169 = Me.Factory.CreateRibbonCheckBox
        Me.ChkR43 = Me.Factory.CreateRibbonCheckBox
        Me.ChkR11 = Me.Factory.CreateRibbonCheckBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.ChkWMatch = Me.Factory.CreateRibbonCheckBox
        Me.ChkHMatch = Me.Factory.CreateRibbonCheckBox
        Me.BtnAddPhoto = Me.Factory.CreateRibbonButton
        Me.ChkComp = Me.Factory.CreateRibbonCheckBox
        Me.BtnComp = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "写真取込"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.BtnAddPhoto)
        Me.Group1.Label = "写真追加"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ChkHCenter)
        Me.Group2.Items.Add(Me.ChkVCenter)
        Me.Group2.Label = "位置"
        Me.Group2.Name = "Group2"
        '
        'ChkHCenter
        '
        Me.ChkHCenter.Label = "横センター"
        Me.ChkHCenter.Name = "ChkHCenter"
        '
        'ChkVCenter
        '
        Me.ChkVCenter.Label = "縦センター"
        Me.ChkVCenter.Name = "ChkVCenter"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.ChkRSel)
        Me.Group3.Items.Add(Me.ChkR169)
        Me.Group3.Items.Add(Me.ChkR43)
        Me.Group3.Items.Add(Me.ChkR11)
        Me.Group3.Label = "トリミング"
        Me.Group3.Name = "Group3"
        '
        'ChkRSel
        '
        Me.ChkRSel.Label = "選択セル"
        Me.ChkRSel.Name = "ChkRSel"
        '
        'ChkR169
        '
        Me.ChkR169.Label = "１６：９"
        Me.ChkR169.Name = "ChkR169"
        '
        'ChkR43
        '
        Me.ChkR43.Label = "４：３"
        Me.ChkR43.Name = "ChkR43"
        '
        'ChkR11
        '
        Me.ChkR11.Label = "１：１"
        Me.ChkR11.Name = "ChkR11"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.ChkWMatch)
        Me.Group4.Items.Add(Me.ChkHMatch)
        Me.Group4.Items.Add(Me.ChkComp)
        Me.Group4.Label = "サイズ"
        Me.Group4.Name = "Group4"
        '
        'ChkWMatch
        '
        Me.ChkWMatch.Label = "横幅で合わせる"
        Me.ChkWMatch.Name = "ChkWMatch"
        '
        'ChkHMatch
        '
        Me.ChkHMatch.Label = "高さで合わせる"
        Me.ChkHMatch.Name = "ChkHMatch"
        '
        'BtnAddPhoto
        '
        Me.BtnAddPhoto.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnAddPhoto.Label = "写真選択"
        Me.BtnAddPhoto.Name = "BtnAddPhoto"
        Me.BtnAddPhoto.OfficeImageId = "PictureReflectionGalleryItem"
        Me.BtnAddPhoto.ShowImage = True
        '
        'ChkComp
        '
        Me.ChkComp.Label = "圧縮する"
        Me.ChkComp.Name = "ChkComp"
        '
        'BtnComp
        '
        Me.BtnComp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnComp.Label = "全て圧縮"
        Me.BtnComp.Name = "BtnComp"
        Me.BtnComp.OfficeImageId = "PictureSwapPicture"
        Me.BtnComp.ShowImage = True
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.BtnComp)
        Me.Group5.Label = "その他"
        Me.Group5.Name = "Group5"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnAddPhoto As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ChkHCenter As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ChkVCenter As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ChkR43 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ChkR169 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ChkR11 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ChkWMatch As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ChkHMatch As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ChkRSel As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ChkComp As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents BtnComp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
