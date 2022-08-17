Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1



    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub BtnAddPhoto_Click(sender As Object, e As RibbonControlEventArgs) Handles BtnAddPhoto.Click
        If IsSelect() Then
            PutPhotos(GetFiles())
        End If
    End Sub

    Private Sub ChkHCenter_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkHCenter.Click
        If ChkHCenter.Checked Then
            ChkVCenter.Checked = False
            ImgSetting.Center = Setting.HVType.HORIZON
        Else
            ImgSetting.Center = Setting.HVType.NONE
        End If
    End Sub

    Private Sub ChkVCenter_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkVCenter.Click
        If ChkVCenter.Checked Then
            ChkHCenter.Checked = False
            ImgSetting.Center = Setting.HVType.VERTICAL
        Else
            ImgSetting.Center = Setting.HVType.NONE
        End If
    End Sub

    Private Sub ChkR43_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkR43.Click
        If ChkR43.Checked Then
            ChkR169.Checked = False
            ChkR11.Checked = False
            ChkRSel.Checked = False
            ImgSetting.Aspect = Setting.AspectType.A43
        Else
            ImgSetting.Aspect = Setting.AspectType.NONE
        End If
    End Sub

    Private Sub ChkR169_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkR169.Click
        If ChkR169.Checked Then
            ChkR43.Checked = False
            ChkR11.Checked = False
            ChkRSel.Checked = False
            ImgSetting.Aspect = Setting.AspectType.A169
        Else
            ImgSetting.Aspect = Setting.AspectType.NONE
        End If
    End Sub

    Private Sub ChkR11_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkR11.Click
        If ChkR11.Checked Then
            ChkR43.Checked = False
            ChkR169.Checked = False
            ChkRSel.Checked = False
            ImgSetting.Aspect = Setting.AspectType.A11
        Else
            ImgSetting.Aspect = Setting.AspectType.NONE
        End If
    End Sub

    Private Sub ChkRSel_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkRSel.Click
        If ChkRSel.Checked Then
            ChkR43.Checked = False
            ChkR169.Checked = False
            ChkR11.Checked = False
            ImgSetting.Aspect = Setting.AspectType.CELL
        Else
            ImgSetting.Aspect = Setting.AspectType.NONE
        End If
    End Sub

    Private Sub ChkWMatch_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkWMatch.Click
        If ChkWMatch.Checked Then
            ChkHMatch.Checked = False
            ImgSetting.Size = Setting.HVType.HORIZON
        Else
            ImgSetting.Size = Setting.HVType.NONE
        End If
    End Sub

    Private Sub ChkHMatch_Click(sender As Object, e As RibbonControlEventArgs) Handles ChkHMatch.Click
        If ChkHMatch.Checked Then
            ChkWMatch.Checked = False
            ImgSetting.Size = Setting.HVType.VERTICAL
        Else
            ImgSetting.Size = Setting.HVType.NONE
        End If
    End Sub
End Class
