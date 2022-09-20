Imports System.Collections
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Forms

Module Photo
    Public ImgSetting As Setting = Setting.GetSingleton()
    Private ReadOnly Application As Excel.Application = Globals.ThisAddIn.Application
    Private OpenDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
    Private PicSize As New Size

    Public Function IsSelect() As Boolean
        If Application.Selection.Columns.Count > 0 Then
            Return True
        End If
        Return False
    End Function

    Public Function GetFiles() As ArrayList
        Dim ofd As New OpenFileDialog With {
            .InitialDirectory = OpenDirectory,
            .Filter = "JPEG画像(*.jpg; *.jpeg)|*.jpg; *.jpeg",
            .Multiselect = True
        }

        Dim al As New ArrayList
        If ofd.ShowDialog() <> DialogResult.Cancel Then
            OpenDirectory = Path.GetDirectoryName(ofd.FileNames(0))
            al.AddRange(ofd.FileNames)
        End If

        Return al
    End Function

    Public Sub PutPhotos(mFileList As ArrayList)
        If mFileList.Count > 0 Then
            Application.ScreenUpdating = False
            PicSize.Width = GetWidth()
            PicSize.Height = GetHeight()
            PutPhoto(mFileList)
            Application.ScreenUpdating = True
        End If
    End Sub

    Private Sub PutPhoto(fl As ArrayList)
        Dim FileName As String
        Dim CopyFile As String
        Dim OriginalFile As FileInfo
        For Each FileName In fl
            OriginalFile = New FileInfo(FileName)
            CopyFile = OriginalFile.DirectoryName & "\copy_" & OriginalFile.Name
            OriginalFile.CopyTo(CopyFile, True)

            ChangeRotate(CopyFile)
            If ImgSetting.Aspect > Setting.AspectType.NONE Then
                ChangeAspect(CopyFile)
            End If

            With Application.ActiveSheet.Shapes.AddPicture(CopyFile, False, True, 0, 0, 0, 0)
                .LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue
                .ScaleHeight(1, Microsoft.Office.Core.MsoTriState.msoTrue)
                .ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue)

                Select Case ImgSetting.Size
                    Case Setting.HVType.NONE
                        .Width = PicSize.Width
                        If .Height > PicSize.Height Then
                            .Height = PicSize.Height
                        End If
                    Case Setting.HVType.HORIZON
                        .Width = PicSize.Width
                    Case Setting.HVType.VERTICAL
                        .Height = PicSize.Height
                End Select

                If ImgSetting.Center = Setting.HVType.HORIZON AndAlso PicSize.Width > .Width Then
                    .Left = Application.ActiveCell.Left + (PicSize.Width - .Width) / 2
                Else
                    .Left = Application.ActiveCell.Left
                End If

                If ImgSetting.Center = Setting.HVType.VERTICAL AndAlso PicSize.Height > .Height Then
                    .Top = Application.ActiveCell.Top + (PicSize.Height - .Height) / 2
                Else
                    .Top = Application.ActiveCell.Top
                End If


                If ImgSetting.Compress Then
                    Dim P As New Point(.Left, .Top)
                    .Copy
                    .Delete
                    Application.ActiveSheet.PasteSpecial(Format:="図 (JPEG)")
                    Application.Selection.Top = P.Y
                    Application.Selection.Left = P.X
                End If
            End With

            Kill(CopyFile)
        Next
    End Sub

    Private Function GetWidth() As Double
        Dim W As Double = 0
        If Application.Selection.Columns.Count = 1 Then
            W = Application.ActiveCell.Width
        Else
            For Each Col As Excel.Range In Application.Selection.Columns
                W += Col.Width
            Next
        End If

        Return W
    End Function

    Private Function GetHeight() As Double
        Dim H As Double = 0
        If Application.Selection.Rows.Count = 1 Then
            H = Application.ActiveCell.Height
        Else
            For Each Row As Excel.Range In Application.Selection.Rows
                H += Row.Height
            Next
        End If

        Return H
    End Function

    Public Sub ChangeRotate(ImageFile As String)
        Dim ReadBitmap As New Bitmap(ImageFile)
        Dim SaveBitmap As Bitmap = DirectCast(ReadBitmap.Clone(), Bitmap)
        Dim Rotation As RotateFlipType = RotateFlipType.RotateNoneFlipNone
        Dim PropItem As PropertyItem = Nothing

        For Each PropItem In ReadBitmap.PropertyItems
            If PropItem.Id = &H112 Then
                Select Case PropItem.Value(0)
                    Case 3
                        Rotation = RotateFlipType.Rotate180FlipNone
                    Case 6
                        Rotation = RotateFlipType.Rotate90FlipNone
                    Case 8
                        Rotation = RotateFlipType.Rotate270FlipNone
                End Select
                Exit For
            End If
        Next
        ReadBitmap.Dispose()

        If Rotation <> RotateFlipType.RotateNoneFlipNone Then
            PropItem.Value(0) = &H1
            PropItem.Len = PropItem.Value.Length
            SaveBitmap.RotateFlip(Rotation)
            SaveBitmap.SetPropertyItem(PropItem)

            Dim eps As New EncoderParameters(1)
            eps.Param(0) = New EncoderParameter(Imaging.Encoder.Quality, 90)

            SaveBitmap.Save(ImageFile, GetEncoderInfo("image/jpeg"), eps)
        End If
        SaveBitmap.Dispose()
    End Sub

    Private Function GetEncoderInfo(mineType As String) As ImageCodecInfo
        For Each enc As ImageCodecInfo In ImageCodecInfo.GetImageEncoders()
            If enc.MimeType = mineType Then
                Return enc
            End If
        Next
        Return Nothing
    End Function

    Private Sub ChangeAspect(ImageFile As String)
        Dim ReadBitmap As New Bitmap(ImageFile)
        Dim NewSize As Size
        Select Case ImgSetting.Aspect
            Case Setting.AspectType.A43
                NewSize = GetSize(ReadBitmap.Width, ReadBitmap.Height, 4, 3)
            Case Setting.AspectType.A169
                NewSize = GetSize(ReadBitmap.Width, ReadBitmap.Height, 16, 9)
            Case Setting.AspectType.A11
                NewSize = GetSize(ReadBitmap.Width, ReadBitmap.Height, 1, 1)
            Case Setting.AspectType.CELL
                NewSize = GetSize(ReadBitmap.Width, ReadBitmap.Height, GetWidth(), GetHeight())
        End Select
        Dim canvas As New Bitmap(NewSize.Width, NewSize.Height)
        Dim g As Graphics = Graphics.FromImage(canvas)
        Dim srcRect As New Rectangle(0, (ReadBitmap.Height - NewSize.Height) / 2, NewSize.Width, NewSize.Height)
        Dim desRect As New Rectangle(0, 0, srcRect.Width, srcRect.Height)
        g.DrawImage(ReadBitmap, desRect, srcRect, GraphicsUnit.Pixel)
        g.Dispose()
        ReadBitmap.Dispose()

        canvas.Save(ImageFile, ImageFormat.Jpeg)
        canvas.Dispose()
    End Sub

    Private Function GetSize(w As Integer, h As Integer, aw As Integer, ah As Integer) As Size
        Dim NewSize As New Size(w, h)
        Dim OldSize As New Size(w, h)
        If w < h Then
            OldSize = New Size(h, w)
        End If

        If OldSize.Width / aw * ah <= OldSize.Height Then
            NewSize.Width = OldSize.Width
            NewSize.Height = OldSize.Width * ah / aw
        Else
            NewSize.Width = OldSize.Height / ah * aw
            NewSize.Height = OldSize.Height
        End If

        If w < h Then
            Return New Size(NewSize.Height, NewSize.Width)
        End If

        Return NewSize
    End Function

    Public Sub CompImage()
        For Each es As Excel.Shape In Application.ActiveSheet.Shapes
            If es.Type = Microsoft.Office.Core.MsoShapeType.msoPicture Then
                Dim P As New Point(es.Left, es.Top)
                es.Copy()
                es.Delete()
                Application.ActiveSheet.PasteSpecial(Format:="図 (JPEG)")
                Application.Selection.Top = P.Y
                Application.Selection.Left = P.X
            End If
        Next
    End Sub
End Module
