Imports System.IO
Imports System.Runtime.InteropServices
Imports GemBox.Document

Public Class Form1
    Private tempPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\Screenshot_Temp_Folder\"
    Private TestResultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Test_Results\"
    Dim line As String

    <DllImport("gdi32.dll")>
    Private Shared Function BitBlt(ByVal hdc As IntPtr,
                                   ByVal nXDest As Integer,
                                   ByVal nYDest As Integer,
                                   ByVal nWidth As Integer,
                                   ByVal nHeight As Integer,
                                   ByVal hdcSrc As IntPtr,
                                   ByVal nXSrc As Integer,
                                   ByVal nYSrc As Integer,
                                   ByVal dwRop As CopyPixelOperation) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(
     ByVal lpClassName As String,
     ByVal lpWindowName As String) As IntPtr
    End Function

    Function GetWindowImage(ByVal WindowHandle As IntPtr, ByVal Area As Rectangle) As Bitmap
        Dim b As New Bitmap(Area.Width, Area.Height, Imaging.PixelFormat.Format24bppRgb)
        Using img As Graphics = Graphics.FromImage(b)
            Dim ImageHDC As IntPtr = img.GetHdc
            Using window As Graphics = Graphics.FromHwnd(WindowHandle)
                Dim WindowHDC As IntPtr = window.GetHdc
                BitBlt(ImageHDC, 0, 0, Area.Width, Area.Height, WindowHDC, Area.X, Area.Y, CopyPixelOperation.SourceCopy)
                window.ReleaseHdc()
            End Using
            img.ReleaseHdc()
        End Using
        Return b
    End Function

    Private Sub ScreenShot_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Directory.CreateDirectory(tempPath)
        Directory.CreateDirectory(TestResultPath)
        Dim FileName = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")
        Dim window As IntPtr = FindWindow(Nothing, "Untitled - Notepad")
        Dim snapshot As Bitmap = GetWindowImage(window, New Rectangle(0, 0, 1366, 768))
        snapshot.Save(tempPath + FileName + ".jpeg", Imaging.ImageFormat.Jpeg)
        CheckBox1_CheckedChanged(sender, e, FileName)
        Return
    End Sub

    Private Sub CommentBox_TextChanged(sender As Object, e As EventArgs, FileName As String)
        Dim comments = TextBox1.Text
        WriteTextFile(FileName, comments)
        TextBox1.Clear()
        Return
    End Sub

    Private Sub WriteTextFile(FileName As String, comments As String)
        Dim FILE_NAME As String = tempPath & "temptext.txt"
        If File.Exists(FILE_NAME) = False Then
            File.Create(FILE_NAME).Dispose()
        End If
        Dim objWriter As New StreamWriter(FILE_NAME, True)
        objWriter.WriteLine(FileName + ": " + comments)
        objWriter.Close()
        Return
    End Sub

    Private Sub ReadTextFile(PicName As String)
        Try
            Using sr As StreamReader = New StreamReader(CStr(tempPath + "temptext.txt"))
                line = sr.ReadLine()
                While (line <> Nothing)
                    If line.Contains(PicName) Then
                        Return
                    End If
                    line = sr.ReadLine()
                End While
            End Using
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
        Return
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs, FileName As String)
        If CheckBox1.Checked = True Then
            CommentBox_TextChanged(sender, e, FileName)
            CheckBox1.CheckState = 0
            Return
        End If
    End Sub

    Private Sub WriteWordDocument()
        ComponentInfo.SetLicense("FREE-LIMITED-KEY")
        Dim document As New DocumentModel()
        Dim di As New IO.DirectoryInfo(tempPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.jpeg")
        Dim fi As IO.FileInfo
        For Each fi In aryFi
            Dim FN As String = fi.Name
            Dim PicName = FN.Split(".jpeg")
            Dim section As New Section(document)
            document.Sections.Add(section)
            Dim paragraph As New Paragraph(document)
            section.Blocks.Add(paragraph)
            Dim picture1 As New Picture(document, CStr(tempPath + FN))
            paragraph.Inlines.Add(picture1)
            ReadTextFile(PicName(0))
            If line <> Nothing Then
                Dim Comments = line.Split(PicName(0) + ":")
                Dim run As New Run(document, "Comment:" + Comments(1))
                paragraph.Inlines.Add(run)
            End If
            Dim layout2 As New FloatingLayout(
                New HorizontalPosition(HorizontalPositionType.Left, HorizontalPositionAnchor.Page),
                New VerticalPosition(2, LengthUnit.Inch, VerticalPositionAnchor.TopMargin),
                New Size(600, 400))
            layout2.WrappingStyle = TextWrappingStyle.InFrontOfText
            picture1.Layout = layout2
        Next
        Dim ResultTime = DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")
        document.Save(TestResultPath + "TestResult_" + ResultTime + ".docx")
        Return
    End Sub

    Private Sub CreateDocument_Click(sender As Object, e As EventArgs) Handles CreateDocument.Click
        WriteWordDocument()
        Try
            For Each filepath As String In Directory.GetFiles(tempPath)
                File.Delete(filepath)
            Next
            Directory.Delete(tempPath)
        Catch f As Exception
            Console.WriteLine(f.Message)
        End Try
        Close()
    End Sub
End Class