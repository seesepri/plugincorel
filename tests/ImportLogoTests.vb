Option Strict Off
Imports System
Imports System.Collections.Generic

' =====================================================================
' Objek COM tiruan yang meniru perilaku object model CorelDRAW untuk
' impor logo. Detail penting: objek hasil impor baru masuk halaman
' SETELAH ImportFilter.Finish() dipanggil.
' =====================================================================

Public Class FakeShape
    Public Property Name As String = ""
    Public Property Type As Integer = 2       ' cdrBitmapShape
End Class

Public Class FakeShapeList
    Private ReadOnly _items As New List(Of FakeShape)()
    Public Sub AddItem(s As FakeShape)
        _items.Add(s)
    End Sub
    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property
    Default Public ReadOnly Property Item(i As Integer) As FakeShape    ' 1-based seperti COM
        Get
            Return _items(i - 1)
        End Get
    End Property
End Class

Public Class FakeImportFilter
    Private ReadOnly _layer As FakeLayer
    Public Sub New(layer As FakeLayer)
        _layer = layer
    End Sub
    ''' <summary>Perilaku nyata: objek masuk ke layer saat Finish dipanggil.</summary>
    Public Sub Finish()
        If _layer.FinishWorks Then _layer.Shapes.AddItem(New FakeShape With {.Type = 2})
    End Sub
End Class

Public Class FakeLayer
    Public ReadOnly Shapes As New FakeShapeList()
    ' ok | importExThrows (ImportEx gagal, Import jalan) | throwAll | finishNoop
    Public Mode As String = "ok"
    Public FinishWorks As Boolean = True
    Public LastFilter As Integer = -999
    Public LastHadOptions As Boolean = False
    Public UsedOptions As Object = Nothing
    Public ImportExCalls As Integer = 0
    Public ImportCalls As Integer = 0
    Public ThrowMessage As String =
        "Type mismatch. (Exception from HRESULT: 0x80020005 (DISP_E_TYPEMISMATCH))"

    Public Function ImportEx(fileName As String, filter As Integer, Optional options As Object = Nothing) As Object
        ImportExCalls += 1
        LastFilter = filter
        LastHadOptions = (options IsNot Nothing)
        UsedOptions = options
        If Mode = "importExThrows" OrElse Mode = "throwAll" Then
            Throw New InvalidOperationException(ThrowMessage)
        End If
        If Mode = "finishNoop" Then FinishWorks = False
        Return New FakeImportFilter(Me)
    End Function

    Public Sub Import(fileName As String, Optional filter As Integer = 0, Optional options As Object = Nothing)
        ImportCalls += 1
        UsedOptions = options
        If Mode = "throwAll" Then Throw New InvalidOperationException(ThrowMessage)
        Shapes.AddItem(New FakeShape With {.Type = 2})
    End Sub
End Class

Public Class FakeOptions
    Public Property Mode As Integer = 0
    Public Property LinkBitmapExternally As Boolean = True
    Public Property MaintainLayers As Boolean = True
    Public Property CombineMultilayerBitmaps As Boolean = False
End Class

Public Class FakeApp
    Public Property VersionMajor As Integer = 24
    Public Calls As Integer = 0
    Public Fail As Boolean = False
    Public Function CreateStructImportOptions() As Object
        Calls += 1
        If Fail Then Throw New InvalidOperationException("tidak didukung versi ini")
        Return New FakeOptions()
    End Function
End Class

Public Class FakeDoc
    Public Property Selection As Object = Nothing
    Public Property Application As Object = Nothing
End Class

Module ImportLogoTests
    Public Sub Run()
        Dim eng As New WatermarkEngine()

        TestMain.Section("Filter impor berdasarkan ekstensi")
        TestMain.Check(WatermarkEngine.FilterForFile("logo.png") = 802, ".png = cdrPNG (802)")
        TestMain.Check(WatermarkEngine.FilterForFile("LOGO.PNG") = 802, ".PNG huruf besar")
        TestMain.Check(WatermarkEngine.FilterForFile("a.jpg") = 774, ".jpg = cdrJPEG (774)")
        TestMain.Check(WatermarkEngine.FilterForFile("a.jpeg") = 774, ".jpeg = cdrJPEG (774)")
        TestMain.Check(WatermarkEngine.FilterForFile("a.tiff") = 772, ".tiff = cdrTIFF (772)")
        TestMain.Check(WatermarkEngine.FilterForFile("a.bmp") = 769, ".bmp = cdrBMP (769)")
        TestMain.Check(WatermarkEngine.FilterForFile("a.gif") = 773, ".gif = cdrGIF (773)")
        TestMain.Check(WatermarkEngine.FilterForFile("a.psd") = 788, ".psd = cdrPSD (788)")
        TestMain.Check(WatermarkEngine.FilterForFile("a.cdr") = 0, "tak dikenal = cdrAutoSense (0)")
        TestMain.Check(WatermarkEngine.FilterForFile("") = 0, "nama kosong = cdrAutoSense")
        TestMain.Check(WatermarkEngine.FilterForFile(Nothing) = 0, "Nothing = cdrAutoSense")

        TestMain.Section("Impor logo lewat Application.CreateStructImportOptions (jalur utama)")
        Dim app1 As New FakeApp()
        Dim layer1 As New FakeLayer()
        Dim shape1 As Object = eng.ImportLogo(New FakeDoc With {.Application = app1}, layer1, "C:\logo\seta.jpg")
        TestMain.Check(shape1 IsNot Nothing, "impor berhasil mengembalikan shape")
        TestMain.Check(app1.Calls = 1, "Application.CreateStructImportOptions() dipanggil")
        TestMain.Check(layer1.LastFilter = 774, "ImportEx memakai filter JPEG (774)")
        TestMain.Check(layer1.LastHadOptions, "Options dikirim sebagai objek (bukan Nothing)")
        TestMain.Check(TypeOf layer1.UsedOptions Is FakeOptions, "options berasal dari aplikasi")
        Dim opt1 As FakeOptions = CType(layer1.UsedOptions, FakeOptions)
        TestMain.Check(opt1.Mode = 1, "Mode = cdrImportFull (1)")
        TestMain.Check(Not opt1.LinkBitmapExternally, "bitmap ditanam, bukan ditautkan")
        TestMain.Check(Not opt1.MaintainLayers, "MaintainLayers = False")
        TestMain.Check(layer1.Shapes.Count = 1, "tepat satu objek masuk (tidak dobel)")

        TestMain.Section("Fallback saat API/versi bermasalah")
        Dim app2 As New FakeApp With {.Fail = True}
        Dim layer2 As New FakeLayer()
        Dim shape2 As Object = eng.ImportLogo(New FakeDoc With {.Application = app2}, layer2, "C:\logo\a.png")
        TestMain.Check(shape2 IsNot Nothing, "CreateStructImportOptions gagal -> impor tetap berhasil")
        TestMain.Check(Not layer2.LastHadOptions, "tanpa options -> dipanggil dengan argumen minimal")

        Dim layer3 As New FakeLayer With {.Mode = "importExThrows"}
        Dim shape3 As Object = eng.ImportLogo(New FakeDoc With {.Application = New FakeApp()}, layer3, "C:\logo\a.png")
        TestMain.Check(shape3 IsNot Nothing, "ImportEx melempar DISP_E_TYPEMISMATCH -> fallback Import berhasil")
        TestMain.Check(layer3.ImportCalls >= 1, "Layer.Import dipakai sebagai fallback")

        Dim layer4 As New FakeLayer With {.Mode = "finishNoop"}
        Dim shape4 As Object = eng.ImportLogo(New FakeDoc With {.Application = New FakeApp()}, layer4, "C:\logo\a.png")
        TestMain.Check(shape4 IsNot Nothing, "Finish() tidak menghasilkan objek -> fallback berikutnya dipakai")

        TestMain.Section("Semua jalur gagal")
        Dim layer5 As New FakeLayer With {.Mode = "throwAll"}
        Dim pesan As String = ""
        Try
            eng.ImportLogo(New FakeDoc With {.Application = New FakeApp With {.Fail = True}}, layer5, "C:\logo\seta.jpg")
        Catch ex As Exception
            pesan = ex.Message
        End Try
        TestMain.Check(pesan.Contains("tidak dapat diimpor"), "pesan jelas untuk pengguna")
        TestMain.Check(pesan.Contains("seta.jpg"), "pesan menyebut nama file")
        TestMain.Check(pesan.Contains("Percobaan:"), "pesan memuat rantai percobaan")
        TestMain.Check(pesan.Contains("0x80020005") OrElse pesan.Contains("Type mismatch"), "pesan memuat detail COM asli")
        TestMain.Check(pesan.Contains("ImportEx") OrElse pesan.Contains("Import"), "pesan menyebut metode yang dicoba")
    End Sub
End Module
