Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  FakeCorel.vb
'  Pengganti objek COM CorelDRAW untuk unit test.
'
'  Semua anggota sengaja publik dan memakai nama properti yang sama dengan API
'  CorelDRAW (Unit, Pages, Layers, Shapes, CenterX, SizeWidth, SaveAs, ...)
'  supaya kode produksi yang memakai late binding bisa dijalankan tanpa
'  CorelDRAW. TIDAK ada pemanggilan COM di sini.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.IO

Public Class FakeShape
    Private _name As String = "Shape"
    Private _typeCode As Integer = 3
    Private _centerX As Double = 0
    Private _centerY As Double = 0
    Private _width As Double = 10
    Private _height As Double = 10

    Public Property [Type] As Integer
        Get
            Return _typeCode
        End Get
        Set(value As Integer)
            _typeCode = value
        End Set
    End Property

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    Public Property CenterX As Double
        Get
            If FailGeometryRead Then Throw New InvalidOperationException("simulasi: CenterX tidak terbaca")
            Return _centerX
        End Get
        Set(value As Double)
            _centerX = value
        End Set
    End Property

    Public Property CenterY As Double
        Get
            Return _centerY
        End Get
        Set(value As Double)
            _centerY = value
        End Set
    End Property

    Public Property SizeWidth As Double
        Get
            Return _width
        End Get
        Set(value As Double)
            _width = value
        End Set
    End Property

    Public Property SizeHeight As Double
        Get
            Return _height
        End Get
        Set(value As Double)
            _height = value
        End Set
    End Property

    Public Property Layer As FakeLayer = Nothing
    Public Property Text As New FakeText()
    Public Property Outline As New FakeOutline()
    Public Property Fill As New FakeFill()
    Public Property Transparency As New FakeTransparency()
    Public Property Children As New List(Of FakeShape)()
    Public Property Deleted As Boolean = False
    ''' <summary>Bila True, membaca CenterX melempar exception (simulasi objek rusak).</summary>
    Public Property FailGeometryRead As Boolean = False

    Public Sub New()
    End Sub

    Public Sub New(typeCode As Integer, centerX As Double, centerY As Double, width As Double, height As Double)
        Me.Type = typeCode
        Me.CenterX = centerX
        Me.CenterY = centerY
        Me.SizeWidth = width
        Me.SizeHeight = height
    End Sub

    Public Sub Rotate(degrees As Double)
        RotationDegrees = degrees
    End Sub

    Public Property RotationDegrees As Double = 0

    Public Function Duplicate(dx As Double, dy As Double) As FakeShape
        Dim copy As New FakeShape(Type, CenterX + dx, CenterY + dy, SizeWidth, SizeHeight)
        copy.Name = Name
        copy.Layer = Layer
        If Layer IsNot Nothing Then Layer.Shapes.Add(copy)
        Return copy
    End Function

    ''' <summary>Disalin ke clipboard palsu oleh selection.Copy().</summary>
    Public Property Copied As Boolean = False

    ''' <summary>Dipilih lewat AddToSelection pada dokumen induk.</summary>
    Public Property Selected As Boolean = False

    ''' <summary>Bila True, AddToSelection tidak berefek (simulasi objek tak dapat dipilih).</summary>
    Public Property FailAddToSelection As Boolean = False

    ''' <summary>Meniru Shape.AddToSelection(): hanya menandai, tidak mengubah geometri.</summary>
    Public Sub AddToSelection()
        If FailAddToSelection Then Return
        Selected = True
        If Layer IsNot Nothing AndAlso Layer.ParentPage IsNot Nothing Then
            Dim doc As FakeDocument = Layer.ParentPage.OwnerDocument
            If doc IsNot Nothing AndAlso Not doc.Selection.Contains(Me) Then doc.Selection.Add(Me)
        End If
    End Sub

    Public Sub Delete()
        Deleted = True
        If Layer IsNot Nothing Then Layer.Shapes.Remove(Me)
    End Sub

    ''' <summary>Koleksi anak untuk grup (dipakai IsWatermarkShape).</summary>
    Public ReadOnly Property Shapes As FakeShapeCollection
        Get
            Return New FakeShapeCollection(Children)
        End Get
    End Property
End Class

Public Class FakeText
    Public Property Story As New FakeStory()
End Class

Public Class FakeStory
    Public Property Font As String = ""
    Public Property Size As Single = 12.0F
End Class

Public Class FakeOutline
    Public Property NoOutline As Boolean = False

    Public Sub SetNoOutline()
        NoOutline = True
    End Sub
End Class

Public Class FakeFill
    Public Property UniformColor As New FakeColor()
End Class

Public Class FakeColor
    Public Property R As Integer = 0
    Public Property G As Integer = 0
    Public Property B As Integer = 0

    Public Sub RGBAssign(red As Integer, green As Integer, blue As Integer)
        R = red
        G = green
        B = blue
    End Sub
End Class

Public Class FakeTransparency
    Public Property Uniform As Integer = 0
    Public Property AppliedTo As Integer = 0
    Public Property MergeMode As Integer = 0

    Public Sub ApplyUniformTransparency(value As Integer)
        Uniform = value
    End Sub
End Class

''' <summary>Koleksi shape ala CorelDRAW: 1-based, urutan z-order dari BAWAH ke ATAS.</summary>
Public Class FakeShapeCollection
    Private ReadOnly _items As List(Of FakeShape)

    Public Sub New(items As List(Of FakeShape))
        _items = items
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakeShape
        Get
            If oneBasedIndex < 1 OrElse oneBasedIndex > _items.Count Then
                Throw New ArgumentOutOfRangeException(NameOf(oneBasedIndex))
            End If
            Return _items(oneBasedIndex - 1)
        End Get
    End Property

    Public Sub Add(shape As FakeShape)
        _items.Add(shape)
    End Sub

    Public Sub Remove(shape As FakeShape)
        _items.Remove(shape)
    End Sub

    Public ReadOnly Property All As List(Of FakeShape)
        Get
            Return _items
        End Get
    End Property
End Class

Public Class FakeLayer
    Public Property Name As String = "Layer 1"
    Public Property Editable As Boolean = True
    Public Property Visible As Boolean = True
    Public Property Printable As Boolean = True
    Public Property ParentPage As FakePage = Nothing
    Public ReadOnly Property Items As New List(Of FakeShape)()
    ''' <summary>Dipanggil saat CreateArtisticText; bisa diisi test untuk memicu kegagalan.</summary>
    Public Property FailCreateText As Boolean = False

    Public ReadOnly Property Shapes As FakeShapeCollection
        Get
            Return New FakeShapeCollection(Items)
        End Get
    End Property

    ''' <summary>Meniru Layer.Paste(): isi clipboard masuk ke layer ini.</summary>
    Public Sub Paste()
        If FakeClipboard.Items Is Nothing Then Return
        For Each source As FakeShape In FakeClipboard.Items
            Dim copy As FakeShape = source.Duplicate(0, 0)
            copy.Layer = Me
            Items.Add(copy)
        Next
    End Sub

    Public Sub MoveToTop()
        MoveToTopCalled = True
    End Sub

    Public Property MoveToTopCalled As Boolean = False

    Public Function CreateArtisticText(x As Double, y As Double, text As String,
                                       startIndex As Integer, endIndex As Integer,
                                       fontName As String, fontSize As Single) As FakeShape
        If FailCreateText Then Throw New InvalidOperationException("simulasi: CreateArtisticText gagal")
        Dim shape As New FakeShape(6, x, y, 40, 8)
        shape.Name = "Text " & (Items.Count + 1)
        shape.Layer = Me
        shape.Text.Story.Font = fontName
        shape.Text.Story.Size = fontSize
        Items.Add(shape)
        Return shape
    End Function

    Public Function Import(filePath As String) As FakeShape
        Dim shape As New FakeShape(5, 0, 0, 20, 20)
        shape.Name = Path.GetFileName(filePath)
        shape.Layer = Me
        Items.Add(shape)
        Return shape
    End Function
End Class

Public Class FakeLayers
    Private ReadOnly _items As New List(Of FakeLayer)()

    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakeLayer
        Get
            Return _items(oneBasedIndex - 1)
        End Get
    End Property

    Public Sub Add(layer As FakeLayer)
        _items.Add(layer)
    End Sub

    Public ReadOnly Property All As List(Of FakeLayer)
        Get
            Return _items
        End Get
    End Property
End Class

Public Class FakePage
    ''' <summary>Dokumen pemilik halaman (dipakai AddToSelection).</summary>
    Public Property OwnerDocument As FakeDocument = Nothing

    Public Property Index As Integer = 1
    Public Property SizeWidth As Double = 210
    Public Property SizeHeight As Double = 297
    Public Property Layers As New FakeLayers()
    Public Property ActiveLayer As FakeLayer = Nothing
    Public Property Activated As Boolean = False
    ''' <summary>Bila True, SizeWidth melempar exception (simulasi halaman rusak).</summary>
    Public Property FailSizeRead As Boolean = False

    ''' <summary>Bila True, semua layer baru pada halaman ini gagal membuat teks.</summary>
    Public Property FailCreateText As Boolean = False

    Public Sub New()
    End Sub

    Public Sub New(pageIndex As Integer)
        Me.Index = pageIndex
    End Sub

    ''' <summary>Bila True, Activate melempar (simulasi halaman tidak dapat diakses).</summary>
    Public Property FailActivate As Boolean = False

    Public Sub Activate()
        If FailActivate Then Throw New InvalidOperationException("simulasi: halaman tidak dapat diaktifkan")
        Activated = True
    End Sub

    ''' <summary>Meniru CorelDRAW: paste menambahkan shape hasil salinan ke layer aktif.</summary>
    Public Sub Paste()
        EnsureActiveLayer().Paste()
    End Sub

    ''' <summary>
    ''' Meniru Page.Shapes: seluruh objek pada semua layer halaman.
    ''' Dipakai AutoSelectService (pemindaian dimulai dari koleksi halaman).
    ''' </summary>
    Public ReadOnly Property Shapes As FakeShapeCollection
        Get
            Dim all As New List(Of FakeShape)()
            For Each layer As FakeLayer In Layers.All
                all.AddRange(layer.Items)
            Next
            Return New FakeShapeCollection(all)
        End Get
    End Property



    Public Function CreateLayer(layerName As String) As FakeLayer
        Dim layer As New FakeLayer()
        layer.Name = layerName
        layer.ParentPage = Me
        layer.FailCreateText = FailCreateText
        Layers.Add(layer)
        Return layer
    End Function

    ''' <summary>Pastikan ada layer aktif untuk menampung shape (meniru dokumen nyata).</summary>
    Public Function EnsureActiveLayer() As FakeLayer
        If ActiveLayer Is Nothing Then
            ActiveLayer = CreateLayer("Layer 1")
        End If
        Return ActiveLayer
    End Function
End Class

Public Class FakePages
    Private ReadOnly _items As New List(Of FakePage)()

    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakePage
        Get
            If oneBasedIndex < 1 OrElse oneBasedIndex > _items.Count Then
                Throw New ArgumentOutOfRangeException(NameOf(oneBasedIndex))
            End If
            Return _items(oneBasedIndex - 1)
        End Get
    End Property

    ''' <summary>Dokumen pemilik (diisi oleh FakeDocument saat halaman ditambahkan).</summary>
    Public Property Owner As FakeDocument = Nothing

    Public Sub Add(page As FakePage)
        page.Index = _items.Count + 1
        page.OwnerDocument = Owner
        _items.Add(page)
    End Sub

    Public ReadOnly Property All As List(Of FakePage)
        Get
            Return _items
        End Get
    End Property
End Class

Public Class FakeApplication
    Public Property RefreshCount As Integer = 0

    Public Sub Refresh()
        RefreshCount += 1
    End Sub
End Class

Public Class FakeDocument
    Private _unit As Integer = 3

    Public Property FilePath As String = ""
    ''' <summary>Path saat dibuka (tidak berubah walau SaveAs dipanggil).</summary>
    Public Property OriginalPath As String = ""
    Public Property Name As String = "dokumen uji"
    Public ReadOnly Property Pages As New FakePages()
    Public Property ActivePage As FakePage = Nothing
    Public Property App As New FakeApplication()

    ' --- pengaturan perilaku untuk test ------------------------------------
    Public Property FailUnitRead As Boolean = False
    Public Property FailUnitSet As Boolean = False
    Public Property FailUnitRestore As Boolean = False
    ''' <summary>
    ''' Satuan tujuan yang penyetelannya SELALU gagal (mis. 1 = inch) — dipakai
    ''' untuk menguji kegagalan PEMULIHAN satuan tanpa menggagalkan normalisasi mm.
    ''' </summary>
    Public Property FailUnitRestoreTo As Integer = -2
    Public Property FailBeginCommandGroup As Boolean = False
    Public Property FailEndCommandGroup As Boolean = False
    Public Property SaveFails As Boolean = False
    Public Property SaveWritesFile As Boolean = True
    Public Property CloseFails As Boolean = False

    ' --- rekaman pemanggilan ----------------------------------------------
    Public ReadOnly Property SavedAs As New List(Of String)()
    Public ReadOnly Property CommandGroupsBegun As New List(Of String)()
    Public Property CommandGroupOpen As Boolean = False
    Public Property Closed As Boolean = False
    Public Property UnitSetAttempts As New List(Of Integer)()

    Public Sub New()
        App = New FakeApplication()
        Pages.Owner = Me
    End Sub

    Public Sub New(sourcePath As String)
        Me.FilePath = sourcePath
        Me.OriginalPath = sourcePath
        Me.Name = Path.GetFileName(sourcePath)
        App = New FakeApplication()
        Pages.Owner = Me
    End Sub

    Public Property Unit As Integer
        Get
            If FailUnitRead Then Throw New InvalidOperationException("simulasi: Document.Unit tidak terbaca")
            Return _unit
        End Get
        Set(value As Integer)
            UnitSetAttempts.Add(value)
            If FailUnitSet AndAlso value <> _unit Then
                Throw New InvalidOperationException("simulasi: Document.Unit tidak dapat diubah")
            End If
            If (FailUnitRestore AndAlso value = 3 AndAlso _unit <> 3) OrElse
               (FailUnitRestoreTo = value AndAlso value <> _unit) Then
                _unit = 3
                Throw New InvalidOperationException("simulasi: pemulihan satuan gagal")
            End If
            _unit = value
        End Set
    End Property

    Public Property Application As FakeApplication
        Get
            Return App
        End Get
        Set(value As FakeApplication)
            App = value
        End Set
    End Property

    Public Sub BeginCommandGroup(name As String)
        If FailBeginCommandGroup Then Throw New InvalidOperationException("simulasi: BeginCommandGroup gagal")
        CommandGroupsBegun.Add(name)
        CommandGroupOpen = True
    End Sub

    Public Sub EndCommandGroup()
        If Not CommandGroupOpen Then Throw New InvalidOperationException("simulasi: EndCommandGroup tanpa group terbuka")
        If FailEndCommandGroup Then Throw New InvalidOperationException("simulasi: EndCommandGroup gagal")
        CommandGroupOpen = False
    End Sub

    Public Sub SaveAs(targetFile As String)
        If SaveFails Then Throw New IOException("simulasi: SaveAs gagal")
        SavedAs.Add(targetFile)
        If SaveWritesFile Then
            Dim folder As String = Path.GetDirectoryName(targetFile)
            If Not String.IsNullOrEmpty(folder) Then Directory.CreateDirectory(folder)
            File.WriteAllText(targetFile, "fake cdr content")
        End If
        FilePath = targetFile
        Name = Path.GetFileName(targetFile)
    End Sub

    Public Sub Close()
        If CloseFails Then Throw New InvalidOperationException("simulasi: Close gagal")
        Closed = True
    End Sub

    ' --- PDF proof (Fokus C) ------------------------------------------------
    ''' <summary>Byte PDF yang ditulis oleh PublishToPDF. Nothing = PDF rusak/omong kosong.</summary>
    ' --- seleksi (untuk CopyService & AutoSelectService) --------------------
    ''' <summary>Shape yang saat ini terpilih.</summary>
    Public ReadOnly Property Selection As New List(Of FakeShape)

    ''' <summary>Jumlah pemanggilan ClearSelection (dipakai test "tidak destruktif").</summary>
    Public Property ClearSelectionCalls As Integer = 0

    ''' <summary>Bila True, ClearSelection melempar (simulasi COM gagal).</summary>
    Public Property FailClearSelection As Boolean = False

    Public Sub ClearSelection()
        ClearSelectionCalls += 1
        If FailClearSelection Then Throw New InvalidOperationException("simulasi: ClearSelection gagal")
        Selection.Clear()
    End Sub

    ''' <summary>Meniru Application.ActiveSelectionRange (Count saja yang dipakai).</summary>
    Public ReadOnly Property ActiveSelectionRange As FakeSelectionRange
        Get
            Return New FakeSelectionRange(Selection, Me)
        End Get
    End Property

    ''' <summary>Meniru Selection.Copy(): isi clipboard dari shape terpilih.</summary>
    Public Sub CopySelection()
        FakeClipboard.Items = New List(Of FakeShape)(Selection)
        For Each shape As FakeShape In Selection
            shape.Copied = True
        Next
    End Sub

    ' --- PDF proof (Fokus C) ------------------------------------------------
    Public Property PdfBytes As Byte() = Nothing
    ''' <summary>True = PublishToPDF melempar exception.</summary>
    Public Property PublishFails As Boolean = False
    ''' <summary>False = PublishToPDF "sukses" tetapi tidak menulis file (kasus nyata di CorelDRAW).</summary>
    Public Property PublishWritesFile As Boolean = True
    Public ReadOnly Property PublishedPdfPaths As New List(Of String)()

    Public Sub PublishToPDF(targetFile As String)
        If PublishFails Then Throw New InvalidOperationException("simulasi: PublishToPDF gagal")
        PublishedPdfPaths.Add(targetFile)
        If Not PublishWritesFile Then Return
        Dim folder As String = Path.GetDirectoryName(targetFile)
        If Not String.IsNullOrEmpty(folder) Then Directory.CreateDirectory(folder)
        If PdfBytes Is Nothing Then
            File.WriteAllText(targetFile, "bukan pdf")
        Else
            File.WriteAllBytes(targetFile, PdfBytes)
        End If
    End Sub

    Public Sub ExportEx(targetFile As String, filter As Integer, options As Integer)
        PublishToPDF(targetFile)
    End Sub
End Class

''' <summary>Aplikasi CorelDRAW palsu: OpenDocument mengembalikan FakeDocument.</summary>
Public Class FakeCorelApplication
    Public ReadOnly Property Opened As New List(Of FakeDocument)()
    Public Property FailOpen As Boolean = False

    ''' <summary>Bila True, file tertentu (berdasarkan nama) gagal dibuka.</summary>
    Public Property FailOpenFor As New List(Of String)()

    ''' <summary>Jumlah dokumen yang masih terbuka (belum Close) — meniru app.Documents.</summary>
    Public ReadOnly Property Documents As New FakeDocuments()

    ''' <summary>Dokumen "aktif" untuk service yang membaca app.ActiveDocument.</summary>
    Public Property ActiveDocumentInstance As FakeDocument = Nothing

    Public ReadOnly Property ActiveDocument As FakeDocument
        Get
            Return ActiveDocumentInstance
        End Get
    End Property

    Public ReadOnly Property ActiveSelectionRange As FakeSelectionRange
        Get
            If ActiveDocumentInstance Is Nothing Then Return New FakeSelectionRange(New List(Of FakeShape)())
            Return ActiveDocumentInstance.ActiveSelectionRange
        End Get
    End Property

    Public Property Application As FakeApplication = New FakeApplication()

    ''' <summary>Isi test untuk mengubah dokumen yang baru dibuka (mis. buat halaman).</summary>
    Public Property Configure As Action(Of FakeDocument) = Nothing

    Public Function OpenDocument(sourcePath As String) As FakeDocument
        If FailOpen Then Throw New IOException("simulasi: OpenDocument gagal")
        For Each blocked As String In FailOpenFor
            If String.Equals(Path.GetFileName(sourcePath), blocked, StringComparison.OrdinalIgnoreCase) Then
                Throw New IOException("simulasi: file sedang dipakai aplikasi lain")
            End If
        Next
        Dim doc As New FakeDocument(sourcePath)
        doc.Application = Application
        doc.Pages.Add(New FakePage(1))
        doc.Pages(1).EnsureActiveLayer()
        doc.ActivePage = doc.Pages(1)
        If Configure IsNot Nothing Then Configure.Invoke(doc)
        Opened.Add(doc)
        Documents.NotifyOpened(doc)
        Return doc
    End Function
End Class

''' <summary>Clipboard palsu bersama (satu proses test = satu clipboard).</summary>
Public Class FakeClipboard
    Public Shared Property Items As List(Of FakeShape) = Nothing

    Public Shared Sub Clear()
        Items = Nothing
    End Sub
End Class

''' <summary>Meniru SelectionRange / selection shape: Count + Copy() + item.</summary>
Public Class FakeSelectionRange
    Private ReadOnly _items As List(Of FakeShape)
    Private ReadOnly _document As FakeDocument

    Public Sub New(items As List(Of FakeShape), Optional document As FakeDocument = Nothing)
        _items = items
        _document = document
    End Sub

    ''' <summary>Meniru Selection.Copy() (isi clipboard dokumen).</summary>
    Public Sub Copy()
        If _document IsNot Nothing Then _document.CopySelection()
    End Sub

    Public ReadOnly Property Count As Integer
        Get
            Return _items.Count
        End Get
    End Property

    Public Default ReadOnly Property Item(oneBasedIndex As Integer) As FakeShape
        Get
            Return _items(oneBasedIndex - 1)
        End Get
    End Property

    Public ReadOnly Property All As List(Of FakeShape)
        Get
            Return _items
        End Get
    End Property
End Class

''' <summary>Koleksi dokumen terbuka (meniru Application.Documents.Count).</summary>
Public Class FakeDocuments
    Private ReadOnly _open As New List(Of FakeDocument)()

    Public ReadOnly Property Count As Integer
        Get
            Dim live As Integer = 0
            For Each doc As FakeDocument In _open
                If Not doc.Closed Then live += 1
            Next
            Return live
        End Get
    End Property

    Public Sub NotifyOpened(doc As FakeDocument)
        If doc IsNot Nothing Then _open.Add(doc)
    End Sub
End Class
