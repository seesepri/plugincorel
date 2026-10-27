Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ShapeNode.vb
'  Abstraksi satu objek CorelDRAW untuk pemindaian Auto Select.
'
'  Kenapa abstraksi:
'    - Logika pemindaian (ShapeScanner) harus bisa diuji tanpa CorelDRAW.
'      Test memakai node palsu; produksi memakai ShapeNodeAdapter (COM).
'    - Semua pembacaan properti COM dilakukan di satu tempat (adapter), dengan
'      pesan error tersimpan — tidak ada Catch kosong yang menyembunyikan
'      kegagalan.
'
'  Yang TIDAK dilakukan di sini: tidak ada operasi yang mengubah atau
'  menghapus objek. Hanya membaca properti dan (lewat adapter) koleksi anak.
' ============================================================================

Imports System
Imports System.Collections.Generic

''' <summary>Satu node objek: fakta klasifikasi + anak-anaknya.</summary>
Public Interface IShapeNode
    ''' <summary>Fakta untuk klasifikasi (tipe, kurva, teks, nama, layer).</summary>
    Function GetFacts() As ShapeFacts

    ''' <summary>Anak grup (1..Count). Kosong untuk objek non-grup.</summary>
    Function GetChildren() As IList(Of IShapeNode)

    ''' <summary>Pesan error bila anak grup tidak dapat dibaca (kosong = sukses).</summary>
    ReadOnly Property ChildrenError As String

    ''' <summary>True bila node ini container PowerClip berisi objek.</summary>
    ReadOnly Property IsPowerClipContainer As Boolean

    ''' <summary>Isi PowerClip (objek di dalam container).</summary>
    Function GetPowerClipChildren() As IList(Of IShapeNode)

    ''' <summary>Pesan error bila isi PowerClip tidak dapat dibaca.</summary>
    ReadOnly Property PowerClipError As String

    ''' <summary>
    ''' Objek yang harus ditambahkan ke seleksi untuk node ini. Untuk objek di
    ''' dalam PowerClip, ini adalah CONTAINER-nya (isi PowerClip tidak dapat
    ''' dipilih sendiri di CorelDRAW).
    ''' </summary>
    ReadOnly Property SelectionTarget As Object

    ''' <summary>Objek COM asli (untuk diagnosa/log).</summary>
    ReadOnly Property WrappedShape As Object
End Interface

''' <summary>
''' Adapter COM (late bound) untuk IShapeNode. Sama seperti bagian lain proyek,
''' tidak memakai referensi interop CorelDRAW: semua properti dibaca dinamis.
''' </summary>
Public Class ShapeNodeAdapter
    Implements IShapeNode

    Private ReadOnly _shape As Object
    Private ReadOnly _selectionTarget As Object
    Private ReadOnly _depthHint As Integer

    Private _facts As ShapeFacts = Nothing
    Private _children As IList(Of IShapeNode) = Nothing
    Private _childrenLoaded As Boolean = False
    Private _childrenError As String = ""
    Private _powerClipLoaded As Boolean = False
    Private _powerClipChildren As IList(Of IShapeNode) = Nothing
    Private _powerClipError As String = ""
    Private _isPowerClip As Boolean = False

    ''' <summary>Batas kedalaman mulai dari node akar (mencegah rekursi tak berujung).</summary>
    Public Const MaxDepth As Integer = 32

    Public Sub New(shape As Object, Optional selectionTarget As Object = Nothing, Optional depthHint As Integer = 0)
        _shape = shape
        _selectionTarget = If(selectionTarget Is Nothing, shape, selectionTarget)
        _depthHint = depthHint
    End Sub

    ''' <summary>Bungkus koleksi shape (mis. Page.Shapes) menjadi daftar node.</summary>
    Public Shared Function WrapCollection(shapes As Object) As List(Of IShapeNode)
        Dim nodes As New List(Of IShapeNode)()
        If shapes Is Nothing Then Return nodes
        Dim count As Integer
        Try
            count = CInt(shapes.Count)
        Catch
            Return nodes
        End Try
        For i As Integer = 1 To count
            Try
                Dim shape As Object = shapes(i)
                If shape IsNot Nothing Then nodes.Add(New ShapeNodeAdapter(shape, Nothing, 1))
            Catch
                ' Shape yang tidak dapat diambil tetap dihitung sebagai kegagalan
                ' baca oleh pemanggil lewat hasil pemindaian (lihat ShapeScanner).
            End Try
        Next
        Return nodes
    End Function

    Public ReadOnly Property WrappedShape As Object Implements IShapeNode.WrappedShape
        Get
            Return _shape
        End Get
    End Property

    Public ReadOnly Property SelectionTarget As Object Implements IShapeNode.SelectionTarget
        Get
            Return _selectionTarget
        End Get
    End Property

    Public ReadOnly Property ChildrenError As String Implements IShapeNode.ChildrenError
        Get
            Dim unused As IList(Of IShapeNode) = GetChildren()
            Return _childrenError
        End Get
    End Property

    Public ReadOnly Property PowerClipError As String Implements IShapeNode.PowerClipError
        Get
            Dim unused As IList(Of IShapeNode) = GetPowerClipChildren()
            Return _powerClipError
        End Get
    End Property

    Public ReadOnly Property IsPowerClipContainer As Boolean Implements IShapeNode.IsPowerClipContainer
        Get
            Dim unused As IList(Of IShapeNode) = GetPowerClipChildren()
            Return _isPowerClip
        End Get
    End Property

    ''' <summary>Baca fakta objek sekali per node (di-cache).</summary>
    Public Function GetFacts() As ShapeFacts Implements IShapeNode.GetFacts
        If _facts IsNot Nothing Then Return _facts
        Dim facts As New ShapeFacts()

        ' --- tipe shape ------------------------------------------------------
        Try
            facts.TypeCode = CInt(_shape.Type)
        Catch ex As Exception
            facts.TypeCode = ShapeKind.Unreadable
            facts.TypeError = Describe(ex)
        End Try

        ' --- nama & layer ----------------------------------------------------
        Try
            facts.ShapeName = CStr(_shape.Name)
        Catch
            facts.ShapeName = ""
        End Try
        Try
            facts.LayerName = CStr(_shape.Layer.Name)
        Catch
            facts.LayerName = ""
        End Try
        Try
            facts.LayerEditable = CBool(_shape.Layer.Editable)
            facts.LayerEditableKnown = True
        Catch
            facts.LayerEditableKnown = False
        End Try

        ' --- status kurva (hanya untuk objek kurva) --------------------------
        If facts.TypeCode = ShapeKind.CurveShape Then
            Try
                facts.CurveClosed = CBool(_shape.Curve.Closed)
                facts.CurveClosedKnown = True
            Catch ex As Exception
                facts.CurveClosedKnown = False
                facts.CurveError = Describe(ex)
            End Try
        End If

        ' --- jenis teks (artistik / paragraf) -------------------------------
        If facts.TypeCode = ShapeKind.TextShape Then
            Try
                facts.TextKind = NormalizeTextKind(CInt(_shape.Text.Type))
            Catch ex As Exception
                facts.TextKind = ShapeKind.TextUnknown
                facts.TextKindError = Describe(ex)
            End Try
        End If

        _facts = facts
        Return facts
    End Function

    ''' <summary>cdrTextType: 1 = artistik, 2 = paragraf. Nilai lain = tidak diketahui.</summary>
    Private Shared Function NormalizeTextKind(value As Integer) As Integer
        If value = 1 Then Return ShapeKind.TextArtistic
        If value = 2 Then Return ShapeKind.TextParagraph
        Return ShapeKind.TextUnknown
    End Function

    ''' <summary>Anak grup. Untuk grup kosong / bukan grup menghasilkan daftar kosong.</summary>
    Public Function GetChildren() As IList(Of IShapeNode) Implements IShapeNode.GetChildren
        If _childrenLoaded Then Return _children
        _childrenLoaded = True
        Dim nodes As New List(Of IShapeNode)()
        _children = nodes

        Dim facts As ShapeFacts = GetFacts()
        If Not facts.TypeReadable OrElse facts.TypeCode <> ShapeKind.GroupShape Then Return nodes

        If _depthHint > MaxDepth Then
            _childrenError = "batas kedalaman " & MaxDepth.ToString() & " terlampaui"
            Return nodes
        End If

        Try
            Dim children As Object = _shape.Shapes
            Dim count As Integer = CInt(children.Count)
            For i As Integer = 1 To count
                Dim child As Object = children(i)
                If child IsNot Nothing Then
                    nodes.Add(New ShapeNodeAdapter(child, Nothing, _depthHint + 1))
                End If
            Next
        Catch ex As Exception
            _childrenError = Describe(ex)
        End Try

        Return nodes
    End Function

    ''' <summary>
    ''' Isi PowerClip. PENTING: pada CorelDRAW, shape.Shapes sebuah container
    ''' PowerClip hanya berisi bingkai container, BUKAN isinya. Isi diambil
    ''' lewat shape.PowerClip. Beberapa versi/API mengeksposnya dengan nama
    ''' berbeda, jadi dicoba berurutan dan kegagalannya dicatat.
    ''' </summary>
    Public Function GetPowerClipChildren() As IList(Of IShapeNode) Implements IShapeNode.GetPowerClipChildren
        If _powerClipLoaded Then Return _powerClipChildren
        _powerClipLoaded = True
        Dim nodes As New List(Of IShapeNode)()
        _powerClipChildren = nodes

        Dim facts As ShapeFacts = GetFacts()
        If Not facts.TypeReadable OrElse facts.TypeCode <> ShapeKind.GroupShape Then Return nodes
        If _depthHint > MaxDepth Then
            _powerClipError = "batas kedalaman " & MaxDepth.ToString() & " terlampaui"
            Return nodes
        End If

        Dim powerClip As Object
        Try
            powerClip = _shape.PowerClip
        Catch ex As Exception
            ' Objek bukan PowerClip: bukan kegagalan (normal), jangan dicatat.
            _powerClipError = ""
            Return nodes
        End Try
        If powerClip Is Nothing Then Return nodes

        _isPowerClip = True

        ' Urutan strategi pembacaan isi PowerClip (late bound, tanpa asumsi).
        Dim attempts As New List(Of String)(New String() {
            "PowerClip.Shapes", "PowerClip.Contents.Shapes", "PowerClip.Contents"
        })
        Dim errors As New List(Of String)()

        For Each attempt As String In attempts
            Try
                Dim collection As Object = Nothing
                Select Case attempt
                    Case "PowerClip.Shapes"
                        collection = powerClip.Shapes
                    Case "PowerClip.Contents.Shapes"
                        Dim contents As Object = powerClip.Contents
                        If contents IsNot Nothing Then collection = contents.Shapes
                    Case "PowerClip.Contents"
                        collection = powerClip.Contents
                End Select

                If collection Is Nothing Then Continue For
                Dim built As List(Of IShapeNode) = WrapChildren(collection, _selectionTarget)
                If built.Count > 0 Then
                    nodes.AddRange(built)
                    _powerClipError = ""
                    Return nodes
                End If
            Catch ex As Exception
                errors.Add(attempt & " -> " & Describe(ex))
            End Try
        Next

        ' Container terdeteksi tetapi isinya tidak terbaca: laporkan, jangan
        ' diam-diam menganggap dokumen tidak berisi objek.
        If nodes.Count = 0 Then
            _powerClipError = "isi PowerClip tidak dapat dibaca" &
                              If(errors.Count = 0, "", " (" & String.Join("; ", errors.ToArray()) & ")")
        End If
        Return nodes
    End Function

    ''' <summary>
    ''' Bungkus isi PowerClip. SelectionTarget diteruskan agar pemilihan memakai
    ''' CONTAINER (isi PowerClip tidak dapat dipilih sendiri).
    ''' </summary>
    Private Function WrapChildren(collection As Object, target As Object) As List(Of IShapeNode)
        Dim nodes As New List(Of IShapeNode)()
        Dim count As Integer = CInt(collection.Count)
        For i As Integer = 1 To count
            Dim child As Object = Nothing
            Try
                child = collection(i)
            Catch
                Continue For
            End Try
            If child IsNot Nothing Then
                nodes.Add(New ShapeNodeAdapter(child, If(target, child), _depthHint + 1))
            End If
        Next
        Return nodes
    End Function

    Private Shared Function Describe(ex As Exception) As String
        If ex Is Nothing Then Return "kesalahan tidak diketahui"
        Return ex.GetType().Name & ": " & ex.Message
    End Function
End Class
