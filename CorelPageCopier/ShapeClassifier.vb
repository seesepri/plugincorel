Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ShapeClassifier.vb
'  Klasifikasi objek CorelDRAW untuk fitur "Auto Select".
'
'  Kenapa file ini terpisah dari Program.vb:
'    - Audit Auto Select menemukan klasifikasi lama menempel di UI dan memakai
'      aturan "vektor = apa pun yang bukan teks/bitmap/grup", sehingga garis
'      bantu (guideline), objek OLE, dan bentuk seleksi ikut terpilih.
'    - Klasifikasi sekarang murni: hanya menerima "fakta" objek (kode tipe,
'      status kurva, jenis teks) sehingga bisa diuji unit tanpa CorelDRAW.
'    - Setiap kegagalan baca (COM) menghasilkan catatan (Note), bukan
'      penelanan senyap.
'
'  DEFINISI KATEGORI (dokumen ini adalah sumber kebenaran; UI memakai label
'  yang sama lewat CategoryLabel/CategoryHelp):
'
'   0 TEKS
'     - Artistic text DAN paragraph text (keduanya shape.Type = cdrTextShape).
'     - Sub-jenis dibaca dari Text.Type (1 = artistik, 2 = paragraf). Bila
'       sub-jenis tidak terbaca, objek TETAP masuk kategori teks (tipe shape
'       sudah pasti teks) dan diberi catatan.
'
'   1 BITMAP
'     - shape.Type = cdrBitmapShape. Termasuk bitmap yang berada DI DALAM grup
'       (grup ditelusuri) dan di dalam PowerClip (isi container dipindai,
'       lihat ShapeScanner).
'
'   2 KURVA TERBUKA
'     - shape.Type = cdrCurveShape DAN Curve.Closed = False (status kurva
'       aktual dari COM, bukan tebakan dari bounding box).
'     - Bila Curve.Closed tidak dapat dibaca: objek TIDAK dimasukkan dan
'       dicatat sebagai pelanggaran keterbacaan (tidak menebak).
'
'   3 VEKTOR
'     Mencakup: rectangle (1), ellipse (2), polygon (4), kurva (3) baik
'     terbuka maupun tertutup, dan objek vektor khusus (mis. metafile/EMF-WMF,
'     connector) yang kodenya tidak termasuk daftar pengecualian.
'     TIDAK mencakup: noShape (0), bitmap (5), teks (6), grup (7),
'     guideline/garis bantu (8), objek OLE (9), bentuk seleksi (10).
'     Catatan: kurva terbuka memang muncul di kategori 2 dan 3 — kategori 2
'     adalah subset yang lebih spesifik.
'
'   4 SEMUA OBJEK
'     - Setiap objek level atas yang bukan watermark ikut terpilih.
'     - Grup dipertahankan sebagai SATU kesatuan (struktur grup tidak dipecah).
'     - Grup yang di dalamnya memuat objek watermark dilewati seluruhnya,
'       karena memilih sebagian isi grup akan memecah struktur grup.
'
'  WATERMARK (semua kategori)
'     - Objek dengan nama berawalan "WM_" atau berada pada layer "WATERMARK"
'       tidak pernah masuk seleksi.
'     - Di dalam grup, hanya anak yang benar-benar watermark yang dilewati;
'       saudaranya yang bukan watermark tetap bisa terpilih.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization

''' <summary>Kategori objek pada UI Auto Select. Urutan = urutan ComboBox.</summary>
Public Enum ShapeCategory
    Text = 0
    Bitmap = 1
    OpenCurve = 2
    Vector = 3
    AllObjects = 4
End Enum

''' <summary>Kode cdrShapeType yang dipakai klasifikasi (nilai stabil lintas versi).</summary>
Public Class ShapeKind
    Public Const NoShape As Integer = 0
    Public Const RectangleShape As Integer = 1
    Public Const EllipseShape As Integer = 2
    Public Const CurveShape As Integer = 3
    Public Const PolygonShape As Integer = 4
    Public Const BitmapShape As Integer = 5
    Public Const TextShape As Integer = 6
    Public Const GroupShape As Integer = 7
    Public Const GuidelineShape As Integer = 8
    Public Const OleObjectShape As Integer = 9
    Public Const SelectionShape As Integer = 10

    ''' <summary>Nilai "tidak terbaca" (bukan kode CorelDRAW).</summary>
    Public Const Unreadable As Integer = -1

    ''' <summary>Jenis teks (cdrTextType). 0 = tidak diketahui.</summary>
    Public Const TextUnknown As Integer = 0
    Public Const TextArtistic As Integer = 1
    Public Const TextParagraph As Integer = 2

    ''' <summary>Label ramah pengguna untuk satu kode tipe.</summary>
    Public Shared Function KindLabel(typeCode As Integer) As String
        Select Case typeCode
            Case RectangleShape
                Return "Rectangle"
            Case EllipseShape
                Return "Ellipse"
            Case CurveShape
                Return "Kurva"
            Case PolygonShape
                Return "Polygon"
            Case BitmapShape
                Return "Bitmap"
            Case TextShape
                Return "Teks"
            Case GroupShape
                Return "Grup"
            Case GuidelineShape
                Return "Garis bantu (guideline)"
            Case OleObjectShape
                Return "Objek OLE"
            Case SelectionShape
                Return "Bentuk seleksi"
            Case NoShape
                Return "Tanpa objek"
            Case Else
                Return "Tipe tidak dikenal (" & typeCode.ToString(CultureInfo.InvariantCulture) & ")"
        End Select
    End Function

    ''' <summary>True bila kode tipe ini objek desain yang bisa masuk kategori Vektor.</summary>
    Public Shared Function IsVectorKind(typeCode As Integer) As Boolean
        Select Case typeCode
            Case RectangleShape, EllipseShape, CurveShape, PolygonShape
                Return True
            Case NoShape, BitmapShape, TextShape, GroupShape, GuidelineShape, OleObjectShape, SelectionShape
                Return False
            Case Else
                ' Kode lain (mis. metafile/connector) diperlakukan sebagai objek
                ' vektor khusus; pemanggil menambahkan catatan agar diverifikasi.
                Return typeCode > 0
        End Select
    End Function

    ''' <summary>True bila kode tipe ini dikenal sebagai objek NON-vektor yang harus dikecualikan.</summary>
    Public Shared Function IsKnownNonVector(typeCode As Integer) As Boolean
        Select Case typeCode
            Case NoShape, BitmapShape, TextShape, GroupShape, GuidelineShape, OleObjectShape, SelectionShape
                Return True
            Case Else
                Return False
        End Select
    End Function
End Class

''' <summary>
''' Fakta mentah satu objek hasil pembacaan COM (atau data uji).
''' Semua properti opsional: nilai "tidak terbaca" disertai pesan error.
''' </summary>
Public Class ShapeFacts
    Public TypeCode As Integer = ShapeKind.Unreadable
    Public TypeError As String = ""
    Public ShapeName As String = ""
    Public LayerName As String = ""
    Public LayerEditable As Boolean = True
    Public LayerEditableKnown As Boolean = True
    Public CurveClosed As Boolean = False
    Public CurveClosedKnown As Boolean = False
    Public CurveError As String = ""
    Public TextKind As Integer = ShapeKind.TextUnknown
    Public TextKindError As String = ""

    Public ReadOnly Property TypeReadable As Boolean
        Get
            Return TypeCode >= 0
        End Get
    End Property

    Public ReadOnly Property IsWatermark As Boolean
        Get
            Return WatermarkLayout.IsWatermarkIdentity(ShapeName, LayerName)
        End Get
    End Property

    Public Shared Function FromValues(typeCode As Integer,
                                      Optional shapeName As String = "",
                                      Optional layerName As String = "") As ShapeFacts
        Dim facts As New ShapeFacts()
        facts.TypeCode = typeCode
        facts.ShapeName = If(shapeName, "")
        facts.LayerName = If(layerName, "")
        Return facts
    End Function
End Class

''' <summary>Hasil klasifikasi satu objek untuk satu kategori.</summary>
Public Class ClassificationResult
    Public Matches As Boolean = False
    ''' <summary>Label jenis objek untuk laporan ("Teks artistik", "Bitmap", …).</summary>
    Public KindLabel As String = ""
    ''' <summary>Catatan non-fatal: keterbacaan, tipe tak dikenal, layer terkunci, dsb.</summary>
    Public Note As String = ""

    Public Shared Function No(note As String) As ClassificationResult
        Dim result As New ClassificationResult()
        result.Matches = False
        result.Note = If(note, "")
        Return result
    End Function

    Public Shared Function Yes(kindLabel As String, Optional note As String = "") As ClassificationResult
        Dim result As New ClassificationResult()
        result.Matches = True
        result.KindLabel = If(kindLabel, "")
        result.Note = If(note, "")
        Return result
    End Function
End Class

''' <summary>Aturan klasifikasi murni (tanpa COM, tanpa WinForms) — bisa diuji unit.</summary>
Public Class ShapeClassifier

    ''' <summary>Label ComboBox — urutan HARUS sama dengan ShapeCategory.</summary>
    Public Shared Function CategoryLabels() As String()
        Return New String() {
            "Teks (Artistik + Paragraf)",
            "Gambar / Bitmap",
            "Garis / Kurva terbuka",
            "Objek vektor (rectangle/ellipse/polygon/kurva)",
            "Semua objek"
        }
    End Function

    ''' <summary>Label singkat satu kategori (untuk laporan/log).</summary>
    Public Shared Function CategoryLabel(category As ShapeCategory) As String
        Dim labels As String() = CategoryLabels()
        Dim index As Integer = CInt(category)
        If index < 0 OrElse index >= labels.Length Then Return "Tidak dikenal"
        Return labels(index)
    End Function

    ''' <summary>
    ''' Penjelasan definisi kategori yang ditampilkan di UI/laporan, supaya
    ''' cakupan "Objek vektor" tidak lagi ambigu.
    ''' </summary>
    Public Shared Function CategoryHelp(category As ShapeCategory) As String
        Select Case category
            Case ShapeCategory.Text
                Return "Artistic text + paragraph text (shape.Type = text)."
            Case ShapeCategory.Bitmap
                Return "Objek bitmap, termasuk bitmap di dalam grup dan di dalam PowerClip."
            Case ShapeCategory.OpenCurve
                Return "Kurva dengan Curve.Closed = False (status kurva aktual). " &
                       "Status yang tidak terbaca tidak dianggap kurva terbuka."
            Case ShapeCategory.Vector
                Return "Rectangle, ellipse, polygon, kurva (terbuka/tertutup), dan objek vektor khusus. " &
                       "Tidak termasuk teks, bitmap, grup, garis bantu, objek OLE, dan bentuk seleksi."
            Case Else
                Return "Semua objek level atas; grup dipilih sebagai satu kesatuan. Watermark selalu dikecualikan."
        End Select
    End Function

    ''' <summary>Kategori berdasarkan indeks ComboBox (dijaga agar aman bila indeks aneh).</summary>
    Public Shared Function CategoryFromIndex(index As Integer) As ShapeCategory
        If index < 0 OrElse index > CInt(ShapeCategory.AllObjects) Then Return ShapeCategory.Text
        Return CType(index, ShapeCategory)
    End Function

    ''' <summary>
    ''' Klasifikasikan satu objek. Tidak pernah melempar exception; semua
    ''' kegagalan keterbacaan dikembalikan sebagai Note.
    ''' </summary>
    Public Shared Function Classify(facts As ShapeFacts, category As ShapeCategory) As ClassificationResult
        If facts Is Nothing Then Return ClassificationResult.No("objek tidak tersedia")

        ' --- Watermark selalu dikecualikan (semua kategori) -------------------
        If facts.IsWatermark Then
            Return ClassificationResult.No("watermark (nama WM_/layer WATERMARK) — dikecualikan")
        End If

        ' --- Tipe tidak terbaca: jangan menebak ------------------------------
        If Not facts.TypeReadable Then
            Return ClassificationResult.No("tipe objek tidak dapat dibaca" &
                                          If(String.IsNullOrEmpty(facts.TypeError), "", ": " & facts.TypeError))
        End If

        ' --- Semua objek: level atas dipilih apa adanya (grup = satu unit) ----
        If category = ShapeCategory.AllObjects Then
            Return ClassificationResult.Yes(DescribeKind(facts), LayerNote(facts))
        End If

        ' --- Grup: bukan anggota kategori spesifik (isinya yang dipindai) ----
        If facts.TypeCode = ShapeKind.GroupShape Then
            Return ClassificationResult.No("")
        End If

        Select Case category
            Case ShapeCategory.Text
                Return ClassifyText(facts)
            Case ShapeCategory.Bitmap
                Return ClassifyBitmap(facts)
            Case ShapeCategory.OpenCurve
                Return ClassifyOpenCurve(facts)
            Case ShapeCategory.Vector
                Return ClassifyVector(facts)
            Case Else
                Return ClassificationResult.No("kategori tidak dikenal")
        End Select
    End Function

    ' --- per kategori -------------------------------------------------------

    Private Shared Function ClassifyText(facts As ShapeFacts) As ClassificationResult
        If facts.TypeCode <> ShapeKind.TextShape Then Return ClassificationResult.No("")

        Dim note As String = LayerNote(facts)
        Select Case facts.TextKind
            Case ShapeKind.TextArtistic
                Return ClassificationResult.Yes("Teks artistik", note)
            Case ShapeKind.TextParagraph
                Return ClassificationResult.Yes("Teks paragraf", note)
            Case Else
                Dim detail As String = "jenis teks tidak terbaca"
                If Not String.IsNullOrEmpty(facts.TextKindError) Then detail &= ": " & facts.TextKindError
                Return ClassificationResult.Yes("Teks (jenis tidak diketahui)", Combine(note, detail))
        End Select
    End Function

    Private Shared Function ClassifyBitmap(facts As ShapeFacts) As ClassificationResult
        If facts.TypeCode <> ShapeKind.BitmapShape Then Return ClassificationResult.No("")
        Return ClassificationResult.Yes("Bitmap", LayerNote(facts))
    End Function

    Private Shared Function ClassifyOpenCurve(facts As ShapeFacts) As ClassificationResult
        If facts.TypeCode <> ShapeKind.CurveShape Then Return ClassificationResult.No("")
        If Not facts.CurveClosedKnown Then
            ' Status kurva aktual tidak diketahui -> JANGAN menebak dari bbox/tipe.
            Return ClassificationResult.No(
                "status Curve.Closed tidak dapat dibaca" &
                If(String.IsNullOrEmpty(facts.CurveError), "", ": " & facts.CurveError) &
                " — objek tidak dimasukkan ke kategori kurva terbuka")
        End If
        If facts.CurveClosed Then Return ClassificationResult.No("")
        Return ClassificationResult.Yes("Kurva terbuka", LayerNote(facts))
    End Function

    Private Shared Function ClassifyVector(facts As ShapeFacts) As ClassificationResult
        Dim note As String = LayerNote(facts)
        Select Case facts.TypeCode
            Case ShapeKind.RectangleShape
                Return ClassificationResult.Yes("Rectangle", note)
            Case ShapeKind.EllipseShape
                Return ClassificationResult.Yes("Ellipse", note)
            Case ShapeKind.PolygonShape
                Return ClassificationResult.Yes("Polygon", note)
            Case ShapeKind.CurveShape
                If Not facts.CurveClosedKnown Then
                    ' Kurva tetap objek vektor walau status tertutupnya tak terbaca.
                    Return ClassificationResult.Yes("Kurva (status tertutup tidak terbaca)",
                                                    Combine(note, "status Curve.Closed tidak dapat dibaca"))
                End If
                Return ClassificationResult.Yes(If(facts.CurveClosed, "Kurva tertutup", "Kurva terbuka"), note)
            Case Else
                If ShapeKind.IsKnownNonVector(facts.TypeCode) Then
                    ' teks/bitmap/grup/guideline/OLE/bentuk seleksi: bukan vektor.
                    Return ClassificationResult.No("")
                End If
                If ShapeKind.IsVectorKind(facts.TypeCode) Then
                    Return ClassificationResult.Yes(
                        ShapeKind.KindLabel(facts.TypeCode) & " (vektor khusus)",
                        Combine(note, "kode tipe " & facts.TypeCode.ToString(CultureInfo.InvariantCulture) &
                                      " tidak ada di daftar eksplisit — diperlakukan sebagai vektor khusus, mohon verifikasi manual"))
                End If
                Return ClassificationResult.No("")
        End Select
    End Function

    ' --- helper -------------------------------------------------------------

    ''' <summary>Label jenis objek (dipakai kategori "Semua objek").</summary>
    Public Shared Function DescribeKind(facts As ShapeFacts) As String
        If facts Is Nothing Then Return "objek"
        Select Case facts.TypeCode
            Case ShapeKind.TextShape
                If facts.TextKind = ShapeKind.TextArtistic Then Return "Teks artistik"
                If facts.TextKind = ShapeKind.TextParagraph Then Return "Teks paragraf"
                Return "Teks"
            Case ShapeKind.CurveShape
                If Not facts.CurveClosedKnown Then Return "Kurva"
                Return If(facts.CurveClosed, "Kurva tertutup", "Kurva terbuka")
            Case Else
                Return ShapeKind.KindLabel(facts.TypeCode)
        End Select
    End Function

    ''' <summary>Catatan bila objek berada pada layer terkunci/tidak dapat diedit.</summary>
    Private Shared Function LayerNote(facts As ShapeFacts) As String
        If facts.LayerEditableKnown AndAlso Not facts.LayerEditable Then
            Return "layer """ & facts.LayerName & """ terkunci (objek tidak dapat dipilih)"
        End If
        Return ""
    End Function

    Private Shared Function Combine(left As String, right As String) As String
        If String.IsNullOrEmpty(left) Then Return If(right, "")
        If String.IsNullOrEmpty(right) Then Return left
        Return left & "; " & right
    End Function
End Class
