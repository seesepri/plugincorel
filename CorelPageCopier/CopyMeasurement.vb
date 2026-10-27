Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyMeasurement.vb
'  Pengukuran objek sumber & objek hasil paste (late bound, TANPA WinForms).
'
'  Kenapa terpisah dari Program.vb:
'    - Bisa diuji unit dengan objek .NET pengganti COM (lihat tests/).
'    - Memusatkan satu perbaikan penting (Temuan A):
'        kode lama mengukur `layer.Shapes(1 .. added)`, yaitu objek TERBAWAH
'        pada layer, bukan objek hasil paste. Objek hasil paste selalu berada
'        di ATAS tumpukan z-order, jadi indeks yang benar adalah
'        (Count - added + 1) .. Count.
'
'  Catatan kompatibilitas:
'    - Semua properti dibaca lewat late binding (Option Strict Off) sama seperti
'      kode lama, sehingga tidak butuh referensi interop CorelDRAW.
'    - Tidak ada objek yang dihapus/diubah di sini; hanya membaca properti.
' ============================================================================

Imports System

''' <summary>Hasil pengukuran satu kelompok objek (kotak + sidik jari struktur).</summary>
Public Class MeasuredShapeSet
    Public Bounds As BoundingBox = BoundingBox.Empty()
    Public Signature As New ShapeSignature()
    ''' <summary>Asal data: "z-order atas", "selection", dsb (untuk transparansi laporan).</summary>
    Public Source As String = ""
    ''' <summary>Alasan bila pengukuran tidak lengkap.</summary>
    Public Reason As String = ""

    Public ReadOnly Property HasGeometry As Boolean
        Get
            Return Bounds.IsValid
        End Get
    End Property

    Public ReadOnly Property HasStructure As Boolean
        Get
            Return Signature.IsComparable
        End Get
    End Property

    ''' <summary>Salinan dengan batas dikalikan faktor konversi satuan (menjadi mm).</summary>
    Public Function Scaled(factor As Double) As MeasuredShapeSet
        Dim copy As New MeasuredShapeSet()
        copy.Signature = Signature
        copy.Source = Source
        copy.Reason = Reason
        If Bounds.IsValid AndAlso factor > 0.0 Then copy.Bounds = Bounds.Scale(factor)
        Return copy
    End Function
End Class

''' <summary>Helper pengukuran shape lewat late binding.</summary>
Public Class CopyMeasurement

    ''' <summary>Baca jumlah item sebuah koleksi/range; -1 bila tidak terbaca.</summary>
    Public Shared Function SafeCount(collection As Object) As Integer
        If collection Is Nothing Then Return -1
        Try
            Return CInt(collection.Count)
        Catch
            Return -1
        End Try
    End Function

    ''' <summary>Baca kode tipe shape; -1 bila tidak terbaca.</summary>
    Public Shared Function SafeTypeCode(shape As Object) As Integer
        If shape Is Nothing Then Return -1
        Try
            Return CInt(shape.Type)
        Catch
            Return -1
        End Try
    End Function

    ''' <summary>Ambil item ke-index (1-based) dari sebuah koleksi.</summary>
    Public Shared Function SafeItem(collection As Object, oneBasedIndex As Integer) As Object
        If collection Is Nothing Then Return Nothing
        Try
            Return collection(oneBasedIndex)
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Ukur potongan koleksi shape (indeks 1-based, inklusif).
    ''' Kotak gabungan dihitung dari CenterX/CenterY/SizeWidth/SizeHeight tiap objek.
    ''' Bila salah satu objek tidak dapat diukur, HasGeometry = False (tidak menebak).
    ''' </summary>
    Public Shared Function MeasureSlice(shapes As Object, firstIndex As Integer, count As Integer,
                                        label As String) As MeasuredShapeSet
        Dim result As New MeasuredShapeSet()
        result.Source = If(label, "")

        If count <= 0 Then
            result.Reason = "tidak ada objek untuk diukur"
            Return result
        End If
        If firstIndex < 1 Then
            result.Reason = "indeks awal tidak valid"
            Return result
        End If

        Dim box As BoundingBox = BoundingBox.Empty()
        Dim measured As Integer = 0
        Dim geometryError As String = ""

        For i As Integer = firstIndex To firstIndex + count - 1
            Dim shape As Object = SafeItem(shapes, i)
            If shape Is Nothing Then
                geometryError = "objek ke-" & i & " tidak dapat diakses"
                Exit For
            End If

            Dim typeCode As Integer = SafeTypeCode(shape)
            If typeCode < 0 Then
                result.Signature.MarkUnreadable()
            Else
                result.Signature.AddShapeType(typeCode)
            End If

            Try
                Dim cx As Double = CDbl(shape.CenterX)
                Dim cy As Double = CDbl(shape.CenterY)
                Dim w As Double = CDbl(shape.SizeWidth)
                Dim h As Double = CDbl(shape.SizeHeight)
                box = box.Union(BoundingBox.FromCenter(cx, cy, w, h))
                measured += 1
            Catch ex As Exception
                geometryError = ex.Message
                Exit For
            End Try
        Next

        If measured = count AndAlso box.IsValid Then
            result.Bounds = box
        Else
            result.Bounds = BoundingBox.Empty()
            result.Reason = "geometri tidak dapat diukur" &
                            If(String.IsNullOrEmpty(geometryError), "", ": " & geometryError)
        End If

        Return result
    End Function

    ''' <summary>
    ''' Ukur N objek TERATAS pada koleksi layer (objek hasil paste selalu di atas).
    ''' Mengembalikan hasil dengan HasGeometry = False bila koleksi lebih pendek
    ''' dari jumlah yang diminta (menandakan paste tidak sesuai harapan).
    ''' </summary>
    Public Shared Function MeasureTopShapes(shapes As Object, addedCount As Integer) As MeasuredShapeSet
        Dim result As New MeasuredShapeSet()
        result.Source = "z-order atas (layer)"
        If addedCount <= 0 Then
            result.Reason = "jumlah objek hasil paste tidak diketahui"
            Return result
        End If

        Dim total As Integer = SafeCount(shapes)
        If total < addedCount Then
            result.Reason = "koleksi layer hanya berisi " & total & " objek, kurang dari " & addedCount & " objek baru"
            Return result
        End If

        Return MeasureSlice(shapes, total - addedCount + 1, addedCount, result.Source)
    End Function

    ''' <summary>
    ''' Ukur sekumpulan objek pilihan (ShapeRange / Shape). Berguna untuk objek
    ''' sumber sebelum di-copy. Tidak mengubah pilihan maupun objek.
    ''' </summary>
    Public Shared Function MeasureSelection(rangeOrShape As Object) As MeasuredShapeSet
        Dim result As New MeasuredShapeSet()
        result.Source = "selection (objek sumber)"

        If rangeOrShape Is Nothing Then
            result.Reason = "selection kosong"
            Return result
        End If

        Dim total As Integer = SafeCount(rangeOrShape)
        If total < 0 Then
            ' Bukan range (Shape tunggal tanpa Count): ukur sebagai satu objek.
            Return MeasureSingleShape(rangeOrShape, result.Source)
        End If
        If total = 0 Then
            result.Reason = "selection kosong"
            Return result
        End If

        Return MeasureSlice(rangeOrShape, 1, total, result.Source)
    End Function

    ''' <summary>Ukur satu objek saja.</summary>
    Public Shared Function MeasureSingleShape(shape As Object, label As String) As MeasuredShapeSet
        Dim result As New MeasuredShapeSet()
        result.Source = If(label, "objek tunggal")
        If shape Is Nothing Then
            result.Reason = "objek tidak tersedia"
            Return result
        End If

        Dim typeCode As Integer = SafeTypeCode(shape)
        If typeCode < 0 Then
            result.Signature.MarkUnreadable()
        Else
            result.Signature.AddShapeType(typeCode)
        End If

        Try
            result.Bounds = BoundingBox.FromCenter(CDbl(shape.CenterX), CDbl(shape.CenterY),
                                                   CDbl(shape.SizeWidth), CDbl(shape.SizeHeight))
        Catch ex As Exception
            result.Bounds = BoundingBox.Empty()
            result.Reason = "geometri tidak dapat diukur: " & ex.Message
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Ukur objek hasil paste. Urutan strategi (didokumentasikan):
    '''   1. N objek teratas pada layer aktif (deterministik, tidak bergantung
    '''      pada perilaku seleksi setelah paste).
    '''   2. Bila geometri z-order gagal diukur, pakai ActiveSelectionRange HANYA
    '''      jika jumlahnya sama dengan jumlah objek yang baru ditambahkan.
    ''' </summary>
    Public Shared Function MeasurePasted(selectionRange As Object, layerShapes As Object,
                                         addedCount As Integer) As MeasuredShapeSet
        Dim primary As MeasuredShapeSet = MeasureTopShapes(layerShapes, addedCount)
        If primary.HasGeometry Then Return primary

        If addedCount > 0 AndAlso selectionRange IsNot Nothing AndAlso SafeCount(selectionRange) = addedCount Then
            Dim fallback As MeasuredShapeSet = MeasureSelection(selectionRange)
            fallback.Source = "selection setelah paste"
            If fallback.HasGeometry Then Return fallback
            If String.IsNullOrEmpty(primary.Reason) Then primary.Reason = fallback.Reason
        End If

        Return primary
    End Function
End Class
