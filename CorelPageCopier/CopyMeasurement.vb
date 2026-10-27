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
Imports System.Collections.Generic
Imports System.Globalization

''' <summary>
''' Jenis input yang diukur. Dicatat di hasil pengukuran supaya laporan bisa
''' menyebut apa yang sebenarnya diukur (bukan menebak).
''' </summary>
Public Enum MeasurementInputKind
    Empty = 0
    ''' <summary>Objek Shape tunggal: TIDAK memiliki anggota Count.</summary>
    SingleShape = 1
    ''' <summary>ShapeRange (nama tipe mengandung "ShapeRange").</summary>
    ShapeRange = 2
    ''' <summary>Koleksi COM lain (mis. Shapes, Selection) yang punya Count.</summary>
    ShapeCollection = 3
    ''' <summary>Punya Count tetapi gagal dibaca, atau jenis tidak dapat dipastikan.</summary>
    Unknown = 4
End Enum

''' <summary>
''' Status pembacaan Count. Perbedaan MemberMissing vs Failed adalah inti
''' perbaikan di file ini:
'''   - MemberMissing : objek memang TIDAK punya Count (Shape tunggal). Aman
'''                     diukur sebagai satu objek.
'''   - Failed        : punya Count tetapi gagal dibaca (COM mati, RPC sibuk,
'''                     akses ditolak). Jumlah objek TIDAK diketahui, sehingga
'''                     TIDAK boleh diperlakukan sebagai satu Shape.
''' </summary>
Public Enum CountReadStatus
    Ok = 0
    MemberMissing = 1
    Failed = 2
End Enum

''' <summary>Hasil satu percobaan membaca Count (tidak pernah melempar).</summary>
Public Structure CountReadResult
    Public Status As CountReadStatus
    Public Count As Integer
    Public ErrorMessage As String

    Public ReadOnly Property IsOk As Boolean
        Get
            Return Status = CountReadStatus.Ok
        End Get
    End Property

    Public ReadOnly Property IsMemberMissing As Boolean
        Get
            Return Status = CountReadStatus.MemberMissing
        End Get
    End Property

    Public Function Describe() As String
        Select Case Status
            Case CountReadStatus.Ok
                Return "Count = " & Count.ToString(CultureInfo.InvariantCulture)
            Case CountReadStatus.MemberMissing
                Return "objek tidak memiliki anggota Count"
            Case Else
                Return "Count tidak dapat dibaca: " & ErrorMessage
        End Select
    End Function
End Structure

''' <summary>Hasil pengukuran satu kelompok objek (kotak + sidik jari struktur).</summary>
Public Class MeasuredShapeSet
    Public Bounds As BoundingBox = BoundingBox.Empty()
    Public Signature As New ShapeSignature()
    ''' <summary>Asal data: "z-order atas", "selection", dsb (untuk transparansi laporan).</summary>
    Public Source As String = ""
    ''' <summary>Alasan bila pengukuran tidak lengkap.</summary>
    Public Reason As String = ""
    ''' <summary>Jenis input yang diukur (Shape tunggal / ShapeRange / koleksi / tidak dikenal).</summary>
    Public InputKind As MeasurementInputKind = MeasurementInputKind.Unknown

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
        copy.InputKind = InputKind
        If Bounds.IsValid AndAlso factor > 0.0 Then copy.Bounds = Bounds.Scale(factor)
        Return copy
    End Function
End Class

''' <summary>Helper pengukuran shape lewat late binding.</summary>
Public Class CopyMeasurement

    ''' <summary>
    ''' Baca jumlah item sebuah koleksi/range; -1 bila tidak terbaca.
    ''' API lama dipertahankan, tetapi TIDAK membedakan "tidak punya Count"
    ''' dari "Count gagal dibaca" — pakai TryReadCount untuk keputusan.
    ''' </summary>
    Public Shared Function SafeCount(collection As Object) As Integer
        Dim read As CountReadResult = TryReadCount(collection)
        If read.IsOk Then Return read.Count
        Return -1
    End Function

    ''' <summary>
    ''' Baca Count sambil membedakan tiga kemungkinan:
    '''   Ok            -> jumlah objek diketahui,
    '''   MemberMissing -> objek tidak punya Count (Shape tunggal),
    '''   Failed        -> Count ada tetapi gagal dibaca (jumlah TIDAK diketahui).
    ''' Tidak pernah melempar.
    ''' </summary>
    Public Shared Function TryReadCount(value As Object) As CountReadResult
        Dim result As New CountReadResult()
        result.Count = -1
        If value Is Nothing Then
            result.Status = CountReadStatus.Failed
            result.ErrorMessage = "objek null"
            Return result
        End If

        Try
            result.Count = CInt(value.Count)
            result.Status = CountReadStatus.Ok
            Return result
        Catch ex As Exception
            result.Count = -1
            result.ErrorMessage = DescribeError(ex)
            If IsMemberMissing(ex) Then
                result.Status = CountReadStatus.MemberMissing
            Else
                result.Status = CountReadStatus.Failed
            End If
            Return result
        End Try
    End Function

    ''' <summary>
    ''' Tentukan jenis input pengukuran. Hasil Unknown berarti jumlah/jenis objek
    ''' tidak dapat dipastikan -> pemanggil wajib memperlakukan hasil sebagai
    ''' TIDAK terverifikasi (Partial), bukan sebagai satu Shape.
    '''
    ''' Penting (review kedua): "anggota Count tidak ada" HANYA boleh dianggap
    ''' Shape tunggal bila objeknya memang terbukti berbentuk Shape, yaitu kode
    ''' tipenya (Shape.Type) terbaca. Objek tanpa Count dan tanpa tipe terbaca
    ''' bisa saja koleksi yang rusak atau objek asing; mengukurnya sebagai satu
    ''' Shape akan menghasilkan geometri yang tampak cocok padahal jumlah objek
    ''' tidak diketahui.
    ''' </summary>
    Public Shared Function ClassifyInput(value As Object) As MeasurementInputKind
        If value Is Nothing Then Return MeasurementInputKind.Empty

        Dim read As CountReadResult = TryReadCount(value)
        Select Case read.Status
            Case CountReadStatus.Ok
                Dim typeLabel As String = SafeTypeName(value)
                If typeLabel.IndexOf("ShapeRange", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return MeasurementInputKind.ShapeRange
                End If
                Return MeasurementInputKind.ShapeCollection
            Case CountReadStatus.MemberMissing
                ' Tidak punya Count: baru dianggap Shape tunggal bila bentuk
                ' objeknya terbukti (kode tipe terbaca).
                If LooksLikeShape(value) Then Return MeasurementInputKind.SingleShape
                Return MeasurementInputKind.Unknown
            Case Else
                Return MeasurementInputKind.Unknown
        End Select
    End Function

    ''' <summary>
    ''' True bila objek memberi bukti bahwa dirinya Shape: kode tipe terbaca.
    ''' Dipakai hanya untuk memverifikasi kandidat "Shape tunggal"; objek yang
    ''' gagal membaca tipe tidak pernah dipaksa menjadi Shape.
    ''' </summary>
    Public Shared Function LooksLikeShape(value As Object) As Boolean
        If value Is Nothing Then Return False
        Return SafeTypeCode(value) >= 0
    End Function

    ''' <summary>Label jenis input untuk laporan (tanpa menebak yang tidak diketahui).</summary>
    Public Shared Function DescribeInputKind(kind As MeasurementInputKind) As String
        Select Case kind
            Case MeasurementInputKind.Empty
                Return "kosong"
            Case MeasurementInputKind.SingleShape
                Return "Shape tunggal (tanpa Count)"
            Case MeasurementInputKind.ShapeRange
                Return "ShapeRange"
            Case MeasurementInputKind.ShapeCollection
                Return "koleksi objek"
            Case Else
                Return "jenis tidak dikenali"
        End Select
    End Function

    ''' <summary>Nama tipe objek (aman; "" bila tidak terbaca).</summary>
    Public Shared Function SafeTypeName(value As Object) As String
        If value Is Nothing Then Return ""
        Try
            Return Microsoft.VisualBasic.Information.TypeName(value)
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' True bila exception berarti "anggota tidak ditemukan" (member tidak ada),
    ''' bukan kegagalan COM/RPC. Dipakai untuk membedakan Shape tunggal dari
    ''' koleksi yang Count-nya gagal dibaca.
    '''
    ''' NotSupportedException TIDAK dihitung sebagai anggota hilang (review
    ''' kedua): anggota yang ada tetapi "tidak didukung" tidak membuktikan
    ''' bahwa objek tidak punya Count. Pembedaan ini hanya salah satu syarat;
    ''' lihat juga LooksLikeShape di ClassifyInput.
    ''' </summary>
    Public Shared Function IsMemberMissing(ex As Exception) As Boolean
        Dim current As Exception = ex
        Do While current IsNot Nothing
            If TypeOf current Is MissingMemberException Then
                Return True
            End If
            ' DISP_E_MEMBERNOTFOUND / DISP_E_UNKNOWNNAME: anggota tidak ada.
            If current.HResult = &H80020003 OrElse current.HResult = &H80020006 Then
                Return True
            End If
            current = current.InnerException
        Loop
        Return False
    End Function

    ''' <summary>Ringkasan exception untuk laporan (dibuka dari wrapper invocation).</summary>
    Public Shared Function DescribeError(ex As Exception) As String
        Dim current As Exception = ex
        While current.InnerException IsNot Nothing AndAlso
              (TypeOf current Is System.Reflection.TargetInvocationException)
            current = current.InnerException
        End While
        If current Is Nothing Then Return "tidak diketahui"
        Return current.GetType().Name & ": " & current.Message
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
            ' Review kedua: pengukuran yang TERPOTONG tidak boleh meninggalkan
            ' sidik jari struktur yang tampak lengkap. Tanpa tanda ini, koleksi
            ' yang jumlah sebenarnya lebih kecil/lebih besar dari Count yang
            ' dilaporkan bisa dibandingkan seolah-olah utuh.
            result.Signature.MarkUnreadable()
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

        Dim countRead As CountReadResult = TryReadCount(shapes)
        If Not countRead.IsOk Then
            ' Jumlah objek layer TIDAK diketahui: jangan mengarang hasil ukur.
            result.Bounds = BoundingBox.Empty()
            result.Signature.MarkUnreadable()
            result.Reason = "jumlah objek pada layer tidak dapat dibaca (" & countRead.Describe() & ")"
            Return result
        End If

        Dim total As Integer = countRead.Count
        If total < addedCount Then
            result.Reason = "koleksi layer hanya berisi " & total & " objek, kurang dari " & addedCount & " objek baru"
            Return result
        End If

        Dim measuredSlice As MeasuredShapeSet = MeasureSlice(shapes, total - addedCount + 1, addedCount, result.Source)
        measuredSlice.InputKind = MeasurementInputKind.ShapeCollection
        Return measuredSlice
    End Function

    ''' <summary>
    ''' Ukur sekumpulan objek pilihan (Shape tunggal / ShapeRange / koleksi COM).
    ''' Berguna untuk objek sumber sebelum di-copy; tidak mengubah pilihan
    ''' maupun objek.
    '''
    ''' Aturan (penting, hasil review):
    '''   1. Shape tunggal = objek yang MEMANG tidak punya anggota Count.
    '''      Kegagalan membaca Count (COM mati/RPC) TIDAK dianggap sebagai bukti
    '''      bahwa input adalah satu Shape.
    '''   2. ShapeRange / koleksi dengan Count terbaca -> setiap objek diukur.
    '''   3. Jumlah atau jenis objek tidak dapat dipastikan -> hasil dikembalikan
    '''      TANPA geometri dan dengan struktur ditandai tidak terbaca, sehingga
    '''      CopyVerification menghasilkan Partial (bukan Verified).
    ''' </summary>
    Public Shared Function MeasureSelection(rangeOrShape As Object) As MeasuredShapeSet
        Dim kind As MeasurementInputKind = ClassifyInput(rangeOrShape)
        Dim result As New MeasuredShapeSet()
        result.InputKind = kind

        If kind = MeasurementInputKind.Empty Then
            result.Source = "selection (kosong)"
            result.Reason = "selection kosong"
            Return result
        End If

        If kind = MeasurementInputKind.SingleShape Then
            Dim singleShape As MeasuredShapeSet =
                MeasureSingleShape(rangeOrShape, "selection (Shape tunggal, tanpa Count)")
            singleShape.InputKind = MeasurementInputKind.SingleShape
            Return singleShape
        End If

        Dim countRead As CountReadResult = TryReadCount(rangeOrShape)
        If kind = MeasurementInputKind.Unknown OrElse Not countRead.IsOk Then
            ' Jumlah objek tidak diketahui -> TIDAK mengukur apa pun.
            result.Source = "selection (" & DescribeInputKind(kind) & ")"
            result.Bounds = BoundingBox.Empty()
            result.Signature.MarkUnreadable()
            If countRead.IsMemberMissing Then
                ' Tidak punya Count, tetapi juga tidak terbukti berbentuk Shape:
                ' jumlah objek TIDAK boleh ditebak sebagai satu.
                result.Reason = "jumlah objek seleksi tidak dapat diverifikasi (" &
                                countRead.Describe() & "; bentuk objek juga tidak dapat dipastikan " &
                                "sebagai satu Shape karena kode tipenya tidak terbaca)"
            Else
                result.Reason = "jumlah objek seleksi tidak dapat diverifikasi (" &
                                countRead.Describe() & ")"
            End If
            Return result
        End If

        If countRead.Count = 0 Then
            result.Source = "selection (" & DescribeInputKind(kind) & ")"
            result.Reason = "selection kosong"
            Return result
        End If

        Dim measuredSlice As MeasuredShapeSet = MeasureSlice(rangeOrShape, 1, countRead.Count,
                                                             "selection (" & DescribeInputKind(kind) & ")")
        measuredSlice.InputKind = kind
        Return measuredSlice
    End Function

    ''' <summary>Ukur satu objek saja.</summary>
    Public Shared Function MeasureSingleShape(shape As Object, label As String) As MeasuredShapeSet
        Dim result As New MeasuredShapeSet()
        result.Source = If(label, "objek tunggal")
        If shape Is Nothing Then
            result.Reason = "objek tidak tersedia"
            Return result
        End If

        result.InputKind = MeasurementInputKind.SingleShape

        Dim typeCode As Integer = SafeTypeCode(shape)
        Dim typeReadable As Boolean = (typeCode >= 0)
        If typeReadable Then
            result.Signature.AddShapeType(typeCode)
        Else
            result.Signature.MarkUnreadable()
        End If

        Dim geometryError As String = ""
        Try
            result.Bounds = BoundingBox.FromCenter(CDbl(shape.CenterX), CDbl(shape.CenterY),
                                                   CDbl(shape.SizeWidth), CDbl(shape.SizeHeight))
        Catch ex As Exception
            result.Bounds = BoundingBox.Empty()
            geometryError = DescribeError(ex)
        End Try

        ' Alasan dilaporkan apa adanya: objek tunggal pun bisa tidak terukur,
        ' dan itu harus terlihat sebagai Partial, bukan sukses.
        Dim reasons As New List(Of String)()
        If Not typeReadable Then reasons.Add("tipe objek tidak dapat dibaca")
        If Not String.IsNullOrEmpty(geometryError) Then
            reasons.Add("geometri tidak dapat diukur: " & geometryError)
        End If
        If reasons.Count > 0 Then result.Reason = String.Join("; ", reasons.ToArray())
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

        ' Review kedua: fallback HANYA dipakai bila jumlah objek seleksi benar-
        ' benar terbaca DAN sama dengan jumlah objek baru. SafeCount (-1 untuk
        ' kedua jenis kegagalan) tidak lagi dipakai sebagai penentu, supaya
        ' "Count gagal dibaca" tidak pernah dianggap cocok.
        Dim selectionCount As CountReadResult = TryReadCount(selectionRange)
        If addedCount > 0 AndAlso selectionCount.IsOk AndAlso selectionCount.Count = addedCount Then
            Dim fallback As MeasuredShapeSet = MeasureSelection(selectionRange)
            fallback.Source = "selection setelah paste"
            If fallback.HasGeometry Then Return fallback
            If String.IsNullOrEmpty(primary.Reason) Then primary.Reason = fallback.Reason
        ElseIf addedCount > 0 AndAlso selectionRange IsNot Nothing AndAlso Not selectionCount.IsOk Then
            ' Alasan fallback tidak dipakai ikut dilaporkan (transparansi).
            Dim note As String = "fallback selection tidak dipakai (" & selectionCount.Describe() & ")"
            primary.Reason = If(String.IsNullOrEmpty(primary.Reason), note, primary.Reason & "; " & note)
        End If

        Return primary
    End Function
End Class
