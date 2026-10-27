Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ShapeScanner.vb
'  Penelusuran pohon objek untuk Auto Select (murni, tanpa COM/WinForms).
'
'  Yang diperbaiki dibanding implementasi lama di Program.vb:
'    1. Grup bersarang: kedalaman dibatasi (MaxDepth) DAN pemotongan dicatat
'       (sebelumnya hanya `Return` tanpa jejak).
'    2. Kegagalan membaca anak grup / isi PowerClip dicatat sebagai diagnostik
'       (sebelumnya `Catch` kosong).
'    3. PowerClip: isi container dibaca lewat jalur khusus (shape.Shapes tidak
'       memuat isi PowerClip). Hasil yang berada di dalam PowerClip diarahkan
'       ke container sebagai target seleksi.
'    4. Watermark: objek WM_/layer WATERMARK dilewati pada level mana pun.
'       Di dalam grup, hanya anak watermark yang dilewati (saudaranya yang sah
'       tetap bisa terpilih) — kecuali mode "Semua objek" yang memilih grup
'       sebagai satu kesatuan, sehingga grup yang memuat watermark dilewati.
'    5. Layer terkunci: objek tetap diklasifikasi tetapi ditandai pada
'       diagnostik dan tidak diklaim sebagai terpilih.
'    6. Hasil seleksi di-dedupe (satu container PowerClip tidak ditambahkan
'       berkali-kali walau beberapa isinya cocok).
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Text

''' <summary>Opsi pemindaian.</summary>
Public Class ShapeScanOptions
    Public Category As ShapeCategory = ShapeCategory.Text
    Public MaxDepth As Integer = 32
    ''' <summary>Pindai isi PowerClip (default True, dengan pencatatan kegagalan).</summary>
    Public ScanPowerClipContents As Boolean = True
    ''' <summary>Lewatkan objek pada layer terkunci (default True).</summary>
    Public SkipLockedLayers As Boolean = True
    ''' <summary>Sertakan objek yang berada di dalam PowerClip (target = container).</summary>
    Public IncludePowerClipHits As Boolean = True

    Public Shared Function ForCategory(category As ShapeCategory) As ShapeScanOptions
        Dim options As New ShapeScanOptions()
        options.Category = category
        Return options
    End Function
End Class

''' <summary>Satu objek yang cocok pada satu halaman.</summary>
Public Class ShapeHit
    Public PageIndex As Integer = 1
    Public KindLabel As String = ""
    Public ShapeName As String = ""
    Public LayerName As String = ""
    Public Depth As Integer = 0
    Public InsidePowerClip As Boolean = False
    Public OnLockedLayer As Boolean = False
    ''' <summary>Objek yang akan ditambahkan ke seleksi (container untuk isi PowerClip).</summary>
    Public SelectionTarget As Object = Nothing
    ''' <summary>Objek sebenarnya yang cocok (bisa di dalam PowerClip/grup).</summary>
    Public MatchedShape As Object = Nothing

    Public Function Describe() As String
        Dim text As String = KindLabel
        If Not String.IsNullOrEmpty(ShapeName) Then text &= " """ & ShapeName & """"
        If InsidePowerClip Then text &= " (di dalam PowerClip)"
        If Depth > 0 Then text &= " (kedalaman " & Depth.ToString(CultureInfo.InvariantCulture) & ")"
        Return text
    End Function
End Class

''' <summary>Hasil pemindaian satu halaman.</summary>
Public Class ShapeScanResult
    Public PageIndex As Integer = 0
    Public Category As ShapeCategory = ShapeCategory.Text
    Public Hits As New List(Of ShapeHit)()
    Public Diagnostics As New List(Of String)()

    ' --- statistik (dipakai laporan & test) ---------------------------------
    Public ScannedShapes As Integer = 0
    Public WatermarksSkipped As Integer = 0
    Public UnreadableShapes As Integer = 0
    Public DepthTruncated As Integer = 0
    Public ChildrenFailures As Integer = 0
    Public PowerClipsScanned As Integer = 0
    Public PowerClipFailures As Integer = 0
    Public LockedLayerSkipped As Integer = 0
    Public GroupsVisited As Integer = 0

    Public ReadOnly Property MatchCount As Integer
        Get
            Return Hits.Count
        End Get
    End Property

    ''' <summary>Jumlah target seleksi unik (container PowerClip tidak dobel).</summary>
    Public ReadOnly Property DistinctTargetCount As Integer
        Get
            Return SelectionTargets().Count
        End Get
    End Property

    ''' <summary>Target seleksi unik, urutan sesuai penemuan pertama.</summary>
    Public Function SelectionTargets() As List(Of Object)
        Dim targets As New List(Of Object)()
        For Each hit As ShapeHit In Hits
            If hit.OnLockedLayer Then Continue For
            Dim target As Object = hit.SelectionTarget
            If target Is Nothing Then Continue For
            If Not ContainsReference(targets, target) Then targets.Add(target)
        Next
        Return targets
    End Function

    ''' <summary>Daftar label hit untuk laporan (maks sejumlah baris).</summary>
    Public Function DescribeHits(Optional maxLines As Integer = 12) As String
        If Hits.Count = 0 Then Return "(tidak ada objek cocok)"
        Dim sb As New StringBuilder()
        Dim shown As Integer = 0
        For Each hit As ShapeHit In Hits
            If shown >= maxLines Then
                sb.AppendLine("  … dan " & (Hits.Count - shown).ToString(CultureInfo.InvariantCulture) & " objek lain")
                Exit For
            End If
            sb.AppendLine("  • " & hit.Describe())
            shown += 1
        Next
        Return sb.ToString()
    End Function

    ''' <summary>Ringkasan diagnostik untuk laporan (kosong bila tidak ada masalah).</summary>
    Public Function DescribeDiagnostics() As String
        If Diagnostics.Count = 0 Then Return ""
        Dim sb As New StringBuilder()
        For Each line As String In Diagnostics
            sb.AppendLine("  • " & line)
        Next
        Return sb.ToString()
    End Function

    Private Shared Function ContainsReference(items As List(Of Object), candidate As Object) As Boolean
        For Each item As Object In items
            If Object.ReferenceEquals(item, candidate) Then Return True
        Next
        Return False
    End Function
End Class

''' <summary>Penelusuran pohon objek yang bisa diuji tanpa CorelDRAW.</summary>
Public Class ShapeScanner

    ''' <summary>
    ''' Pindai daftar node akar (mis. isi Page.Shapes) untuk satu kategori.
    ''' Tidak pernah melempar exception karena masalah keterbacaan; semuanya
    ''' menjadi diagnostik.
    ''' </summary>
    Public Shared Function Scan(nodes As IList(Of IShapeNode), pageIndex As Integer,
                                options As ShapeScanOptions) As ShapeScanResult
        Dim result As New ShapeScanResult()
        result.PageIndex = pageIndex
        If options Is Nothing Then options = New ShapeScanOptions()
        result.Category = options.Category

        If nodes Is Nothing Then
            result.Diagnostics.Add("Daftar objek halaman kosong/tidak terbaca.")
            Return result
        End If

        For Each node As IShapeNode In nodes
            Visit(node, pageIndex, options, result, 0, Nothing, False)
        Next

        Return result
    End Function

    ''' <summary>
    ''' Telusuri satu node. rootShape = objek induk untuk konteks PowerClip,
    ''' target = target seleksi paksa (container PowerClip).
    ''' </summary>
    Private Shared Sub Visit(node As IShapeNode, pageIndex As Integer, options As ShapeScanOptions,
                             result As ShapeScanResult, depth As Integer,
                             forcedTarget As Object, insidePowerClip As Boolean)
        If node Is Nothing Then Return
        result.ScannedShapes += 1

        If depth > options.MaxDepth Then
            result.DepthTruncated += 1
            result.Diagnostics.Add("Kedalaman " & (depth).ToString(CultureInfo.InvariantCulture) &
                                   " melebihi batas " & options.MaxDepth.ToString(CultureInfo.InvariantCulture) &
                                   " — isi objek ini tidak dipindai.")
            Return
        End If

        Dim facts As ShapeFacts
        Try
            facts = node.GetFacts()
        Catch ex As Exception
            facts = Nothing
            result.UnreadableShapes += 1
            result.Diagnostics.Add("Fakta objek tidak dapat dibaca (" & ex.GetType().Name & ": " & ex.Message & ").")
        End Try

        If facts Is Nothing Then Return

        ' --- Watermark: selalu dikecualikan (termasuk isinya) ------------------
        If facts.IsWatermark Then
            result.WatermarksSkipped += 1
            Return
        End If

        If Not facts.TypeReadable Then
            result.UnreadableShapes += 1
            result.Diagnostics.Add("Tipe objek tidak dapat dibaca" &
                                   If(String.IsNullOrEmpty(facts.TypeError), "", ": " & facts.TypeError) & ".")
            Return
        End If

        Dim isGroup As Boolean = (facts.TypeCode = ShapeKind.GroupShape)
        If isGroup Then result.GroupsVisited += 1

        ' --- Mode "Semua objek": grup dipilih utuh, tidak ditelusuri ----------
        If options.Category = ShapeCategory.AllObjects Then
            Dim containsWatermark As Boolean = False
            If isGroup Then containsWatermark = GroupHasWatermark(node, options, result, depth + 1)

            If containsWatermark Then
                result.WatermarksSkipped += 1
                result.Diagnostics.Add("Grup """ & SafeName(facts) &
                                       """ dilewati: memuat objek watermark, dan grup dipilih sebagai satu kesatuan.")
                Return
            End If

            Dim classification As ClassificationResult = ShapeClassifier.Classify(facts, options.Category)
            If classification.Matches Then
                AddHit(result, pageIndex, facts, classification, depth, insidePowerClip,
                       If(forcedTarget, node.SelectionTarget))
            End If
            Return
        End If

        ' --- Kategori spesifik ------------------------------------------------
        Dim categoryResult As ClassificationResult = ShapeClassifier.Classify(facts, options.Category)
        If categoryResult.Matches Then
            ' Objek pada layer terkunci tetap DIHITUNG sebagai cocok (agar angka
            ' di laporan jujur: "N cocok, M tidak dapat dipilih"), tetapi tidak
            ' dimasukkan ke target seleksi (lihat ShapeScanResult.SelectionTargets).
            Dim lockedLayer As Boolean = options.SkipLockedLayers AndAlso
                                         facts.LayerEditableKnown AndAlso Not facts.LayerEditable
            If lockedLayer Then
                result.LockedLayerSkipped += 1
                result.Diagnostics.Add("Objek " & categoryResult.KindLabel & " """ & SafeName(facts) &
                                       """ pada layer terkunci """ & facts.LayerName &
                                       """ tidak dapat dipilih.")
            End If

            AddHit(result, pageIndex, facts, categoryResult, depth, insidePowerClip,
                   If(forcedTarget, node.SelectionTarget), lockedLayer)
        ElseIf Not String.IsNullOrEmpty(categoryResult.Note) Then
            ' Catatan non-fatal (mis. Curve.Closed tak terbaca) tetap dilaporkan.
            result.Diagnostics.Add(categoryResult.Note)
        End If

        ' --- Telusuri isi grup ------------------------------------------------
        If isGroup Then
            Dim children As IList(Of IShapeNode) = Nothing
            Try
                children = node.GetChildren()
            Catch ex As Exception
                result.ChildrenFailures += 1
                result.Diagnostics.Add("Isi grup """ & SafeName(facts) & """ tidak dapat dibaca (" &
                                       ex.GetType().Name & ": " & ex.Message & ").")
            End Try
            If children IsNot Nothing Then
                If children.Count = 0 AndAlso Not String.IsNullOrEmpty(node.ChildrenError) Then
                    result.ChildrenFailures += 1
                    result.Diagnostics.Add("Isi grup """ & SafeName(facts) & """ tidak dapat dibaca: " & node.ChildrenError)
                End If
                For Each child As IShapeNode In children
                    Visit(child, pageIndex, options, result, depth + 1, forcedTarget, insidePowerClip)
                Next
            End If

            ' --- Telusuri isi PowerClip (jalur terpisah!) ---------------------
            If options.ScanPowerClipContents Then
                VisitPowerClip(node, pageIndex, options, result, depth, insidePowerClip)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Isi PowerClip TIDAK berada di shape.Shapes, jadi dipindai lewat jalur
    ''' sendiri. Semua hit di dalamnya memakai container sebagai target seleksi.
    ''' </summary>
    Private Shared Sub VisitPowerClip(node As IShapeNode, pageIndex As Integer, options As ShapeScanOptions,
                                      result As ShapeScanResult, depth As Integer, insidePowerClip As Boolean)
        Dim isContainer As Boolean
        Try
            isContainer = node.IsPowerClipContainer
        Catch ex As Exception
            result.PowerClipFailures += 1
            result.Diagnostics.Add("Status PowerClip tidak dapat dibaca (" & ex.GetType().Name & ": " & ex.Message & ").")
            Return
        End Try
        If Not isContainer Then Return

        result.PowerClipsScanned += 1

        Dim contents As IList(Of IShapeNode) = Nothing
        Try
            contents = node.GetPowerClipChildren()
        Catch ex As Exception
            result.PowerClipFailures += 1
            result.Diagnostics.Add("Isi PowerClip tidak dapat dibaca (" & ex.GetType().Name & ": " & ex.Message & ").")
            Return
        End Try

        If contents Is Nothing OrElse contents.Count = 0 Then
            If Not String.IsNullOrEmpty(node.PowerClipError) Then
                result.PowerClipFailures += 1
                result.Diagnostics.Add("PowerClip terdeteksi tetapi isinya tidak dapat dibaca: " & node.PowerClipError &
                                       " — objek di dalamnya TIDAK dipindai (jumlah hasil bisa kurang).")
            End If
            Return
        End If

        For Each child As IShapeNode In contents
            Visit(child, pageIndex, options, result, depth + 1, node.SelectionTarget, True)
        Next
    End Sub

    ''' <summary>
    ''' True bila grup memuat objek watermark pada level mana pun.
    ''' Dipakai mode "Semua objek" (grup tidak boleh dipilih sebagian).
    ''' </summary>
    Private Shared Function GroupHasWatermark(node As IShapeNode, options As ShapeScanOptions,
                                              result As ShapeScanResult, depth As Integer) As Boolean
        If depth > options.MaxDepth Then
            ' Tidak dapat dipastikan: perlakukan sebagai "mengandung" agar aman
            ' (watermark tidak pernah ikut terpilih).
            result.Diagnostics.Add("Kedalaman grup melebihi batas saat memeriksa watermark — grup dilewati demi keamanan.")
            Return True
        End If

        Dim children As IList(Of IShapeNode) = Nothing
        Try
            children = node.GetChildren()
        Catch
            ' Gagal membaca isi: tidak bisa memastikan aman -> anggap memuat watermark.
            result.ChildrenFailures += 1
            result.Diagnostics.Add("Isi grup tidak dapat dibaca saat memeriksa watermark — grup dilewati demi keamanan.")
            Return True
        End Try
        If children Is Nothing Then Return False

        For Each child As IShapeNode In children
            Dim facts As ShapeFacts = Nothing
            Try
                facts = child.GetFacts()
            Catch
                Return True
            End Try
            If facts Is Nothing Then Continue For
            If facts.IsWatermark Then Return True
            If facts.TypeReadable AndAlso facts.TypeCode = ShapeKind.GroupShape Then
                If GroupHasWatermark(child, options, result, depth + 1) Then Return True
            End If
        Next
        Return False
    End Function

    Private Shared Sub AddHit(result As ShapeScanResult, pageIndex As Integer, facts As ShapeFacts,
                              classification As ClassificationResult, depth As Integer,
                              insidePowerClip As Boolean, target As Object, Optional lockedLayer As Boolean = False)
        Dim hit As New ShapeHit()
        hit.PageIndex = pageIndex
        hit.KindLabel = classification.KindLabel
        hit.ShapeName = facts.ShapeName
        hit.LayerName = facts.LayerName
        hit.Depth = depth
        hit.InsidePowerClip = insidePowerClip
        hit.OnLockedLayer = lockedLayer OrElse (facts.LayerEditableKnown AndAlso Not facts.LayerEditable)
        hit.MatchedShape = Nothing
        hit.SelectionTarget = target
        result.Hits.Add(hit)

        If Not String.IsNullOrEmpty(classification.Note) Then
            result.Diagnostics.Add(classification.Note & " [" & classification.KindLabel & "]")
        End If
    End Sub

    Private Shared Function SafeName(facts As ShapeFacts) As String
        If facts Is Nothing Then Return "?"
        If String.IsNullOrEmpty(facts.ShapeName) Then Return "(tanpa nama)"
        Return facts.ShapeName
    End Function
End Class
