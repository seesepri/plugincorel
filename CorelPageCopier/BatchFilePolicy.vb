Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  BatchFilePolicy.vb
'  Kebijakan file untuk batch watermark (Temuan D & E).
'
'  Masalah pada kode lama (WatermarkEngine.ApplyToFolder):
'    - `Directory.GetFiles(folder, "*.cdr")` ikut mengambil output lama
'      (*_wm.cdr) sehingga eksekusi berikutnya menghasilkan *_wm_wm.cdr dan
'      watermark berlapis.
'    - Status file dihitung dari partResult.Ok saja; kegagalan SaveAs masuk ke
'      daftar kegagalan BESAR tetapi file tetap dihitung "OK".
'    - Tidak ada pemeriksaan benturan nama / penimpaan file sumber.
'
'  Semua aturan di sini murni (tanpa COM / WinForms) sehingga bisa diuji unit.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>Status akhir satu file pada batch watermark.</summary>
Public Enum BatchFileStatus
    Ok = 0
    ' 'Partial' adalah kata kunci VB, jadi dinamai PartialSuccess.
    PartialSuccess = 1
    Failed = 2
End Enum

''' <summary>Satu file yang sengaja dilewati saat pengumpulan sumber batch.</summary>
Public Class SkippedBatchFile
    Public Path As String = ""
    Public Reason As String = ""

    Public Sub New()
    End Sub

    Public Sub New(filePath As String, reason As String)
        Me.Path = If(filePath, "")
        Me.Reason = If(reason, "")
    End Sub
End Class

''' <summary>Hasil pengumpulan sumber: yang diproses dan yang dilewati (dengan alasan).</summary>
Public Class BatchSourceSelection
    Public Sources As New List(Of String)()
    Public Skipped As New List(Of SkippedBatchFile)()
End Class

''' <summary>
''' Hasil satu file batch, dengan status terpisah untuk:
'''   - processing (penerapan watermark per halaman),
'''   - save       (penyimpanan output),
'''   - final      (Status, dipakai CSV + ringkasan + UI sekaligus).
''' </summary>
Public Class BatchFileOutcome
    Public FileName As String = ""
    Public SourcePath As String = ""
    Public OutputPath As String = ""

    ' --- processing ---------------------------------------------------------
    Public ApplyAttempted As Boolean = False
    Public PagesIntended As Integer = 0
    Public PagesApplied As Integer = 0
    Public PagesFailed As Integer = 0

    ' --- save ---------------------------------------------------------------
    ''' <summary>True bila SaveAs dipanggil untuk file ini.</summary>
    Public SaveAttempted As Boolean = False
    ''' <summary>
    ''' True bila PEMANGGILAN SaveAs tidak melempar. Terpisah dari SaveSucceeded:
    ''' pemanggilan bisa "berhasil" tanpa menulis apa pun.
    ''' </summary>
    Public SaveCallSucceeded As Boolean = False
    ''' <summary>True HANYA bila penulisan output terbukti (lihat SaveOutputVerification).</summary>
    Public SaveSucceeded As Boolean = False
    ''' <summary>Label hasil verifikasi output (BARU/DIPERBARUI/TIDAK BERUBAH/...).</summary>
    Public SaveVerificationLabel As String = ""
    ''' <summary>Alasan verifikasi output (ditampilkan apa adanya).</summary>
    Public SaveVerificationReason As String = ""

    ' --- statistik ----------------------------------------------------------
    Public ShapesCreated As Integer = 0
    Public DurationSeconds As Double = 0

    ' --- kegagalan fatal (file tidak bisa dibuka / exception tak terduga) ----
    Public FatalError As String = ""

    Public Messages As New List(Of String)()

    Public Sub New()
    End Sub

    Public Sub New(sourcePath As String)
        Me.SourcePath = If(sourcePath, "")
        Me.FileName = Path.GetFileName(Me.SourcePath)
    End Sub

    ''' <summary>Status akhir — SATU sumber kebenaran untuk CSV, ringkasan, dan UI.</summary>
    Public ReadOnly Property Status As BatchFileStatus
        Get
            Return BatchStatusRules.Evaluate(ApplyAttempted, PagesIntended, PagesApplied, PagesFailed,
                                             SaveAttempted, SaveSucceeded,
                                             Not String.IsNullOrEmpty(FatalError))
        End Get
    End Property

    Public Function StatusLabel() As String
        Return BatchStatusRules.Label(Status)
    End Function

    ''' <summary>Ringkasan pesan untuk CSV (tanpa newline).</summary>
    Public Function MessageText() As String
        Dim parts As New List(Of String)()
        If Not String.IsNullOrEmpty(FatalError) Then parts.Add(FatalError)
        For Each message As String In Messages
            If Not String.IsNullOrEmpty(message) Then parts.Add(message)
        Next
        If PagesIntended > 0 Then
            parts.Add("halaman " & PagesApplied.ToString(CultureInfo.InvariantCulture) & "/" &
                      PagesIntended.ToString(CultureInfo.InvariantCulture) & " diterapkan")
        End If
        If PagesFailed > 0 Then
            parts.Add(PagesFailed.ToString(CultureInfo.InvariantCulture) & " halaman gagal")
        End If
        If SaveAttempted Then
            If SaveSucceeded Then
                parts.Add("save OK" & If(String.IsNullOrEmpty(SaveVerificationLabel),
                                          "", " (" & SaveVerificationLabel & ")"))
            ElseIf SaveCallSucceeded Then
                ' SaveAs tidak melempar, tetapi penulisan tidak terbukti.
                parts.Add("SAVE TIDAK TERVERIFIKASI (" &
                          If(String.IsNullOrEmpty(SaveVerificationLabel), "?", SaveVerificationLabel) & ")")
            Else
                parts.Add("SAVE GAGAL")
            End If
        Else
            parts.Add("save tidak dijalankan")
        End If
        Return String.Join(" | ", parts.ToArray())
    End Function

    Public Function ToReportLine() As String
        Dim target As String = If(String.IsNullOrEmpty(OutputPath), "(tidak ada output)", OutputPath)
        Return "[" & StatusLabel() & "] " & FileName & " → " & target & " — " & MessageText()
    End Function
End Class

''' <summary>Aturan evaluasi status batch (murni, terpisah dari IO & COM).</summary>
Public Class BatchStatusRules

    ''' <summary>
    ''' Status OK hanya bila SEMUA langkah wajib berhasil:
    ''' watermark diterapkan ke seluruh halaman target DAN output tersimpan.
    ''' </summary>
    Public Shared Function Evaluate(applyAttempted As Boolean,
                                    pagesIntended As Integer,
                                    pagesApplied As Integer,
                                    pagesFailed As Integer,
                                    saveAttempted As Boolean,
                                    saveSucceeded As Boolean,
                                    hasFatalError As Boolean) As BatchFileStatus
        If hasFatalError Then Return BatchFileStatus.Failed
        If Not applyAttempted Then Return BatchFileStatus.Failed
        If pagesApplied <= 0 Then Return BatchFileStatus.Failed
        If Not saveAttempted Then Return BatchFileStatus.Failed
        If Not saveSucceeded Then Return BatchFileStatus.Failed
        If pagesFailed > 0 Then Return BatchFileStatus.PartialSuccess
        If pagesIntended > 0 AndAlso pagesApplied < pagesIntended Then Return BatchFileStatus.PartialSuccess
        Return BatchFileStatus.Ok
    End Function

    ''' <summary>
    ''' Evaluasi untuk operasi SATU dokumen (tanpa langkah save di dalam engine).
    ''' Dipakai agar status CSV operasi tunggal memakai aturan yang sama.
    ''' </summary>
    Public Shared Function EvaluateApplyOnly(applyAttempted As Boolean,
                                             pagesIntended As Integer,
                                             pagesApplied As Integer,
                                             pagesFailed As Integer,
                                             hasErrors As Boolean) As BatchFileStatus
        If hasErrors Then Return BatchFileStatus.Failed
        If Not applyAttempted Then Return BatchFileStatus.Failed
        If pagesApplied <= 0 Then Return BatchFileStatus.Failed
        If pagesFailed > 0 Then Return BatchFileStatus.PartialSuccess
        If pagesIntended > 0 AndAlso pagesApplied < pagesIntended Then Return BatchFileStatus.PartialSuccess
        Return BatchFileStatus.Ok
    End Function

    ''' <summary>Label status yang dipakai bersama oleh CSV, ringkasan, dan UI.</summary>
    Public Shared Function Label(status As BatchFileStatus) As String
        Select Case status
            Case BatchFileStatus.Ok
                Return "OK"
            Case BatchFileStatus.PartialSuccess
                Return "SEBAGIAN"
            Case Else
                Return "GAGAL"
        End Select
    End Function

    ''' <summary>Hitung jumlah file per status dari daftar hasil (satu sumber data).</summary>
    Public Shared Function CountStatus(outcomes As List(Of BatchFileOutcome), status As BatchFileStatus) As Integer
        Dim count As Integer = 0
        If outcomes Is Nothing Then Return 0
        For Each outcome As BatchFileOutcome In outcomes
            If outcome.Status = status Then count += 1
        Next
        Return count
    End Function
End Class

''' <summary>
''' Kebijakan nama file batch: filter output lama, penentuan path output,
''' deteksi benturan, dan validasi mode timpa.
''' </summary>
Public Class BatchFilePolicy

    ''' <summary>Akhiran default hasil batch (file asli tidak pernah ditimpa).</summary>
    Public Const DefaultOutputSuffix As String = "_wm"

    ''' <summary>Ekstensi yang diproses batch.</summary>
    Public Const SourceExtension As String = ".cdr"

    ''' <summary>
    ''' Rapikan akhiran output: tidak kosong, tanpa ekstensi, tanpa karakter
    ''' ilegal, dan selalu diawali pemisah supaya tidak menempel ke nama dasar.
    ''' </summary>
    Public Shared Function NormalizeOutputSuffix(suffix As String) As String
        Dim text As String = If(suffix, "").Trim()
        If text.Length = 0 Then Return DefaultOutputSuffix
        If text.EndsWith(SourceExtension, StringComparison.OrdinalIgnoreCase) Then
            text = text.Substring(0, text.Length - SourceExtension.Length).Trim()
        End If
        Dim cleaned As New StringBuilder()
        For Each ch As Char In text
            If Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "-"c Then
                cleaned.Append(ch)
            End If
        Next
        Dim result As String = cleaned.ToString()
        If result.Length = 0 Then Return DefaultOutputSuffix
        If Not (result.StartsWith("_") OrElse result.StartsWith("-")) Then result = "_" & result
        Return result
    End Function

    ''' <summary>True bila nama file terlihat sebagai output batch (mis. desain_wm.cdr).</summary>
    Public Shared Function IsGeneratedOutput(filePath As String, suffix As String) As Boolean
        If String.IsNullOrEmpty(filePath) Then Return False
        ' Hanya file .cdr yang bisa menjadi output batch; file lain diabaikan.
        If Not String.Equals(Path.GetExtension(filePath), SourceExtension, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If
        Dim normalized As String = NormalizeOutputSuffix(suffix)
        Dim fileName As String = Path.GetFileName(filePath)
        If String.IsNullOrEmpty(fileName) Then Return False
        Dim baseName As String = Path.GetFileNameWithoutExtension(fileName)
        If baseName Is Nothing Then Return False
        Return baseName.EndsWith(normalized, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Path output untuk satu file sumber (mode "simpan ke file baru").</summary>
    Public Shared Function BuildOutputPath(sourcePath As String, suffix As String) As String
        Dim normalized As String = NormalizeOutputSuffix(suffix)
        Dim folder As String = Path.GetDirectoryName(sourcePath)
        Dim baseName As String = Path.GetFileNameWithoutExtension(sourcePath)
        Return Path.Combine(If(folder, ""), baseName & normalized & SourceExtension)
    End Function

    ''' <summary>
    ''' Kumpulkan file sumber: keluarkan output lama (akhiran suffix) dan file
    ''' yang bukan .cdr. File yang dilewati dicatat beserta alasannya agar bisa
    ''' dilaporkan (tidak hilang diam-diam).
    ''' </summary>
    Public Shared Function SelectSources(candidates As IEnumerable(Of String),
                                         suffix As String) As BatchSourceSelection
        Dim selection As New BatchSourceSelection()
        If candidates Is Nothing Then Return selection
        Dim normalized As String = NormalizeOutputSuffix(suffix)
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each candidate As String In candidates
            If String.IsNullOrEmpty(candidate) Then Continue For
            If Not String.Equals(Path.GetExtension(candidate), SourceExtension, StringComparison.OrdinalIgnoreCase) Then
                selection.Skipped.Add(New SkippedBatchFile(candidate, "bukan file " & SourceExtension))
                Continue For
            End If
            If IsGeneratedOutput(candidate, normalized) Then
                selection.Skipped.Add(New SkippedBatchFile(
                    candidate, "output batch sebelumnya (akhiran """ & normalized & """); tidak diproses ulang"))
                Continue For
            End If
            If Not seen.Add(Path.GetFullPath(candidate)) Then
                selection.Skipped.Add(New SkippedBatchFile(candidate, "duplikat path"))
                Continue For
            End If
            selection.Sources.Add(candidate)
        Next

        selection.Sources.Sort(StringComparer.OrdinalIgnoreCase)
        Return selection
    End Function

    ''' <summary>
    ''' Cari nama output yang belum dipakai bila path kandidat benturan.
    ''' isTaken = fungsi pemeriksa (dipakai agar bisa diuji tanpa file system).
    ''' </summary>
    Public Shared Function ResolveUniqueOutputPath(candidatePath As String,
                                                   isTaken As Func(Of String, Boolean)) As String
        If String.IsNullOrEmpty(candidatePath) Then Return candidatePath
        If isTaken Is Nothing Then Return candidatePath
        If Not isTaken(candidatePath) Then Return candidatePath

        Dim folder As String = Path.GetDirectoryName(candidatePath)
        Dim baseName As String = Path.GetFileNameWithoutExtension(candidatePath)
        For i As Integer = 2 To 1000
            Dim attempt As String = Path.Combine(If(folder, ""),
                                                 baseName & "_" & i.ToString(CultureInfo.InvariantCulture) & SourceExtension)
            If Not isTaken(attempt) Then Return attempt
        Next
        Return candidatePath
    End Function

    ''' <summary>True bila dua path menunjuk file yang sama (case-insensitive).</summary>
    Public Shared Function SamePath(left As String, right As String) As Boolean
        If String.IsNullOrEmpty(left) OrElse String.IsNullOrEmpty(right) Then Return False
        Return String.Equals(TrimPath(left), TrimPath(right), StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function TrimPath(value As String) As String
        Try
            Return Path.GetFullPath(value).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Catch
            Return value.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        End Try
    End Function

    ''' <summary>True bila path berada di dalam folder (opsional termasuk subfolder).</summary>
    Public Shared Function IsUnderFolder(filePath As String, folder As String, includeSubfolders As Boolean) As Boolean
        If String.IsNullOrEmpty(filePath) OrElse String.IsNullOrEmpty(folder) Then Return False
        Dim fullFile As String = TrimPath(filePath)
        Dim fullFolder As String = TrimPath(folder)
        If SamePath(fullFile, fullFolder) Then Return False

        Dim prefix As String = fullFolder & Path.DirectorySeparatorChar
        If Not fullFile.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return False
        If includeSubfolders Then Return True

        Dim remainder As String = fullFile.Substring(prefix.Length)
        Return remainder.IndexOf(Path.DirectorySeparatorChar) < 0 AndAlso
               remainder.IndexOf(Path.AltDirectorySeparatorChar) < 0
    End Function

    ''' <summary>
    ''' Validasi ketat mode timpa (saveInPlace). Mengembalikan pesan kesalahan,
    ''' atau "" bila target aman ditimpa.
    ''' </summary>
    Public Shared Function ValidateOverwriteTarget(sourcePath As String, folder As String,
                                                   includeSubfolders As Boolean) As String
        If String.IsNullOrEmpty(sourcePath) Then Return "path file sumber kosong."
        If Not String.Equals(Path.GetExtension(sourcePath), SourceExtension, StringComparison.OrdinalIgnoreCase) Then
            Return "target bukan file " & SourceExtension & "."
        End If
        If Not IsUnderFolder(sourcePath, folder, includeSubfolders) Then
            Return "target berada di luar folder batch yang dipilih."
        End If
        If Not File.Exists(sourcePath) Then Return "file sumber tidak ditemukan." 
        If (File.GetAttributes(sourcePath) And FileAttributes.Directory) = FileAttributes.Directory Then
            Return "target adalah folder, bukan file."
        End If
        Try
            Dim attributes As FileAttributes = File.GetAttributes(sourcePath)
            If (attributes And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
                Return "file bersifat read-only; tidak dapat ditimpa."
            End If
        Catch ex As Exception
            Return "atribut file tidak dapat dibaca: " & ex.Message
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Validasi target pada mode "file baru": tidak boleh sama dengan sumber.
    ''' Mengembalikan pesan kesalahan atau "".
    ''' </summary>
    Public Shared Function ValidateNewOutputPath(sourcePath As String, outputPath As String) As String
        If String.IsNullOrEmpty(outputPath) Then Return "path output kosong."
        If SamePath(sourcePath, outputPath) Then
            Return "output akan menimpa file sumber; dijalankan dalam mode file baru."
        End If
        If Not String.Equals(Path.GetExtension(outputPath), SourceExtension, StringComparison.OrdinalIgnoreCase) Then
            Return "output bukan file " & SourceExtension & "."
        End If
        Return ""
    End Function
End Class
