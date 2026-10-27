Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  SaveOutputVerification.vb
'  Verifikasi output penyimpanan (review lanjutan PR #2, butir 2).
'
'  Masalah pada kode sebelumnya (WatermarkEngine.ApplyToFolder):
'    - keberhasilan pemanggilan SaveAs dan validasi file output dikerjakan
'      dalam satu blok yang sama, sehingga tidak jelas mana yang gagal;
'    - validasi hanya File.Exists + Length > 0. Pada mode TIMPA, file tujuan
'      adalah file sumber itu sendiri: file lama yang masih ada otomatis
'      "lulus" walau SaveAs tidak menulis apa pun (atau COM menelan
'      kegagalan). Pada mode file baru, file dari eksekusi sebelumnya juga
'      membuat File.Exists bernilai True tanpa membuktikan apa pun.
'
'  Aturan di sini:
'    1. Pemanggilan SaveAs dicatat terpisah (SaveCallSucceeded).
'    2. Keberhasilan penyimpanan HARUS dibuktikan lewat perubahan keadaan file
'       yang diambil SEBELUM SaveAs: file belum ada, atau ukuran/waktu
'       modifikasi berubah.
'    3. Bila tidak ada perubahan, status = TIDAK TERVERIFIKASI (Unchanged),
'       bukan sukses.
'    4. TIDAK membuka ulang dokumen untuk memeriksa hasil: dokumen aktif dan
'       instance COM milik pengguna tidak boleh terganggu. Semua pemeriksaan
'       memakai metadata sistem file + beberapa byte pertama isi file.
'
'  Tambahan pada review follow-up kedua (temuan G):
'    5. Ukuran + waktu modifikasi saja dapat DUSTA: banyak filesystem menyimpan
'       waktu dengan granularitas 1-2 detik, sehingga file yang benar-benar
'       ditulis ulang dengan panjang sama bisa terlihat "tidak berubah"
'       (menghasilkan GAGAL palsu). Karena itu isi file diberi sidik jari
'       (SHA-256 atas sebagian isi), dan perbedaan sidik jari dihitung sebagai
'       bukti perubahan.
'    6. "Ada + tidak kosong" bukan bukti file VALID. Berkas .cdr adalah
'       kontainer RIFF (CorelDRAW <= X3) atau paket ZIP (X4+, termasuk 2022),
'       jadi bentuk dasarnya diperiksa dari 12 byte pertama. Bentuk yang tidak
'       dikenal -> status Invalid (bukan sukses).
'
'  Kelas di file ini murni operasi file (tanpa COM, tanpa WinForms) sehingga
'  seluruh cabang dapat diuji unit.
' ============================================================================

Imports System
Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Threading

''' <summary>Hasil pemeriksaan satu output penyimpanan.</summary>
Public Enum SaveOutputStatus
    ''' <summary>File output belum ada sebelum SaveAs dan sekarang ada.</summary>
    VerifiedNew = 0
    ''' <summary>File output sudah ada sebelumnya dan terbukti berubah.</summary>
    VerifiedChanged = 1
    ''' <summary>File output tidak ada setelah SaveAs.</summary>
    Missing = 2
    ''' <summary>File output ada tetapi kosong (0 byte).</summary>
    Empty = 3
    ''' <summary>File sudah ada sebelumnya dan TIDAK berubah -> tidak dapat dibuktikan.</summary>
    Unchanged = 4
    ''' <summary>Keadaan file tidak dapat diperiksa (sebelum atau sesudah SaveAs).</summary>
    Unreadable = 5
    ''' <summary>
    ''' File ada dan tidak kosong, tetapi bentuk isinya bukan berkas CorelDRAW
    ''' (tanda tangan RIFF/CDR atau paket ZIP tidak ditemukan).
    ''' </summary>
    Invalid = 6
End Enum

''' <summary>Mode penyimpanan batch; memengaruhi cara bukti penulisan dinilai.</summary>
Public Enum SaveOutputMode
    ''' <summary>Output baru (&lt;nama&gt;_wm.cdr); file boleh belum ada.</summary>
    NewFile = 0
    ''' <summary>Menimpa file sumber; file tujuan WAJIB sudah ada sebelum SaveAs.</summary>
    OverwriteInPlace = 1
End Enum

''' <summary>
''' Keadaan satu file pada satu titik waktu. Diambil sebelum dan sesudah
''' SaveAs. Kegagalan membaca metadata dicatat, tidak disembunyikan.
'''
''' Metadata (ada/ukuran/waktu) saja TIDAK cukup untuk membuktikan penulisan:
''' granularitas waktu filesystem bisa 1-2 detik. Karena itu, bila diminta,
''' isi file juga diberi sidik jari dan bentuk dasarnya diklasifikasi.
''' </summary>
Public Class SaveOutputSnapshot
    ''' <summary>Batas percobaan membuka file (CorelDRAW bisa masih memegang file sesaat setelah SaveAs).</summary>
    Public Const ContentReadAttempts As Integer = 3
    ''' <summary>Jeda antarpercobaan (milidetik).</summary>
    Public Const ContentRetryDelayMs As Integer = 40
    ''' <summary>Jumlah byte awal yang dibaca untuk bentuk file.</summary>
    Public Const HeaderBytes As Integer = 12
    ''' <summary>Jumlah byte awal/akhir yang ikut di-sidik-jari.</summary>
    Public Const FingerprintChunkBytes As Integer = 4096

    ''' <summary>Tanda tangan berkas CorelDRAW klasik (RIFF + form type "CD...").</summary>
    Public Const SignatureRiffCdr As String = "RIFF/CDR"
    ''' <summary>Tanda tangan paket ZIP (CDR CorelDRAW X4 ke atas, termasuk 2022).</summary>
    Public Const SignatureZip As String = "ZIP"
    ''' <summary>Bentuk file tidak dikenali.</summary>
    Public Const SignatureUnknown As String = "TIDAK DIKENALI"

    Public Path As String = ""
    Public Exists As Boolean = False
    Public Length As Long = 0
    Public LastWriteUtc As DateTime = DateTime.MinValue
    ''' <summary>Alasan bila metadata tidak dapat dibaca (kosong = berhasil dibaca).</summary>
    Public CaptureError As String = ""

    ''' <summary>Jenis bentuk isi file (SignatureRiffCdr/SignatureZip/SignatureUnknown); "" bila tidak dibaca.</summary>
    Public SignatureKind As String = ""
    ''' <summary>Sidik jari isi file (SHA-256 heksadesimal); "" bila tidak dihitung.</summary>
    Public Fingerprint As String = ""
    ''' <summary>True bila isi file berhasil dibaca (bentuk + sidik jari tersedia).</summary>
    Public ContentRead As Boolean = False
    ''' <summary>Alasan bila isi file tidak dapat dibaca (kosong = berhasil).</summary>
    Public ContentError As String = ""

    Public ReadOnly Property IsValid As Boolean
        Get
            Return String.IsNullOrEmpty(CaptureError)
        End Get
    End Property

    ''' <summary>Ambil keadaan file (metadata saja); TIDAK pernah melempar.</summary>
    Public Shared Function Capture(path As String) As SaveOutputSnapshot
        Return Capture(path, False)
    End Function

    ''' <summary>
    ''' Ambil keadaan file. readContent:=True juga membaca bentuk isi + sidik
    ''' jari (memakai FileShare.ReadWrite supaya file yang masih dipegang
    ''' CorelDRAW tidak langsung dianggap gagal). TIDAK pernah melempar.
    ''' </summary>
    Public Shared Function Capture(path As String, readContent As Boolean) As SaveOutputSnapshot
        Dim snapshot As New SaveOutputSnapshot()
        snapshot.Path = If(path, "")
        If String.IsNullOrEmpty(snapshot.Path) Then
            snapshot.CaptureError = "path output kosong"
            Return snapshot
        End If

        Try
            Dim info As New FileInfo(snapshot.Path)
            snapshot.Exists = info.Exists
            If info.Exists Then
                snapshot.Length = info.Length
                snapshot.LastWriteUtc = info.LastWriteTimeUtc
            End If
        Catch ex As Exception
            snapshot.Exists = False
            snapshot.CaptureError = ex.GetType().Name & ": " & ex.Message
            Return snapshot
        End Try

        If readContent AndAlso snapshot.Exists Then snapshot.ReadContent()
        Return snapshot
    End Function

    ''' <summary>
    ''' Bangun snapshot untuk test tanpa menyentuh sistem file.
    ''' fingerprint/signatureKind opsional: bila diisi, ContentRead dianggap True.
    ''' </summary>
    Public Shared Function DescribeState(exists As Boolean, length As Long,
                                         lastWriteUtc As DateTime,
                                         Optional fingerprint As String = "",
                                         Optional signatureKind As String = "") As SaveOutputSnapshot
        Dim snapshot As New SaveOutputSnapshot()
        snapshot.Exists = exists
        snapshot.Length = length
        snapshot.LastWriteUtc = lastWriteUtc
        snapshot.Fingerprint = If(fingerprint, "")
        snapshot.SignatureKind = If(signatureKind, "")
        snapshot.ContentRead = (snapshot.Fingerprint <> "" OrElse snapshot.SignatureKind <> "")
        Return snapshot
    End Function

    ''' <summary>Baca bentuk isi + sidik jari. Kegagalan dicatat, tidak dilempar.</summary>
    Private Sub ReadContent()
        Dim lastError As String = ""
        For attempt As Integer = 1 To ContentReadAttempts
            Try
                Using stream As New FileStream(Me.Path, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite Or FileShare.Delete)
                    Dim header(HeaderBytes - 1) As Byte
                    Dim headerLength As Integer = ReadUpTo(stream, header, HeaderBytes)
                    Me.SignatureKind = ClassifySignature(header, headerLength)
                    Me.Fingerprint = ComputeFingerprint(stream, header, headerLength)
                    Me.ContentRead = True
                    Me.ContentError = ""
                    Return
                End Using
            Catch ex As Exception
                lastError = DescribeError(ex)
                If attempt < ContentReadAttempts Then
                    Try
                        Thread.Sleep(ContentRetryDelayMs)
                    Catch
                    End Try
                End If
            End Try
        Next
        Me.ContentRead = False
        Me.SignatureKind = ""
        Me.Fingerprint = ""
        Me.ContentError = lastError
    End Sub

    ''' <summary>
    ''' Klasifikasi bentuk berkas dari byte awal:
    '''   "RIFF" + form type "CD.." -> RIFF/CDR (CorelDRAW <= X3),
    '''   "PK" (ZIP)                -> ZIP      (CorelDRAW X4+ / 2022),
    '''   lainnya                   -> TIDAK DIKENALI.
    ''' Pemeriksaan ini SENGAJA dangkal: hanya membuktikan bentuk wadah, bukan
    ''' bahwa seluruh isi berkas valid. Karena itu ia dipakai untuk menolak,
    ''' bukan untuk menjamin.
    ''' </summary>
    Public Shared Function ClassifySignature(header As Byte(), length As Integer) As String
        If header Is Nothing OrElse length < 4 Then Return SignatureUnknown

        If header(0) = &H52 AndAlso header(1) = &H49 AndAlso header(2) = &H46 AndAlso header(3) = &H46 Then
            ' "RIFF": form type pada offset 8 harus diawali "CD" (CDR/CDRX/CDT...).
            If length >= 10 AndAlso header(8) = &H43 AndAlso header(9) = &H44 Then Return SignatureRiffCdr
            Return SignatureUnknown
        End If

        ' "PK" 03/05/07: paket ZIP (CDR X4+, termasuk CorelDRAW 2022).
        If header(0) = &H50 AndAlso header(1) = &H4B AndAlso
           (header(2) = &H3 OrElse header(2) = &H5 OrElse header(2) = &H7) Then
            Return SignatureZip
        End If

        Return SignatureUnknown
    End Function

    Private Shared Function ReadUpTo(stream As FileStream, buffer As Byte(), count As Integer) As Integer
        Dim total As Integer = 0
        While total < count
            Dim read As Integer = stream.Read(buffer, total, count - total)
            If read <= 0 Then Exit While
            total += read
        End While
        Return total
    End Function

    ''' <summary>
    ''' Sidik jari isi: SHA-256 atas (panjang file + potongan awal + potongan akhir).
    ''' Murah untuk file besar, tetapi tetap membedakan dua berkas berukuran sama.
    ''' </summary>
    Private Shared Function ComputeFingerprint(stream As FileStream, header As Byte(), headerLength As Integer) As String
        Dim builder As New System.Text.StringBuilder()
        builder.Append(stream.Length.ToString(CultureInfo.InvariantCulture)).Append("|")
        builder.Append(ToHex(header, headerLength)).Append("|")

        Dim tailLength As Integer = CInt(Math.Min(CLng(FingerprintChunkBytes), stream.Length))
        If tailLength > 0 Then
            stream.Seek(-tailLength, SeekOrigin.End)
            Dim tail(tailLength - 1) As Byte
            Dim read As Integer = ReadUpTo(stream, tail, tailLength)
            builder.Append(ToHex(tail, read))
        End If

        Using sha As SHA256 = SHA256.Create()
            Dim data As Byte() = System.Text.Encoding.UTF8.GetBytes(builder.ToString())
            Dim hash As Byte() = sha.ComputeHash(data)
            Return ToHex(hash, hash.Length)
        End Using
    End Function

    Private Shared Function ToHex(data As Byte(), length As Integer) As String
        If data Is Nothing OrElse length <= 0 Then Return ""
        Dim chars As New System.Text.StringBuilder(length * 2)
        For i As Integer = 0 To length - 1
            chars.Append(data(i).ToString("x2", CultureInfo.InvariantCulture))
        Next
        Return chars.ToString()
    End Function

    Private Shared Function DescribeError(ex As Exception) As String
        If ex Is Nothing Then Return "tidak diketahui"
        Return ex.GetType().Name & ": " & ex.Message
    End Function

    Public Function Describe() As String
        If Not IsValid Then Return "tidak dapat diperiksa (" & CaptureError & ")"
        If Not Exists Then Return "tidak ada"
        Dim text As String = Length.ToString(CultureInfo.InvariantCulture) & " byte, diubah " &
               LastWriteUtc.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) & " UTC"
        If ContentRead Then
            text &= ", bentuk " & If(String.IsNullOrEmpty(SignatureKind), "?", SignatureKind)
            text &= If(String.IsNullOrEmpty(Fingerprint), "", ", sidik jari " & Fingerprint.Substring(0, Math.Min(12, Fingerprint.Length)))
        ElseIf Not String.IsNullOrEmpty(ContentError) Then
            text &= ", isi tidak terbaca (" & ContentError & ")"
        End If
        Return text
    End Function
End Class


''' <summary>Keputusan verifikasi satu output penyimpanan.</summary>
Public Class SaveOutputVerification
    Public Status As SaveOutputStatus = SaveOutputStatus.Unreadable
    Public Reason As String = ""
    Public Mode As SaveOutputMode = SaveOutputMode.NewFile

    ''' <summary>
    ''' True HANYA bila penulisan terbukti: file baru muncul, atau file lama
    ''' berubah ukuran/waktu modifikasi setelah SaveAs.
    ''' </summary>
    Public ReadOnly Property IsVerified As Boolean
        Get
            Return Status = SaveOutputStatus.VerifiedNew OrElse Status = SaveOutputStatus.VerifiedChanged
        End Get
    End Property

    Public ReadOnly Property StatusLabel As String
        Get
            Select Case Status
                Case SaveOutputStatus.VerifiedNew
                    Return "BARU"
                Case SaveOutputStatus.VerifiedChanged
                    Return "DIPERBARUI"
                Case SaveOutputStatus.Missing
                    Return "TIDAK DITEMUKAN"
                Case SaveOutputStatus.Empty
                    Return "KOSONG"
                Case SaveOutputStatus.Unchanged
                    Return "TIDAK BERUBAH"
                Case SaveOutputStatus.Invalid
                    Return "TIDAK VALID"
                Case Else
                    Return "TIDAK DAPAT DIPERIKSA"
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Nilai satu penyimpanan dari keadaan file SEBELUM dan SESUDAH SaveAs.
    ''' Tidak ada asumsi: setiap temuan punya alasan yang bisa ditampilkan.
    '''
    ''' requireCorelSignature:=True menambahkan pemeriksaan bentuk berkas:
    '''   - isi terbaca & bukan RIFF/CDR maupun ZIP  -> Invalid (bukan sukses),
    '''   - isi TIDAK terbaca (mis. masih dikunci CorelDRAW) -> pemeriksaan
    '''     bentuk dilewati dan dicatat di alasan; bukti penulisan tetap dari
    '''     metadata/sidik jari, sehingga hasil tidak pernah mengklaim lebih.
    ''' </summary>
    Public Shared Function Verify(before As SaveOutputSnapshot, after As SaveOutputSnapshot,
                                  Optional mode As SaveOutputMode = SaveOutputMode.NewFile,
                                  Optional requireCorelSignature As Boolean = False) As SaveOutputVerification
        Dim result As New SaveOutputVerification()
        result.Mode = mode

        ' 1. Keadaan sesudah SaveAs wajib terbaca.
        If after Is Nothing OrElse Not after.IsValid Then
            result.Status = SaveOutputStatus.Unreadable
            result.Reason = "keadaan file output sesudah SaveAs tidak dapat diperiksa" &
                            If(after Is Nothing OrElse String.IsNullOrEmpty(after.CaptureError),
                               "", ": " & after.CaptureError)
            Return result
        End If

        ' 2. File wajib ada.
        If Not after.Exists Then
            result.Status = SaveOutputStatus.Missing
            result.Reason = "file output tidak ditemukan setelah SaveAs selesai."
            Return result
        End If

        ' 3. File wajib berisi.
        If after.Length <= 0 Then
            result.Status = SaveOutputStatus.Empty
            result.Reason = "file output ada tetapi berukuran 0 byte; penyimpanan tidak dianggap berhasil."
            Return result
        End If

        ' 4. Bentuk berkas (opsional): "ada + tidak kosong" BUKAN bukti valid.
        Dim signatureNote As String = ""
        If requireCorelSignature Then
            If after.ContentRead Then
                If after.SignatureKind <> SaveOutputSnapshot.SignatureRiffCdr AndAlso
                   after.SignatureKind <> SaveOutputSnapshot.SignatureZip Then
                    result.Status = SaveOutputStatus.Invalid
                    result.Reason = "file output ada dan tidak kosong, tetapi bentuknya bukan berkas " &
                                    "CorelDRAW (tanda tangan: " &
                                    If(String.IsNullOrEmpty(after.SignatureKind), "?", after.SignatureKind) &
                                    "); isi 12 byte pertama: " & DescribeHeader(after) & "."
                    Return result
                End If
            Else
                signatureNote = " Bentuk berkas tidak dapat diperiksa (" &
                                If(String.IsNullOrEmpty(after.ContentError), "isi tidak terbaca", after.ContentError) & ")."
            End If
        End If

        ' 5. Keadaan sebelum SaveAs wajib diketahui, kalau tidak tidak ada bukti.
        If before Is Nothing OrElse Not before.IsValid Then
            result.Status = SaveOutputStatus.Unreadable
            result.Reason = "keadaan file sebelum SaveAs tidak dapat dicatat, sehingga keberhasilan " &
                            "penyimpanan terbaru tidak dapat dibuktikan." & signatureNote
            Return result
        End If

        ' 6. Mode timpa: file sumber harus sudah ada sebelumnya.
        If mode = SaveOutputMode.OverwriteInPlace AndAlso Not before.Exists Then
            result.Status = SaveOutputStatus.Unreadable
            result.Reason = "mode timpa dipakai, tetapi file sumber tidak ada sebelum SaveAs; " &
                            "bukti penyimpanan tidak dapat dinilai." & signatureNote
            Return result
        End If

        ' 7. File belum ada sebelumnya -> bukti paling kuat.
        If Not before.Exists Then
            result.Status = SaveOutputStatus.VerifiedNew
            result.Reason = "file output baru dibuat (" & after.Length.ToString(CultureInfo.InvariantCulture) &
                            " byte)." & signatureNote
            Return result
        End If

        ' 8. File sudah ada sebelumnya -> wajib ada perubahan yang terbukti.
        '    Perubahan isi (sidik jari) ikut dihitung supaya penulisan ulang
        '    dengan panjang sama pada filesystem bergranularitas kasar tidak
        '    dilaporkan sebagai "tidak berubah".
        Dim changedByMetadata As Boolean = (after.Length <> before.Length OrElse
                                            after.LastWriteUtc <> before.LastWriteUtc)
        Dim fingerprintCompared As Boolean = (before.Fingerprint <> "" AndAlso after.Fingerprint <> "")
        Dim changedByContent As Boolean = fingerprintCompared AndAlso before.Fingerprint <> after.Fingerprint

        If changedByMetadata OrElse changedByContent Then
            result.Status = SaveOutputStatus.VerifiedChanged
            result.Reason = "file output diperbarui: " &
                            before.Length.ToString(CultureInfo.InvariantCulture) & " byte / " &
                            before.LastWriteUtc.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) &
                            " UTC -> " &
                            after.Length.ToString(CultureInfo.InvariantCulture) & " byte / " &
                            after.LastWriteUtc.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) & " UTC" &
                            If(changedByContent AndAlso Not changedByMetadata,
                               " (ukuran & waktu sama, tetapi isi file berubah)", ".") &
                            signatureNote
            Return result
        End If

        ' 9. Tidak ada perubahan: file lama TIDAK boleh dianggap bukti sukses.
        result.Status = SaveOutputStatus.Unchanged
        result.Reason = "file output sudah ada sebelum SaveAs dan tidak berubah (" &
                        If(fingerprintCompared,
                           "ukuran, waktu modifikasi, dan isi file sama",
                           "ukuran & waktu modifikasi sama; isi file tidak dibandingkan") &
                        "), sehingga penyimpanan terbaru tidak dapat dibuktikan." & signatureNote
        Return result
    End Function

    ''' <summary>Cuplikan tanda tangan isi untuk laporan (tidak membaca ulang file).</summary>
    Private Shared Function DescribeHeader(after As SaveOutputSnapshot) As String
        If after Is Nothing Then Return "?"
        If String.IsNullOrEmpty(after.Fingerprint) Then Return "tidak tersedia"
        Return "sidik jari " & after.Fingerprint.Substring(0, Math.Min(12, after.Fingerprint.Length))
    End Function
End Class
