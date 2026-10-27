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
'       memakai metadata sistem file.
'
'  Kelas di file ini murni operasi file (tanpa COM, tanpa WinForms) sehingga
'  seluruh cabang dapat diuji unit.
' ============================================================================

Imports System
Imports System.Globalization
Imports System.IO

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
''' </summary>
Public Class SaveOutputSnapshot
    Public Path As String = ""
    Public Exists As Boolean = False
    Public Length As Long = 0
    Public LastWriteUtc As DateTime = DateTime.MinValue
    ''' <summary>Alasan bila metadata tidak dapat dibaca (kosong = berhasil dibaca).</summary>
    Public CaptureError As String = ""

    Public ReadOnly Property IsValid As Boolean
        Get
            Return String.IsNullOrEmpty(CaptureError)
        End Get
    End Property

    ''' <summary>Ambil keadaan file; TIDAK pernah melempar.</summary>
    Public Shared Function Capture(path As String) As SaveOutputSnapshot
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
        End Try
        Return snapshot
    End Function

    ''' <summary>Bangun snapshot untuk test tanpa menyentuh sistem file.</summary>
    Public Shared Function DescribeState(exists As Boolean, length As Long,
                                         lastWriteUtc As DateTime) As SaveOutputSnapshot
        Dim snapshot As New SaveOutputSnapshot()
        snapshot.Exists = exists
        snapshot.Length = length
        snapshot.LastWriteUtc = lastWriteUtc
        Return snapshot
    End Function

    Public Function Describe() As String
        If Not IsValid Then Return "tidak dapat diperiksa (" & CaptureError & ")"
        If Not Exists Then Return "tidak ada"
        Return Length.ToString(CultureInfo.InvariantCulture) & " byte, diubah " &
               LastWriteUtc.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture) & " UTC"
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
                Case Else
                    Return "TIDAK DAPAT DIPERIKSA"
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Nilai satu penyimpanan dari keadaan file SEBELUM dan SESUDAH SaveAs.
    ''' Tidak ada asumsi: setiap temuan punya alasan yang bisa ditampilkan.
    ''' </summary>
    Public Shared Function Verify(before As SaveOutputSnapshot, after As SaveOutputSnapshot,
                                  Optional mode As SaveOutputMode = SaveOutputMode.NewFile) As SaveOutputVerification
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

        ' 4. Keadaan sebelum SaveAs wajib diketahui, kalau tidak tidak ada bukti.
        If before Is Nothing OrElse Not before.IsValid Then
            result.Status = SaveOutputStatus.Unreadable
            result.Reason = "keadaan file sebelum SaveAs tidak dapat dicatat, sehingga keberhasilan " &
                            "penyimpanan terbaru tidak dapat dibuktikan."
            Return result
        End If

        ' 5. Mode timpa: file sumber harus sudah ada sebelumnya.
        If mode = SaveOutputMode.OverwriteInPlace AndAlso Not before.Exists Then
            result.Status = SaveOutputStatus.Unreadable
            result.Reason = "mode timpa dipakai, tetapi file sumber tidak ada sebelum SaveAs; " &
                            "bukti penyimpanan tidak dapat dinilai."
            Return result
        End If

        ' 6. File belum ada sebelumnya -> bukti paling kuat.
        If Not before.Exists Then
            result.Status = SaveOutputStatus.VerifiedNew
            result.Reason = "file output baru dibuat (" & after.Length.ToString(CultureInfo.InvariantCulture) & " byte)."
            Return result
        End If

        ' 7. File sudah ada sebelumnya -> wajib ada perubahan.
        If after.Length <> before.Length OrElse after.LastWriteUtc <> before.LastWriteUtc Then
            result.Status = SaveOutputStatus.VerifiedChanged
            result.Reason = "file output diperbarui: " &
                            before.Length.ToString(CultureInfo.InvariantCulture) & " byte / " &
                            before.LastWriteUtc.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) &
                            " UTC -> " &
                            after.Length.ToString(CultureInfo.InvariantCulture) & " byte / " &
                            after.LastWriteUtc.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) & " UTC."
            Return result
        End If

        ' 8. Tidak ada perubahan: file lama TIDAK boleh dianggap bukti sukses.
        result.Status = SaveOutputStatus.Unchanged
        result.Reason = "file output sudah ada sebelum SaveAs dan tidak berubah (ukuran & waktu " &
                        "modifikasi sama), sehingga penyimpanan terbaru tidak dapat dibuktikan."
        Return result
    End Function
End Class
