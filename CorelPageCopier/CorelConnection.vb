Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CorelConnection.vb
'  Identitas dokumen aktif + pemeriksaan kesehatan objek COM (Temuan F).
'
'  Masalah pada kode lama (ConnectToCorel):
'    - Bila instance CorelDRAW yang sedang dipakai tidak ditemukan di ROT,
'      aplikasi DIAM-DIAM membuat instance kedua. Pengguna mengira sedang
'      memproses dokumennya sendiri padahal dokumen baru/kosong.
'    - Tidak ada verifikasi bahwa dokumen target masih dokumen yang sama
'      (pengguna bisa berganti dokumen / instance di antara dua klik).
'    - Tidak ada penanganan objek COM yang sudah tidak valid (instance ditutup).
'
'  Kelas di sini murni (tanpa WinForms) sehingga bisa diuji dengan objek
'  pengganti (fake) tanpa CorelDRAW.
' ============================================================================

Imports System
Imports System.Globalization
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

''' <summary>Identitas dokumen + instance CorelDRAW yang sedang dipakai.</summary>
Public Class DocumentIdentity

    Public Name As String = ""
    Public FilePath As String = ""
    Public PageCount As Integer = -1

    ''' <summary>Penanda instance (hash dari proxy COM) agar dokumen dari instance lain tidak dianggap sama.</summary>
    Public InstanceToken As Integer = 0

    Public Sub New()
    End Sub

    Public Sub New(documentName As String, filePath As String, pageCount As Integer, instanceToken As Integer)
        Me.Name = If(documentName, "")
        Me.FilePath = If(filePath, "")
        Me.PageCount = pageCount
        Me.InstanceToken = instanceToken
    End Sub

    Public ReadOnly Property IsEmpty As Boolean
        Get
            Return String.IsNullOrEmpty(Name) AndAlso String.IsNullOrEmpty(FilePath) AndAlso PageCount < 0
        End Get
    End Property

    ''' <summary>
    ''' Kunci pembanding target. Dokumen dari instance berbeda TIDAK dianggap sama.
    ''' </summary>
    Public ReadOnly Property Key As String
        Get
            Dim path As String = If(FilePath, "").Trim()
            If path.Length > 0 Then
                Return InstanceToken.ToString(CultureInfo.InvariantCulture) & "|" & path.ToUpperInvariant()
            End If
            Return InstanceToken.ToString(CultureInfo.InvariantCulture) & "|" &
                   If(Name, "").ToUpperInvariant() & "|" & PageCount.ToString(CultureInfo.InvariantCulture)
        End Get
    End Property

    ''' <summary>True bila instance + dokumen sama (jumlah halaman boleh berubah).</summary>
    Public Function SameTarget(other As DocumentIdentity) As Boolean
        If other Is Nothing Then Return False
        If InstanceToken <> other.InstanceToken Then Return False
        Dim leftPath As String = If(FilePath, "").Trim()
        Dim rightPath As String = If(other.FilePath, "").Trim()
        If leftPath.Length > 0 AndAlso rightPath.Length > 0 Then
            Return String.Equals(leftPath, rightPath, StringComparison.OrdinalIgnoreCase)
        End If
        Return String.Equals(If(Name, ""), If(other.Name, ""), StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>Deskripsi untuk status bar / dialog konfirmasi.</summary>
    Public Function Describe() As String
        Dim label As String = If(FilePath, "").Trim()
        If label.Length = 0 Then label = If(Name, "").Trim()
        If label.Length = 0 Then label = "(dokumen tanpa nama)"
        Dim pages As String = If(PageCount >= 0,
                                 PageCount.ToString(CultureInfo.InvariantCulture) & " halaman",
                                 "jumlah halaman tidak terbaca")
        Return label & " — " & pages
    End Function

    ''' <summary>
    ''' Penanda instance CorelDRAW. Prioritas: handle jendela aplikasi (stabil per
    ''' proses), lalu hash proxy COM (best effort).
    ''' </summary>
    Public Shared Function InstanceTokenFor(app As Object) As Integer
        If app Is Nothing Then Return 0
        Try
            Dim handle As Integer = CInt(app.AppWindow.Handle)
            If handle <> 0 Then Return handle
        Catch
        End Try
        Try
            Return RuntimeHelpers.GetHashCode(app)
        Catch
        End Try
        Return 0
    End Function

    ''' <summary>
    ''' Bangun identitas dari objek Document (late bound). Tidak pernah melempar
    ''' exception; bagian yang gagal dibaca dibiarkan kosong.
    ''' </summary>
    Public Shared Function FromDocument(doc As Object, instanceToken As Integer) As DocumentIdentity
        Dim identity As New DocumentIdentity()
        identity.InstanceToken = instanceToken
        If doc Is Nothing Then
            identity.PageCount = -1
            Return identity
        End If

        Try
            identity.Name = CStr(doc.Name)
        Catch
        End Try
        For Each propertyName As String In New String() {"FilePath", "FileName", "FullFileName"}
            If Not String.IsNullOrEmpty(identity.FilePath) Then Exit For
            Try
                Dim value As String = ""
                Select Case propertyName
                    Case "FilePath" : value = CStr(doc.FilePath)
                    Case "FileName" : value = CStr(doc.FileName)
                    Case Else : value = CStr(doc.FullFileName)
                End Select
                If Not String.IsNullOrEmpty(value) Then identity.FilePath = value
            Catch
            End Try
        Next
        Try
            identity.PageCount = CInt(doc.Pages.Count)
        Catch
            identity.PageCount = -1
        End Try
        Return identity
    End Function
End Class

''' <summary>Pemeriksaan kesehatan objek COM late bound.</summary>
Public Class ComObjectHealth

    ''' <summary>
    ''' False bila objek COM sudah dilepas / instance-nya sudah ditutup.
    ''' Objek non-COM dianggap hidup (memudahkan unit test dengan fake).
    ''' </summary>
    Public Shared Function IsAlive(obj As Object) As Boolean
        If obj Is Nothing Then Return False
        Try
            If Marshal.IsComObject(obj) Then
                ' Membaca properti apa pun akan melempar COMException bila server
                ' sudah mati, dan InvalidComObjectException bila RCW dilepas.
                Dim probe As Object = obj.Name
                Return True
            End If
            Return True
        Catch ex As InvalidComObjectException
            Return False
        Catch ex As COMException
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>Lepas satu referensi COM; kegagalan tidak fatal.</summary>
    Public Shared Function TryRelease(obj As Object) As Boolean
        If obj Is Nothing Then Return True
        Try
            If Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
            Return True
        Catch
            Return False
        End Try
    End Function
End Class
