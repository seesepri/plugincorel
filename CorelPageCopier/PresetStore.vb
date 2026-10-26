Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  PresetStore.vb
'  Penyimpanan preset watermark di %APPDATA%\CorelPageCopier\presets.ini
'
'  Format sengaja dibuat INI (key=value) alih-alih JSON:
'    - tanpa dependensi tambahan (.NET Framework 4.8 tidak punya System.Text.Json)
'    - bisa dibaca/diedit manusia, aman kalau rusak sebagian (baris tak dikenal diabaikan)
'    - seluruh logika serialize/parse murni -> bisa diuji unit tanpa Windows
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Text

''' <summary>Satu preset: nama + seluruh pengaturan watermark.</summary>
Public Class WatermarkPreset
    Public Name As String = ""
    Public Settings As New WatermarkSettings()

    Public Sub New()
    End Sub

    Public Sub New(presetName As String, source As WatermarkSettings)
        Name = If(presetName, "")
        If source IsNot Nothing Then Settings = source.Clone()
    End Sub
End Class

''' <summary>Baca/tulis daftar preset ke file di %APPDATA%.</summary>
Public Class PresetStore

    Public Const FileVersion As String = "v1"
    Public Const FolderName As String = "CorelPageCopier"
    Public Const FileName As String = "presets.ini"

    Private ReadOnly _filePath As String

    ''' <summary>path = Nothing -> %APPDATA%\CorelPageCopier\presets.ini</summary>
    Public Sub New(Optional filePath As String = Nothing)
        _filePath = If(String.IsNullOrEmpty(filePath), DefaultFilePath, filePath)
    End Sub

    Public ReadOnly Property FilePath As String
        Get
            Return _filePath
        End Get
    End Property

    ''' <summary>Folder aplikasi di %APPDATA%. Fallback ke folder temp bila %APPDATA% tidak ada.</summary>
    Public Shared Function DefaultDirectory() As String
        Dim root As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        If String.IsNullOrEmpty(root) Then root = Path.GetTempPath()
        Return Path.Combine(root, FolderName)
    End Function

    Public Shared Function DefaultFilePath() As String
        Return Path.Combine(DefaultDirectory(), FileName)
    End Function

    ' ========================================================================
    '  LOAD / SAVE
    ' ========================================================================

    Public Function Load() As List(Of WatermarkPreset)
        Try
            If Not File.Exists(_filePath) Then Return New List(Of WatermarkPreset)()
            Return Deserialize(File.ReadAllText(_filePath, Encoding.UTF8))
        Catch
            ' File rusak / tidak terbaca tidak boleh menjatuhkan aplikasi.
            Return New List(Of WatermarkPreset)()
        End Try
    End Function

    ''' <summary>Simpan atomik: tulis ke .tmp lalu pindahkan, agar file tidak korup saat gagal.</summary>
    Public Function Save(presets As List(Of WatermarkPreset)) As Boolean
        Try
            Dim folder As String = Path.GetDirectoryName(_filePath)
            If Not String.IsNullOrEmpty(folder) Then Directory.CreateDirectory(folder)

            Dim tempPath As String = _filePath & ".tmp"
            File.WriteAllText(tempPath, Serialize(presets), New UTF8Encoding(False))

            If File.Exists(_filePath) Then File.Delete(_filePath)
            File.Move(tempPath, _filePath)
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>Tambah/ganti preset berdasarkan nama (case-insensitive).</summary>
    Public Shared Function Upsert(presets As List(Of WatermarkPreset), name As String,
                                  settings As WatermarkSettings) As Boolean
        If presets Is Nothing OrElse String.IsNullOrEmpty(name) Then Return False
        Dim replaced As Boolean = False
        For Each existing As WatermarkPreset In presets
            If String.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase) Then
                existing.Settings = settings.Clone()
                replaced = True
                Exit For
            End If
        Next
        If Not replaced Then presets.Add(New WatermarkPreset(name, settings))
        Return replaced
    End Function

    Public Shared Function Remove(presets As List(Of WatermarkPreset), name As String) As Boolean
        If presets Is Nothing Then Return False
        For i As Integer = presets.Count - 1 To 0 Step -1
            If String.Equals(presets(i).Name, name, StringComparison.OrdinalIgnoreCase) Then
                presets.RemoveAt(i)
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Function Find(presets As List(Of WatermarkPreset), name As String) As WatermarkPreset
        If presets Is Nothing OrElse name Is Nothing Then Return Nothing
        For Each preset As WatermarkPreset In presets
            If String.Equals(preset.Name, name, StringComparison.OrdinalIgnoreCase) Then Return preset
        Next
        Return Nothing
    End Function

    ' ========================================================================
    '  SERIALISASI (murni)
    ' ========================================================================

    Public Shared Function Serialize(presets As List(Of WatermarkPreset)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("; Corel Page Copier - watermark presets (" & FileVersion & ")")
        sb.AppendLine("; Aman diedit manual. Baris tak dikenal akan diabaikan.")
        If presets IsNot Nothing Then
            For Each preset As WatermarkPreset In presets
                Dim s As WatermarkSettings = preset.Settings
                sb.AppendLine()
                sb.AppendLine("[preset " & EscapeValue(preset.Name) & "]")
                AppendPair(sb, "Kind", CInt(s.Kind).ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "Text", s.Text)
                AppendPair(sb, "FontName", s.FontName)
                AppendPair(sb, "FontSize", s.FontSize.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "ColorArgb", s.Color.ToArgb().ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "UseBundledLogo", If(s.UseBundledLogo, "1", "0"))
                AppendPair(sb, "BundledLogoWhite", If(s.BundledLogoWhite, "1", "0"))
                AppendPair(sb, "LogoPath", s.LogoPath)
                AppendPair(sb, "LogoWidthMm", s.LogoWidthMm.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "OpacityPercent", s.OpacityPercent.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "RotationDeg", s.RotationDeg.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "Anchor", CInt(s.Anchor).ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "MarginMm", s.MarginMm.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "Tile", If(s.Tile, "1", "0"))
                AppendPair(sb, "TileSpacingX", s.TileSpacingXMm.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "TileSpacingY", s.TileSpacingYMm.ToString(CultureInfo.InvariantCulture))
                AppendPair(sb, "AllPages", If(s.AllPages, "1", "0"))
                AppendPair(sb, "TargetPages", JoinPages(s.TargetPages))
                AppendPair(sb, "ReplaceExisting", If(s.ReplaceExisting, "1", "0"))
                AppendPair(sb, "LayerName", s.LayerName)
                AppendPair(sb, "PerPageText", If(s.PerPageText, "1", "0"))
                AppendPair(sb, "LockLayer", If(s.LockLayer, "1", "0"))
                AppendPair(sb, "ExportProofPdf", If(s.ExportProofPdf, "1", "0"))
            Next
        End If
        Return sb.ToString()
    End Function

    Public Shared Function Deserialize(text As String) As List(Of WatermarkPreset)
        Dim result As New List(Of WatermarkPreset)()
        If String.IsNullOrEmpty(text) Then Return result

        Dim current As WatermarkPreset = Nothing
        For Each rawLine As String In text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(ChrW(10))
            Dim line As String = rawLine.Trim()
            If line.Length = 0 OrElse line.StartsWith(";") OrElse line.StartsWith("#") Then Continue For

            If line.StartsWith("[preset") AndAlso line.EndsWith("]") Then
                Dim inner As String = line.Substring("[preset".Length, line.Length - "[preset".Length - 1).Trim()
                current = New WatermarkPreset()
                current.Name = UnescapeValue(inner)
                result.Add(current)
                Continue For
            End If

            If current Is Nothing Then Continue For
            Dim eq As Integer = line.IndexOf("="c)
            If eq <= 0 Then Continue For

            Dim key As String = line.Substring(0, eq).Trim()
            Dim value As String = line.Substring(eq + 1).Trim()
            ApplyKey(current.Settings, key, value)
        Next

        ' Buang preset tanpa nama (mis. hasil edit manual yang rusak).
        For i As Integer = result.Count - 1 To 0 Step -1
            If String.IsNullOrEmpty(result(i).Name) Then result.RemoveAt(i)
        Next
        Return result
    End Function

    Private Shared Sub ApplyKey(s As WatermarkSettings, key As String, value As String)
        Select Case key.ToLowerInvariant()
            Case "kind" : s.Kind = CType(ToInt(value, CInt(s.Kind)), WatermarkKind)
            Case "text" : s.Text = UnescapeValue(value)
            Case "fontname" : s.FontName = UnescapeValue(value)
            Case "fontsize" : s.FontSize = CSng(ToDouble(value, s.FontSize))
            Case "colorargb" : s.Color = Color.FromArgb(ToInt(value, s.Color.ToArgb()))
            Case "usebundledlogo" : s.UseBundledLogo = ToBool(value, s.UseBundledLogo)
            Case "bundledlogowhite" : s.BundledLogoWhite = ToBool(value, s.BundledLogoWhite)
            Case "logopath" : s.LogoPath = UnescapeValue(value)
            Case "logowidthmm" : s.LogoWidthMm = ToDouble(value, s.LogoWidthMm)
            Case "opacitypercent" : s.OpacityPercent = ToInt(value, s.OpacityPercent)
            Case "rotationdeg" : s.RotationDeg = ToDouble(value, s.RotationDeg)
            Case "anchor" : s.Anchor = CType(ToInt(value, CInt(s.Anchor)), WatermarkAnchor)
            Case "marginmm" : s.MarginMm = ToDouble(value, s.MarginMm)
            Case "tile" : s.Tile = ToBool(value, s.Tile)
            Case "tilespacingx" : s.TileSpacingXMm = ToDouble(value, s.TileSpacingXMm)
            Case "tilespacingy" : s.TileSpacingYMm = ToDouble(value, s.TileSpacingYMm)
            Case "allpages" : s.AllPages = ToBool(value, s.AllPages)
            Case "targetpages"
                s.TargetPages.Clear()
                For Each part As String In value.Split(","c)
                    Dim pageIndex As Integer = ToInt(part.Trim(), 0)
                    If pageIndex > 0 AndAlso Not s.TargetPages.Contains(pageIndex) Then s.TargetPages.Add(pageIndex)
                Next
            Case "replaceexisting" : s.ReplaceExisting = ToBool(value, s.ReplaceExisting)
            Case "layername" : s.LayerName = UnescapeValue(value)
            Case "perpagetext" : s.PerPageText = ToBool(value, s.PerPageText)
            Case "locklayer" : s.LockLayer = ToBool(value, s.LockLayer)
            Case "exportproofpdf" : s.ExportProofPdf = ToBool(value, s.ExportProofPdf)
        End Select
    End Sub

    ' --- helper konversi toleran -------------------------------------------
    Private Shared Function ToInt(value As String, fallback As Integer) As Integer
        Dim parsed As Integer
        If Integer.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then Return parsed
        Return fallback
    End Function

    Private Shared Function ToDouble(value As String, fallback As Double) As Double
        Dim parsed As Double
        If Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, parsed) Then Return parsed
        Return fallback
    End Function

    Private Shared Function ToBool(value As String, fallback As Boolean) As Boolean
        Select Case value.Trim().ToLowerInvariant()
            Case "1", "true", "yes", "ya" : Return True
            Case "0", "false", "no", "tidak" : Return False
        End Select
        Return fallback
    End Function

    Private Shared Function JoinPages(pages As List(Of Integer)) As String
        If pages Is Nothing OrElse pages.Count = 0 Then Return ""
        Dim parts As New List(Of String)()
        For Each pageIndex As Integer In pages
            parts.Add(pageIndex.ToString(CultureInfo.InvariantCulture))
        Next
        Return String.Join(",", parts.ToArray())
    End Function

    ' Newline di-escape supaya satu nilai tetap satu baris.
    Private Shared Function EscapeValue(value As String) As String
        If value Is Nothing Then Return ""
        Return value.Replace("\", "\\").Replace(vbCrLf, "\n").Replace(vbLf, "\n").Replace(vbCr, "\n")
    End Function

    Private Shared Function UnescapeValue(value As String) As String
        If value Is Nothing Then Return ""
        Dim sb As New StringBuilder()
        Dim i As Integer = 0
        While i < value.Length
            Dim c As Char = value(i)
            If c = "\"c AndAlso i + 1 < value.Length Then
                Dim n As Char = value(i + 1)
                If n = "n"c Then
                    sb.Append(vbLf)
                Else
                    sb.Append(n)
                End If
                i += 2
            Else
                sb.Append(c)
                i += 1
            End If
        End While
        Return sb.ToString()
    End Function

    Private Shared Sub AppendPair(sb As StringBuilder, key As String, value As String)
        sb.AppendLine(key & "=" & EscapeValue(value))
    End Sub
End Class
