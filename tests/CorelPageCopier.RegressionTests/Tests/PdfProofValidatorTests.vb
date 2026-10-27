Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  PdfProofValidatorTests.vb
'  Fokus C — PDF proof: struktur, jumlah halaman, ukuran/orientasi.
'
'  Test di sini menjawab pertanyaan "apakah PDF-nya benar-benar bisa dibuka
'  dan sesuai dokumen sumber", bukan sekadar "file-nya ada". Yang TIDAK bisa
'  diuji di sini (watermark terlihat, opasitas, clipping) dinyatakan sebagai
'  pengujian manual di PANDUAN-UJI.md §12.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.IO

Public Module PdfProofValidatorTests

    Private Function TempFile(bytes As Byte()) As String
        Dim pdfPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                                                       "cppdf-" & Guid.NewGuid().ToString("N") & ".pdf")
        File.WriteAllBytes(pdfPath, bytes)
        Return pdfPath
    End Function

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("PDF proof — validasi struktur")

        ' --- konversi satuan ---------------------------------------------------
        suite.Add("mm -> point PDF (A4 = 595 x 842 pt)", Sub()
            Check.Near(595.28, PdfProofValidator.MillimetersToPoints(210.0), 0.05, "lebar A4")
            Check.Near(841.89, PdfProofValidator.MillimetersToPoints(297.0), 0.05, "tinggi A4")
        End Sub)

        ' --- struktur dasar ----------------------------------------------------
        suite.Add("PDF valid 3 halaman -> status Valid", Sub()
            Dim path As String = TempFile(PdfFixture.A4ThreePages())
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {210, 210, 210}, New Double() {297, 297, 297}))
                Check.AreEqual(PdfProofStatus.Valid, result.Status, "status")
                Check.IsTrue(result.HeaderOk, "header")
                Check.IsTrue(result.EofOk, "%%EOF")
                Check.IsTrue(result.StartXRefOk, "startxref")
                Check.AreEqual(3, result.PageCount, "jumlah halaman")
                Check.AreEqual(0, result.Problems.Count, "tidak ada masalah")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("file tidak ada -> Invalid (bukan sukses)", Sub()
            Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                Path.Combine(Path.GetTempPath(), "tidak-ada-" & Guid.NewGuid().ToString("N") & ".pdf"))
            Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
            Check.Contains(String.Join(" ", result.Problems.ToArray()), "tidak ditemukan", "alasan")
        End Sub)

        suite.Add("file 0 byte -> Invalid", Sub()
            Dim path As String = TempFile(New Byte() {})
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(path)
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "0 byte", "alasan")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("bukan PDF (teks biasa) -> Invalid, header dilaporkan", Sub()
            Dim path As String = TempFile(System.Text.Encoding.ASCII.GetBytes("fake cdr content"))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(path)
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.IsFalse(result.HeaderOk, "header tidak ada")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "%PDF", "alasan menyebut header")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("tanpa %%EOF (file terpotong) -> Invalid", Sub()
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {595.28}, New Double() {841.89},
                                                              Nothing, False))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(path)
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.IsFalse(result.EofOk, "%%EOF tidak ada")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "%%EOF", "alasan")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("tanpa startxref -> Invalid", Sub()
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {595.28}, New Double() {841.89},
                                                              Nothing, True, False))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(path)
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "startxref", "alasan")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        ' --- jumlah halaman ----------------------------------------------------
        suite.Add("jumlah halaman berbeda dari dokumen sumber -> Invalid", Sub()
            ' PDF 3 halaman, dokumen sumber 2 halaman.
            Dim path As String = TempFile(PdfFixture.A4ThreePages())
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {210, 210}, New Double() {297, 297}))
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "jumlah halaman", "alasan")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        ' --- ukuran & orientasi ------------------------------------------------
        suite.Add("ukuran halaman berbeda (A4 vs A5) -> Invalid", Sub()
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {595.28}, New Double() {419.53}))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {210}, New Double() {297}))
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "ukuran PDF", "alasan")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("halaman landscape cocok bila dokumen juga landscape", Sub()
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {841.89}, New Double() {595.28}))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {297}, New Double() {210}))
                Check.AreEqual(PdfProofStatus.Valid, result.Status, "status")
                Check.IsTrue(result.Pages(0).IsLandscape, "terdeteksi landscape")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("orientasi tertukar (portrait vs landscape) -> Invalid", Sub()
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {595.28}, New Double() {841.89}))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {297}, New Double() {210}))
                Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
                Check.Contains(String.Join(" ", result.Problems.ToArray()), "tertukar", "alasan")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("/Rotate 90 dengan ukuran tertukar tetap cocok", Sub()
            ' Ukuran mentah portrait + /Rotate 90 -> tampil landscape.
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {595.28}, New Double() {841.89},
                                                              New Integer() {90}))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {297}, New Double() {210}))
                Check.AreEqual(PdfProofStatus.Valid, result.Status, "status")
                Check.AreEqual(90, result.Pages(0).Rotate, "rotate terbaca")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("halaman berbeda ukuran dalam satu dokumen divalidasi per halaman", Sub()
            ' Halaman 1 A4, halaman 2 A5 landscape.
            Dim path As String = TempFile(PdfFixture.BuildPdf(
                New Double() {595.28, 595.28}, New Double() {841.89, 419.53}))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {210, 210}, New Double() {297, 148}))
                Check.AreEqual(PdfProofStatus.Valid, result.Status, "status")
                Check.AreEqual(2, result.PageCount, "dua halaman")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        ' --- keterbatasan yang TIDAK diklaim --------------------------------
        suite.Add("MediaBox tidak terbaca -> ukuran tidak diklaim cocok", Sub()
            Dim path As String = TempFile(PdfFixture.BuildPdf(New Double() {595.28}, New Double() {841.89},
                                                              Nothing, True, True, False))
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {210}, New Double() {297}))
                Check.IsFalse(result.Pages(0).SizeKnown, "ukuran ditandai tidak diketahui")
                Check.IsTrue(result.Notes.Count > 0, "catatan bahwa ukuran tidak dibandingkan")
                Check.DoesNotContain(String.Join(" ", result.Problems.ToArray()), "ukuran PDF",
                                     "tidak mengarang masalah ukuran")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("object stream -> Unverifiable (perlu cek manual), bukan Valid", Sub()
            Dim path As String = TempFile(PdfFixture.ObjectStreamPdf())
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(
                    path, PdfFixture.ExpectationMm(New Double() {210}, New Double() {297}))
                Check.AreEqual(PdfProofStatus.Unverifiable, result.Status, "status")
                Check.Contains(result.SummaryLine(), "TIDAK DAPAT DIVERIFIKASI", "ringkasan jujur")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("ringkasan valid menyebut jumlah halaman", Sub()
            Dim path As String = TempFile(PdfFixture.A4ThreePages())
            Try
                Dim result As PdfProofValidation = PdfProofValidator.ValidateFile(path)
                Check.Contains(result.SummaryLine(), "PDF VALID", "ringkasan")
                Check.Contains(result.Describe(), "Halaman    : 3", "deskripsi memuat jumlah halaman")
            Finally
                File.Delete(path)
            End Try
        End Sub)

        suite.Add("path kosong -> Invalid tanpa exception", Sub()
            Dim result As PdfProofValidation = PdfProofValidator.ValidateFile("")
            Check.AreEqual(PdfProofStatus.Invalid, result.Status, "status")
            Check.IsTrue(result.Problems.Count > 0, "alasan ada")
        End Sub)
    End Sub
End Module
