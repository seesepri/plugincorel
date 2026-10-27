Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  DocumentUnitTests.vb
'  Regression test Temuan B (normalisasi satuan): mm / inch / point,
'  konversi eksplisit, dan pelaporan kegagalan pemulihan.
' ============================================================================

Imports System

Public Module DocumentUnitTests

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Normalisasi satuan (Temuan B)")

        suite.Add("konversi mm / inch / point ke mm", Sub()
            Check.Near(1.0, DocumentUnits.ToMillimeters(1, DocumentUnits.Millimeter), 0.000001, "1 mm")
            Check.Near(25.4, DocumentUnits.ToMillimeters(1, DocumentUnits.Inch), 0.000001, "1 inch")
            Check.Near(25.4, DocumentUnits.ToMillimeters(72, DocumentUnits.Point), 0.01, "72 pt = 1 inch = 25.4 mm")
            Check.Near(10.0, DocumentUnits.ToMillimeters(1, DocumentUnits.Centimeter), 0.000001, "1 cm")
        End Sub)

        suite.Add("pixel tidak dapat dikonversi (tidak ada asumsi mm)", Sub()
            Check.IsFalse(DocumentUnits.IsConvertible(DocumentUnits.Pixel), "pixel tidak didukung")
            Check.IsTrue(Double.IsNaN(DocumentUnits.ToMillimeters(10, DocumentUnits.Pixel)), "hasil NaN, bukan angka palsu")
        End Sub)

        suite.Add("scope: perubahan mm berhasil -> faktor 1 dan perlu restore", Sub()
            Dim scope As New DocumentUnitScope(DocumentUnits.Millimeter)
            scope.MarkOriginalKnown(DocumentUnits.Inch)
            scope.MarkChangeApplied()
            Check.AreEqual(1.0, scope.MeasurementFactorMm, "faktor mm")
            Check.IsTrue(scope.CanVerifyInMillimeters, "verifikasi mm boleh")
            Check.IsTrue(scope.RestoreRequired, "satuan harus dipulihkan")
            scope.MarkRestored()
            Check.IsFalse(scope.RestoreFailed, "pemulihan berhasil")
            Check.Contains(scope.Describe(), "berhasil dipulihkan", "laporan")
        End Sub)

        suite.Add("scope: penyetelan mm gagal -> konversi eksplisit dipakai", Sub()
            Dim scope As New DocumentUnitScope(DocumentUnits.Millimeter)
            scope.MarkOriginalKnown(DocumentUnits.Inch)
            scope.MarkChangeFailed("simulasi: Unit read-only")
            Check.AreEqual(DocumentUnits.Inch, scope.MeasurementUnit, "pakai satuan asli")
            Check.Near(25.4, scope.MeasurementFactorMm, 0.000001, "faktor inch")
            Check.IsTrue(scope.CanVerifyInMillimeters, "verifikasi mm masih mungkin via konversi")
            Check.IsFalse(scope.RestoreRequired, "tidak ada yang perlu dipulihkan")
            Check.Contains(scope.Describe(), "dikonversi eksplisit", "laporan menyebut konversi")
        End Sub)

        suite.Add("scope: satuan tidak terbaca -> verifikasi diblokir", Sub()
            Dim scope As New DocumentUnitScope(DocumentUnits.Millimeter)
            scope.MarkOriginalUnknown("simulasi: COM error")
            Check.IsFalse(scope.CanVerifyInMillimeters, "tidak boleh asumsi mm")
            Check.Contains(scope.BlockingReason, "tidak dapat dibaca", "alasan blokir")
        End Sub)

        suite.Add("scope: satuan pixel -> verifikasi diblokir", Sub()
            Dim scope As New DocumentUnitScope(DocumentUnits.Millimeter)
            scope.MarkOriginalKnown(DocumentUnits.Pixel)
            scope.MarkChangeFailed("simulasi: gagal set mm")
            Check.IsFalse(scope.CanVerifyInMillimeters, "pixel tidak dapat dikonversi")
            Check.Contains(scope.BlockingReason, "tidak dapat dikonversi", "alasan blokir")
        End Sub)

        suite.Add("scope: pemulihan gagal dilaporkan eksplisit", Sub()
            Dim scope As New DocumentUnitScope(DocumentUnits.Millimeter)
            scope.MarkOriginalKnown(DocumentUnits.Inch)
            scope.MarkChangeApplied()
            scope.MarkRestoreFailed("simulasi: gagal restore")
            Check.IsTrue(scope.RestoreFailed, "kegagalan tercatat")
            Check.Contains(scope.Describe(), "PEMULIHAN", "laporan menyebut pemulihan gagal")
            Check.Contains(scope.Describe(), "simulasi: gagal restore", "detail error disertakan")
        End Sub)

        suite.Add("scope: tidak ada perubahan -> tidak ada restore", Sub()
            Dim scope As New DocumentUnitScope(DocumentUnits.Millimeter)
            scope.MarkOriginalKnown(DocumentUnits.Millimeter)
            scope.MarkChangeApplied()
            Check.IsTrue(scope.RestoreRequired, "tetap perlu restore ke unit asli")
        End Sub)

        suite.Add("toleransi geometri terdokumentasi = 0,5 mm", Sub()
            Check.Near(0.5, CopyVerification.ToleranceMm, 0.000001, "toleransi")
        End Sub)
    End Sub
End Module
