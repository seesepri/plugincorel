Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  CopyVerificationTests.vb
'  Regression test untuk Temuan A (verifikasi copy):
'  Verified / Partial / Failed dipisah, dan Partial tidak pernah dihitung
'  sebagai sukses penuh.
' ============================================================================

Imports System

Public Module CopyVerificationTests

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Verifikasi copy (Temuan A)")

        ' --- objek tunggal -------------------------------------------------
        suite.Add("objek tunggal cocok -> Verified", Sub()
                                                        Dim delta = WatermarkLayout.CompareGeometry(100, 100, 50, 30, 100.2, 99.8, 50.2, 30.1)
                                                        Dim result = CopyVerification.Classify(1, True, delta, True, True)
                                                        Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status")
                                                    End Sub)

        ' --- multi selection ----------------------------------------------
        suite.Add("multi-selection (3 objek) cocok -> Verified", Sub()
                                                                     Dim delta = WatermarkLayout.CompareGeometry(0, 0, 100, 100, 0.1, -0.1, 100.1, 99.9)
                                                                     Dim result = CopyVerification.Classify(3, True, delta, True, True)
                                                                     Check.AreEqual(CopyVerifyStatus.Verified, result.Status, "status")
                                                                 End Sub)

        ' --- grup bersarang: 1 shape grup, struktur sama -------------------
        suite.Add("grup bersarang (1 objek tipe 7) -> Verified", Sub()
                                                                     Dim signature As ShapeSignature = ShapeSignature.FromTypeCodes(New Integer() {7})
                                                                     Check.IsTrue(signature.Matches(ShapeSignature.FromTypeCodes(New Integer() {7})), "struktur grup cocok")
                                                                     Dim delta = WatermarkLayout.CompareGeometry(50, 50, 20, 20, 50, 50, 20, 20)
                                                                     Check.AreEqual(CopyVerifyStatus.Verified,
                                                                                    CopyVerification.Classify(1, True, delta, True, True).Status, "status")
                                                                 End Sub)

        ' --- bitmap & teks: geometri beda -> Failed -----------------------
        suite.Add("bitmap/teks dengan posisi berbeda -> Failed", Sub()
                                                                    Dim delta = WatermarkLayout.CompareGeometry(0, 0, 10, 10, 3, 3, 11, 11)
                                                                    Dim result = CopyVerification.Classify(1, True, delta, True, True)
                                                                    Check.AreEqual(CopyVerifyStatus.Failed, result.Status, "status")
                                                                    Check.Contains(result.Reason, "berbeda", "alasan menyebut perbedaan geometri")
                                                                End Sub)

        ' --- struktur berubah (mis. 3 objek jadi 1 grup) -> Partial -------
        suite.Add("geometri cocok tetapi struktur berbeda -> Partial", Sub()
                                                                          Dim delta = WatermarkLayout.CompareGeometry(0, 0, 100, 100, 0, 0, 100, 100)
                                                                          Dim result = CopyVerification.Classify(3, True, delta, True, False)
                                                                          Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "status")
                                                                      End Sub)

        ' --- ukuran campuran: toleransi 0,5 mm ----------------------------
        suite.Add("ukuran campuran tepat di batas toleransi -> Verified", Sub()
                                                                             Dim delta = WatermarkLayout.CompareGeometry(0, 0, 100, 100, 0.5, -0.5, 100.5, 99.5)
                                                                             Check.AreEqual(CopyVerifyStatus.Verified,
                                                                                            CopyVerification.Classify(2, True, delta, True, True).Status,
                                                                                            "status pada batas toleransi")
                                                                         End Sub)

        suite.Add("ukuran campuran melebihi toleransi -> Failed", Sub()
                                                                     Dim delta = WatermarkLayout.CompareGeometry(0, 0, 100, 100, 0.51, 0, 100, 100)
                                                                     Check.AreEqual(CopyVerifyStatus.Failed,
                                                                                    CopyVerification.Classify(2, True, delta, True, True).Status,
                                                                                    "status di atas toleransi")
                                                                 End Sub)

        ' --- geometri TIDAK tersedia -> Partial, BUKAN Verified -----------
        suite.Add("geometri tidak terukur -> Partial (bukan Verified)", Sub()
                                                                           Dim delta As New WatermarkLayout.GeometryDelta()
                                                                           Dim result = CopyVerification.Classify(2, False, delta, True, True)
                                                                           Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "status")
                                                                           Check.IsFalse(result.IsVerified, "tidak boleh dianggap verified")
                                                                           Check.Contains(result.Reason, "geometri tidak dapat diukur", "alasan eksplisit")
                                                                       End Sub)

        suite.Add("geometri & struktur tidak terukur -> Partial", Sub()
                                                                     Dim result = CopyVerification.Classify(1, False, New WatermarkLayout.GeometryDelta(), False, False)
                                                                     Check.AreEqual(CopyVerifyStatus.PartiallyVerified, result.Status, "status")
                                                                 End Sub)

        ' --- tidak ada objek hasil paste -> Failed ------------------------
        suite.Add("paste tidak menambah objek -> Failed", Sub()
                                                             Dim result = CopyVerification.Classify(0, False, New WatermarkLayout.GeometryDelta(), False, False)
                                                             Check.AreEqual(CopyVerifyStatus.Failed, result.Status, "status")
                                                         End Sub)

        ' --- ringkasan: Partial tidak dihitung sebagai sukses -------------
        suite.Add("CopyRunSummary: partial tidak masuk hitungan verified", Sub()
                                                                              Dim summary As New CopyRunSummary()
                                                                              summary.Add(2, CopyVerifyStatus.Verified, "ok")
                                                                              summary.Add(3, CopyVerifyStatus.PartiallyVerified, "geometri tak terukur")
                                                                              summary.Add(4, CopyVerifyStatus.Failed, "posisi beda")
                                                                              Check.AreEqual(1, summary.VerifiedCount, "verified")
                                                                              Check.AreEqual(1, summary.PartialCount, "partial")
                                                                              Check.AreEqual(1, summary.FailedCount, "failed")
                                                                              Check.IsFalse(summary.AllVerified, "tidak semua verified")
                                                                              Dim report As String = summary.BuildReport("uji.cdr — 3 halaman", "satuan mm", "undo ok", 1.5)
                                                                              Check.Contains(report, "Partial       : 1", "laporan memisahkan Partial")
                                                                              Check.Contains(report, "Halaman 3 [Partial]", "halaman partial dirinci")
                                                                              Check.IsFalse(report.Contains("1/3 halaman terverifikasi"),
                                                                                            "tidak ada klaim sukses untuk halaman partial")
                                                                          End Sub)

        suite.Add("CopyRunSummary: semua verified -> AllVerified True", Sub()
                                                                           Dim summary As New CopyRunSummary()
                                                                           summary.Add(2, CopyVerifyStatus.Verified, "ok")
                                                                           Check.IsTrue(summary.AllVerified, "semua verified")
                                                                           Check.Contains(summary.Headline(0.5), "1 verified, 0 partial, 0 gagal", "headline")
                                                                       End Sub)

        ' --- kotak pembatas -------------------------------------------------
        suite.Add("BoundingBox: gabungan dua kotak", Sub()
                                                          Dim first As BoundingBox = BoundingBox.FromCenter(0, 0, 10, 10)
                                                          Dim second As BoundingBox = BoundingBox.FromCenter(20, 0, 10, 10)
                                                          Dim merged As BoundingBox = first.Union(second)
                                                          Check.Near(-5, merged.MinX, 0.0001, "minX")
                                                          Check.Near(25, merged.MaxX, 0.0001, "maxX")
                                                          Check.Near(10, merged.CenterX, 0.0001, "centerX")
                                                          Check.Near(30, merged.Width, 0.0001, "width")
                                                      End Sub)

        suite.Add("BoundingBox: skala konversi satuan", Sub()
                                                            Dim box As BoundingBox = BoundingBox.FromCenter(0, 0, 1, 1)
                                                            Dim mm As BoundingBox = box.Scale(25.4)
                                                            Check.Near(25.4, mm.Width, 0.0001, "1 inch -> 25.4 mm")
                                                        End Sub)

        ' --- sidik jari struktur -------------------------------------------
        suite.Add("ShapeSignature: tipe berbeda tidak cocok", Sub()
                                                                  Dim left As ShapeSignature = ShapeSignature.FromTypeCodes(New Integer() {6, 5})
                                                                  Dim right As ShapeSignature = ShapeSignature.FromTypeCodes(New Integer() {6, 3})
                                                                  Check.IsFalse(left.Matches(right), "tipe berbeda")
                                                              End Sub)

        suite.Add("ShapeSignature: jumlah berbeda tidak cocok", Sub()
                                                                    Check.IsFalse(ShapeSignature.FromTypeCodes(New Integer() {6}).Matches(
                                                                                  ShapeSignature.FromTypeCodes(New Integer() {6, 6})), "jumlah berbeda")
                                                                End Sub)
    End Sub
End Module
