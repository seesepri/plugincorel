Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  ExistingFeatureTests.vb
'  Regression test untuk fitur yang sudah ada (token, log CSV, preset,
'  tata letak watermark, OperationGate) supaya perbaikan audit tidak
'  merusak perilaku lama.
' ============================================================================

Imports System
Imports System.Collections.Generic
Imports System.Drawing

Public Module ExistingFeatureTests

    Public Sub Register(runner As TestRunner)
        Dim suite As TestSuite = runner.AddSuite("Fitur lama (tidak boleh rusak)")

        ' --- token ----------------------------------------------------------
        suite.Add("token {page}/{pages}/{order} dievaluasi", Sub()
            Dim ctx As New WatermarkTokenContext()
            ctx.PageIndex = 3
            ctx.PageCount = 9
            ctx.FileName = "C:\x\PO_2026_0042.cdr"
            Check.AreEqual("Hal 3/9 PO_2026_0042",
                           WatermarkTokens.Expand("Hal {page}/{pages} {order}", ctx), "hasil expand")
            Check.IsTrue(WatermarkTokens.HasToken("Hal {page}", "page"), "HasToken")
            Check.AreEqual("a {unknown} b", WatermarkTokens.Expand("a {unknown} b", ctx), "token asing dibiarkan")
        End Sub)

        ' --- log CSV --------------------------------------------------------
        suite.Add("CSV escape & header", Sub()
            Check.AreEqual("""a;b""", BatchLog.EscapeCell("a;b"), "titik koma di-escape")
            Check.Contains(BatchLog.Header(), "Status", "header memuat Status")
            Dim entry As New BatchLogEntry()
            entry.FileName = "a.cdr"
            entry.Status = "OK"
            Check.Contains(BatchLog.BuildLine(entry), ";OK;", "status masuk baris")
        End Sub)

        ' --- preset ---------------------------------------------------------
        suite.Add("preset: serialize & deserialize bolak-balik", Sub()
            Dim settings As New WatermarkSettings()
            settings.Text = "RAHASIA"
            settings.AllPages = False
            settings.TargetPages.Add(2)
            settings.TargetPages.Add(4)
            settings.OutputSuffix = "_v2"

            Dim presets As New List(Of WatermarkPreset)()
            PresetStore.Upsert(presets, "uji", settings)
            Dim roundTrip As List(Of WatermarkPreset) = PresetStore.Deserialize(PresetStore.Serialize(presets))

            Check.AreEqual(1, roundTrip.Count, "satu preset")
            Check.AreEqual("uji", roundTrip(0).Name, "nama preset")
            Check.AreEqual("RAHASIA", roundTrip(0).Settings.Text, "teks")
            Check.AreEqual(2, roundTrip(0).Settings.TargetPages.Count, "halaman target")
            Check.AreEqual("_v2", roundTrip(0).Settings.OutputSuffix, "akhiran output tersimpan")
        End Sub)

        ' --- tata letak -----------------------------------------------------
        suite.Add("tata letak: anchor tengah & tile", Sub()
            Dim point As PointF = WatermarkLayout.AnchorLeftBottom(200, 300, 100, 50, WatermarkAnchor.Center, 5)
            Check.Near(50, point.X, 0.01, "kiri tengah")
            Check.Near(125, point.Y, 0.01, "bawah tengah")
            Dim positions As List(Of PointF) = WatermarkLayout.TilePositions(200, 300, 50, 20, 60, 45)
            Check.IsTrue(positions.Count > 1, "tile menghasilkan banyak posisi")
            Check.AreEqual(0, WatermarkLayout.TilePositions(0, 300, 50, 20, 60, 45).Count, "halaman tidak valid -> kosong")
        End Sub)

        suite.Add("transparency: opasitas 50% -> 50", Sub()
            Check.AreEqual(50, WatermarkLayout.TransparencyFromOpacity(50), "50%")
            Check.AreEqual(0, WatermarkLayout.TransparencyFromOpacity(150), "dibatasi 100")
        End Sub)

        suite.Add("identitas watermark WM_ / layer WATERMARK", Sub()
            Check.IsTrue(WatermarkLayout.IsWatermarkObjectName("WM_TEXT"), "prefix WM_")
            Check.IsTrue(WatermarkLayout.IsWatermarkLayerName("watermark"), "layer case-insensitive")
            Check.IsFalse(WatermarkLayout.IsWatermarkObjectName("Logo klien"), "objek biasa")
        End Sub)

        suite.Add("settings.Validate menolak teks kosong", Sub()
            Dim settings As New WatermarkSettings()
            settings.Text = "  "
            Check.IsTrue(settings.Validate().Count > 0, "validasi gagal")
        End Sub)

        ' --- OperationGate ---------------------------------------------------
        suite.Add("OperationGate mencegah re-entrancy", Sub()
            Dim gate As New OperationGate()
            Check.IsTrue(gate.TryBegin("uji"), "kunci pertama")
            Check.IsFalse(gate.TryBegin("uji2"), "kunci kedua ditolak")
            Check.Contains(gate.BusyMessage, "uji", "pesan sibuk")
            gate.Release()
            Check.IsTrue(gate.TryBegin("uji3"), "setelah release boleh lagi")
        End Sub)
    End Sub
End Module
