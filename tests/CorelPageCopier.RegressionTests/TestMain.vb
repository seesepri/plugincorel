Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  TestMain.vb
'  Entry point unit test Corel Page Copier.
'
'  Semua test di sini TIDAK membutuhkan CorelDRAW: logika murni diuji langsung,
'  dan WatermarkEngine diuji memakai objek COM palsu (FakeCorel.vb).
'  Integration test dengan CorelDRAW 2022 sungguhan ada di PANDUAN-UJI.md dan
'  HARUS dijalankan manual di Windows (status: NOT RUN bila tidak tersedia).
' ============================================================================

Imports System

Module TestMain

    Function Main() As Integer
        Console.WriteLine("Corel Page Copier — regression tests (tanpa CorelDRAW)")
        Console.WriteLine("Runtime: " & Environment.Version.ToString() & " / " & Environment.OSVersion.ToString())
        Console.WriteLine()

        Dim runner As New TestRunner()
        CopyVerificationTests.Register(runner)
        CopyMeasurementTests.Register(runner)
        CopyPipelineTests.Register(runner)
        DocumentUnitTests.Register(runner)
        UndoGroupTests.Register(runner)
        BatchPolicyTests.Register(runner)
        ConnectionIdentityTests.Register(runner)
        AutoSelectClassifierTests.Register(runner)
        ShapeScannerTests.Register(runner)
        SelectionNavigationTests.Register(runner)
        CorelConnectionServiceTests.Register(runner)
        CopyServiceTests.Register(runner)
        AutoSelectServiceTests.Register(runner)
        PdfProofValidatorTests.Register(runner)
        WatermarkPlacementTests.Register(runner)
        BatchLogFormatTests.Register(runner)
        BatchReliabilityTests.Register(runner)
        EngineTests.Register(runner)
        ExistingFeatureTests.Register(runner)

        Dim exitCode As Integer = runner.RunAll()
        Console.WriteLine()
        Console.WriteLine("Catatan: integration test CorelDRAW 2022 TIDAK dijalankan oleh test ini (NOT RUN).")
        Return exitCode
    End Function

End Module
