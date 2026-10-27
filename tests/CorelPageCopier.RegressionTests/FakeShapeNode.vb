Option Explicit On
Option Strict Off
Option Infer On

' ============================================================================
'  FakeShapeNode.vb
'  Node objek palsu untuk menguji ShapeClassifier / ShapeScanner / MatchNavigator
'  tanpa CorelDRAW. Meniru struktur ShapeNodeAdapter (grup, PowerClip, error baca).
' ============================================================================

Imports System
Imports System.Collections.Generic

Public Class FakeShapeNode
    Implements IShapeNode

    Public Facts As New ShapeFacts()
    Public ChildNodes As New List(Of IShapeNode)()
    Public PowerClipNodes As New List(Of IShapeNode)()

    Public Property ChildrenErrorMessage As String = ""
    Public Property IsPowerClip As Boolean = False
    Public Property PowerClipErrorMessage As String = ""
    Public Property ThrowOnGetFacts As Boolean = False
    ''' <summary>Objek yang dikembalikan sebagai target seleksi (default: dirinya sendiri).</summary>
    Public Property SelectionTargetOverride As Object = Nothing
    Public Property Marker As String = ""

    Public Sub New()
    End Sub

    Public Function GetFacts() As ShapeFacts Implements IShapeNode.GetFacts
        If ThrowOnGetFacts Then Throw New InvalidOperationException("simulasi: fakta objek tidak terbaca")
        Return Facts
    End Function

    Public Function GetChildren() As IList(Of IShapeNode) Implements IShapeNode.GetChildren
        Return ChildNodes
    End Function

    Public ReadOnly Property ChildrenError As String Implements IShapeNode.ChildrenError
        Get
            Return ChildrenErrorMessage
        End Get
    End Property

    Public ReadOnly Property IsPowerClipContainer As Boolean Implements IShapeNode.IsPowerClipContainer
        Get
            Return IsPowerClip
        End Get
    End Property

    Public Function GetPowerClipChildren() As IList(Of IShapeNode) Implements IShapeNode.GetPowerClipChildren
        Return PowerClipNodes
    End Function

    Public ReadOnly Property PowerClipError As String Implements IShapeNode.PowerClipError
        Get
            Return PowerClipErrorMessage
        End Get
    End Property

    Public ReadOnly Property SelectionTarget As Object Implements IShapeNode.SelectionTarget
        Get
            If SelectionTargetOverride IsNot Nothing Then Return SelectionTargetOverride
            Return Me
        End Get
    End Property

    Public ReadOnly Property WrappedShape As Object Implements IShapeNode.WrappedShape
        Get
            Return Me
        End Get
    End Property

    ' =====================================================================
    '  Pembuat node (memudahkan dataset uji terbaca)
    ' =====================================================================

    Public Shared Function Curve(centerX As Double, width As Double, closedCurve As Boolean,
                                 Optional shapeName As String = "", Optional layerName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.CurveShape, shapeName, layerName)
        node.Facts.CurveClosed = closedCurve
        node.Facts.CurveClosedKnown = True
        node.Facts.ShapeName = shapeName
        node.Marker = "curve@" & centerX.ToString()
        Return node
    End Function

    Public Shared Function CurveWithUnknownClosed(Optional shapeName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.CurveShape, shapeName)
        node.Facts.CurveClosedKnown = False
        node.Facts.CurveError = "simulasi: Curve.Closed tidak terbaca"
        Return node
    End Function

    Public Shared Function Text(kind As Integer, Optional shapeName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.TextShape, shapeName)
        node.Facts.TextKind = kind
        Return node
    End Function

    Public Shared Function Bitmap(Optional shapeName As String = "", Optional layerName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.BitmapShape, shapeName, layerName)
        Return node
    End Function

    Public Shared Function Rectangle(Optional shapeName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.RectangleShape, shapeName)
        Return node
    End Function

    Public Shared Function Ellipse(Optional shapeName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.EllipseShape, shapeName)
        Return node
    End Function

    Public Shared Function Polygon(Optional shapeName As String = "") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.PolygonShape, shapeName)
        Return node
    End Function

    Public Shared Function Guideline() As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.GuidelineShape, "Garis bantu")
        Return node
    End Function

    Public Shared Function OleObject() As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.OleObjectShape, "OLE")
        Return node
    End Function

    Public Shared Function UnreadableShape() As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = New ShapeFacts()
        node.Facts.TypeCode = ShapeKind.Unreadable
        node.Facts.TypeError = "simulasi: shape.Type tidak terbaca"
        Return node
    End Function

    Public Shared Function WatermarkShape() As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.TextShape, "WM_TEXT")
        node.Facts.TextKind = ShapeKind.TextArtistic
        Return node
    End Function

    Public Shared Function WatermarkLayerBitmap() As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.BitmapShape, "logo", "WATERMARK")
        Return node
    End Function

    Public Shared Function LockedLayerRectangle() As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.RectangleShape, "Kotak terkunci", "Layer kunci")
        node.Facts.LayerEditable = False
        node.Facts.LayerEditableKnown = True
        Return node
    End Function

    Public Shared Function Group(children As IEnumerable(Of IShapeNode),
                                 Optional shapeName As String = "Grup") As FakeShapeNode
        Dim node As New FakeShapeNode()
        node.Facts = ShapeFacts.FromValues(ShapeKind.GroupShape, shapeName)
        For Each child As IShapeNode In children
            node.ChildNodes.Add(child)
        Next
        Return node
    End Function

    Public Shared Function GroupWithUnreadableChildren() As FakeShapeNode
        Dim node As FakeShapeNode = Group(New List(Of IShapeNode)(), "Grup rusak")
        node.ChildrenErrorMessage = "simulasi: isi grup tidak terbaca"
        Return node
    End Function

    ''' <summary>PowerClip container yang berisi objek (mis. bitmap di dalamnya).</summary>
    Public Shared Function PowerClip(contents As IEnumerable(Of IShapeNode),
                                     Optional shapeName As String = "PowerClip") As FakeShapeNode
        Dim node As FakeShapeNode = Group(New List(Of IShapeNode)(), shapeName)
        node.IsPowerClip = True
        For Each child As IShapeNode In contents
            node.PowerClipNodes.Add(child)
        Next
        Return node
    End Function

    Public Shared Function PowerClipWithUnreadableContents() As FakeShapeNode
        Dim node As FakeShapeNode = Group(New List(Of IShapeNode)(), "PowerClip rusak")
        node.IsPowerClip = True
        node.PowerClipErrorMessage = "simulasi: isi PowerClip tidak terbaca"
        Return node
    End Function
End Class
