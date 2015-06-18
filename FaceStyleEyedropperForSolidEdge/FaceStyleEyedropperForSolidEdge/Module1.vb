'Written by Jason Titcomb while working for LMGi www.tlmgi.com 6/18/2015
'This code is provided AS IS.
Imports System.Runtime.InteropServices
Imports SolidEdgeConstants, SolidEdgeFramework
Imports SolidEdgeFramework.DocumentTypeConstants

Module Module1
    Private mSolidApp As SolidEdgeFramework.Application = Nothing
    Private mCommand As SolidEdgeFramework.Command = Nothing
    Private mMouse As SolidEdgeFramework.Mouse = Nothing
    Public Sub main()
        Try
            mSolidApp = TryCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            MessageFilter.Register()

            Select Case mSolidApp.ActiveDocumentType
                Case igPartDocument, igSyncPartDocument, igSheetMetalDocument, igSyncSheetMetalDocument
                Case Else
                    ReportStatus("Part or Sheetmetal only!")
                    Exit Sub
            End Select

            mCommand = mSolidApp.CreateCommand(SolidEdgeConstants.seCmdFlag.seNoDeactivate)
            AddHandler mCommand.Terminate, AddressOf command_Terminate

            mCommand.Start()
            mMouse = mCommand.Mouse
            With mMouse
                .LocateMode = 1
                .WindowTypes = 1
                .EnabledMove = True
                .AddToLocateFilter(SolidEdgeConstants.seLocateFilterConstants.seLocateFace)
                AddHandler .MouseDown, AddressOf mouse_MouseDown
                AddHandler .MouseMove, AddressOf mouse_MouseMove
            End With
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            ReportStatus(ex.Message)
        Finally
            MessageFilter.Revoke()
        End Try
    End Sub

    Private Sub ReportStatus(msg As String)
        mSolidApp.StatusBar = msg
    End Sub

    Private Sub mouse_MouseDown(ByVal sButton As Short, ByVal sShift As Short, ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, ByVal pWindowDispatch As Object, ByVal lKeyPointType As Integer, ByVal pGraphicDispatch As Object)
        If sButton = 2 Or pGraphicDispatch Is Nothing Then
            mSolidApp.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.sePartSelectCommand)
        End If
        ReportStyleName(pGraphicDispatch, True)
    End Sub

    Private Sub mouse_MouseMove(ByVal sButton As Short, ByVal sShift As Short, ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, ByVal pWindowDispatch As Object, ByVal lKeyPointType As Integer, ByVal pGraphicDispatch As Object)
        ReportStyleName(pGraphicDispatch, False)
    End Sub

    Private Sub command_Terminate()
        ReleaseRCW(mMouse)
        ReleaseRCW(mCommand)
        ReleaseRCW(mSolidApp)
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub ReportStyleName(myOface As Object, openStyle As Boolean)
        Dim fs As FaceStyle = Nothing
        Dim haveStyle As Boolean = False
        If myOface IsNot Nothing Then
            Dim theFace As SolidEdgeGeometry.Face = DirectCast(myOface, SolidEdgeGeometry.Face)
            If theFace.Style IsNot Nothing Then
                fs = theFace.Style
                mSolidApp.StatusBar = fs.StyleName
                haveStyle = True
            Else
                'feature style
                Dim p As Object = Nothing
                If PropExists(theFace, "Parent", p) Then
                    fs = p.GetStyle()
                    If fs IsNot Nothing Then
                        mSolidApp.StatusBar = fs.StyleName
                        haveStyle = True
                    End If
                End If

                'look for body style
                If theFace.Body IsNot Nothing And Not haveStyle Then
                    Dim bdy As SolidEdgeGeometry.Body = theFace.Body
                    If bdy.Style IsNot Nothing Then
                        fs = bdy.Style
                        mSolidApp.StatusBar = fs.StyleName
                        haveStyle = True
                    End If
                End If
            End If

            If Not haveStyle Then
                mSolidApp.StatusBar = "Material table style"
            End If
        Else
            mSolidApp.StatusBar = "Position mouse over face"
        End If

        If openStyle And haveStyle Then
            mSolidApp.StartCommand(SolidEdgeConstants.PartCommandConstants.PartFormatStyle)
        End If

    End Sub

    Private Function PropExists(o As Object, name As String, ByRef retProp As Object) As Boolean
        retProp = o.GetType.InvokeMember(name, Reflection.BindingFlags.GetProperty, Nothing, o, Nothing)
        Return retProp IsNot Nothing
    End Function


    Private Sub ReleaseRCW(ByRef o As Object)
        If o IsNot Nothing Then
            Dim ret As Integer = Marshal.ReleaseComObject(o)
            Debug.Assert(0 = ret)
            o = Nothing
        End If
    End Sub
End Module
