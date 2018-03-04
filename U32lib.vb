Imports System.Drawing
Imports System.Windows.Forms

Public Class User32
    Enum MouseEvents

        leftdown = &H2
        leftup = &H4
        rightdown = &H8
        rightup = &H10
        middledown = &H20
        middleup = &H40

    End Enum

    Enum WindowSettings
        hidewindow = &H80
        bringtofrontbutdontactivate = &H10
        retainposition = &H2
        donotchangezorder = &H200
        noresize = &H1
        showwindow = &H40
    End Enum
    Enum FlashWindowSettings
        stopflashing = &H0
        flashcaption = &H1
        flashtray = &H2
        flashall = &H3
        flashtimer = &H4
        flashtillforgound = &H12
    End Enum

    Enum windowdetailsflags
        asyncwindowplacement = &H4
        restoretomaximized = &H2
        setminposition = &H1


    End Enum

    Enum windowdetailsshowCMD
        hide = &H0
        maximize = &H3
        minimize = &H6
        restore = &H9
        show = &H5
        showmaximized = &H3
        showminimized = &H2
        showminimizednotative = &H7
        shownotactive = &H8
        showtopmostnotactive = &H4
        shownormal = &H1

    End Enum

    Public Structure flashwindowstruct
        Shared cbSize As UInteger = Convert.ToUInt32(Runtime.InteropServices.Marshal.SizeOf(GetType(flashwindowstruct)))
        Public hwnd As IntPtr
        Public dwFlags As UInteger
        Public uCount As UInteger
        Public dwTimeout As Integer
    End Structure
    Public Structure cursorinfostruct
        Shared cbSize As UInteger = Convert.ToUInt32(Runtime.InteropServices.Marshal.SizeOf(GetType(cursorinfostruct)))
        Public flags As UInteger
        Public ptScreenPos As Point
        Public hCursor As IntPtr
    End Structure

    Public Structure getwindowplacementstruct
        Shared length As UInteger = Convert.ToUInt32(Runtime.InteropServices.Marshal.SizeOf(GetType(getwindowplacementstruct)))
        Public flags As UInteger
        Public showCmd As UInteger
        Public ptMinPosition As Point
        Public ptMaxPosition As Point
        Public rcNormalPosition As Rectangle


    End Structure

    Public Shared Sub bringwindowtofront(proc As Process)
        If proc.Handle <> 0 Then
            SetWindowPos(proc.Handle, -1, 0, 0, 0, 0, WindowSettings.retainposition Or WindowSettings.noresize Or WindowSettings.showwindow)
        End If
    End Sub

    Public Shared Function getMouseButtonHex(button As MouseButtons) As SortedList
        Dim rxcx As New MouseButtons
        Dim crs As New SortedList

        rxcx = button
        crs.Add("down", System.Enum.Parse(GetType(MouseEvents), rxcx.GetName(rxcx.GetType, button).ToLower & "down"))
        crs.Add("up", System.Enum.Parse(GetType(MouseEvents), rxcx.GetName(rxcx.GetType, button).ToLower & "up"))

        Return crs

    End Function

    <Runtime.InteropServices.DllImport("user32")> Shared Function GetWindowPlacement(windowid As IntPtr, ByRef windowdetails As getwindowplacementstruct) As Boolean

    End Function

    <Runtime.InteropServices.DllImport("user32")> Shared Function GetCursorInfo(ByRef cursordetails As cursorinfostruct) As Boolean

    End Function

    <Runtime.InteropServices.DllImport("user32")> Public Shared Function FlashWindowEx(ByRef windowdetails As flashwindowstruct) As Boolean

    End Function
    <Runtime.InteropServices.DllImport("user32")> Shared Function GetCursorPos(ByRef Location As Point) As Boolean

    End Function
    <Runtime.InteropServices.DllImport("user32")> Shared Function SetCursorPos(x As Integer, y As Integer) As Boolean

    End Function
    <Runtime.InteropServices.DllImport("user32")> Shared Function mouse_event(eventflags As UInteger, xloc As UInteger, yloc As UInteger, extra1 As UInteger, extra2 As UInteger) As Boolean

    End Function
    <Runtime.InteropServices.DllImport("user32")> Shared Function SetWindowPos(windowhandle As IntPtr, windowinsertafter As IntPtr, xpos As Integer, ypos As Integer, width As Integer, height As Integer, windowsettings As UInteger) As Boolean

    End Function


    <Runtime.InteropServices.DllImport("user32")> Shared Function GetWindowRect(windowhandle As IntPtr, ByRef windowbounds As Rectangle) As Boolean

    End Function

    <Runtime.InteropServices.DllImport("user32")> Shared Function IsZoomed(windowhandle As IntPtr) As Boolean

    End Function
End Class