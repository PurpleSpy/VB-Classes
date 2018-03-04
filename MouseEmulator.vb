
Public Class MouseEmulator

    Public Property DoubleClickDelay As Integer = 200
    Public Property MouseMoveSteps As Integer = 2000
    Public Property MouseMoveFPS As Integer = 30
    Sub MouseDoubleClick(x As Integer, y As Integer, button As MouseButtons)
        MouseDoubleClick(New Point(x, y), button)
    End Sub

    Sub MouseClick(x As Integer, y As Integer, button As MouseButtons)
        MouseClick(New Point(x, y), button)
    End Sub

    Sub MouseUp(x As Integer, y As Integer, button As MouseButtons)
        MouseUp(New Point(x, y), button)
    End Sub

    Sub MouseDown(x As Integer, y As Integer, button As MouseButtons)
        MouseDown(New Point(x, y), button)
    End Sub

    Sub MouseDrag(fromx As Integer, fromy As Integer, tox As Integer, toy As Integer, button As MouseButtons)
        MouseDrag(New Point(fromx, fromy), New Point(tox, toy), button)
    End Sub

    Sub MouseMove(x As Integer, y As Integer)
        MouseMove(New Point(x, y))
    End Sub
    Sub MouseMove(location As Point)
        User32.SetCursorPos(location.X, location.Y)
    End Sub

    Sub MouseMoveTimed(fromlocation As Point, tolocation As Point)
        Dim deltapt As Point = tolocation - fromlocation
        For i = 1 To MouseMoveSteps
            MouseMove(tolocation + New Point(deltapt.X / i, deltapt.Y / i))
            Threading.Thread.Sleep(1000 / MouseMoveFPS)
        Next
    End Sub
    Sub MouseDragtimed(fromlocation As Point, tolocation As Point, button As MouseButtons)
        MouseMove(fromlocation)
        MouseDown(fromlocation, button)

        MouseMoveTimed(fromlocation, tolocation)

        MouseMove(tolocation)
        MouseUp(fromlocation, button)
    End Sub

    Sub MouseDrag(fromlocation As Point, tolocation As Point, button As MouseButtons)
        MouseMove(fromlocation)
        MouseDown(fromlocation, button)
        MouseMove(tolocation)
        MouseUp(fromlocation, button)
    End Sub


    Sub MouseDoubleClick(location As Point, button As MouseButtons)
        MouseClick(location, button)
        Threading.Thread.Sleep(DoubleClickDelay)
        MouseClick(location, button)
    End Sub
    Sub MouseClick(location As Point, button As MouseButtons)
        MouseMove(location)
        MouseDown(location, button)
        Threading.Thread.Sleep(100)
        MouseUp(location, button)


    End Sub
    Sub MouseUp(location As Point, button As MouseButtons)
        Dim butstat As SortedList = User32.getMouseButtonHex(button)
        User32.mouse_event(butstat("up"), Convert.ToUInt32(location.X), Convert.ToUInt32(location.Y), 0, 0)
    End Sub

    Sub MouseDown(location As Point, button As MouseButtons)
        Dim butstat As SortedList = User32.getMouseButtonHex(button)
        User32.mouse_event(butstat("down"), Convert.ToUInt32(location.X), Convert.ToUInt32(location.Y), 0, 0)

    End Sub

    Sub playmacro(macro As MouseMacro)
        Dim flm As IEnumerator = macro.GetEnumerator

        While flm.MoveNext
            Dim fx As MouseMacro.MouseTask = flm.Current

            Select Case fx.MouseEvent
                Case MouseMacro.mouseevents.mouseclick
                    MouseClick(fx.location, fx.Button)
                Case MouseMacro.mouseevents.mousedoubleclick
                    MouseDoubleClick(fx.location, fx.Button)
                Case MouseMacro.mouseevents.mousedown
                    MouseDown(fx.location, fx.Button)
                Case MouseMacro.mouseevents.mousedrag
                    MouseDrag(fx.FromPt, fx.ToPt, fx.Button)
                Case MouseMacro.mouseevents.mousemove
                    MouseMove(fx.location)
                Case MouseMacro.mouseevents.mouseup
                    MouseUp(fx.location, fx.Button)
                Case MouseMacro.mouseevents.pause
                    Threading.Thread.Sleep(fx.delay)
            End Select
        End While
    End Sub
    <Serializable> Public Class MouseMacro
        Inherits Collections.ArrayList

        Enum mouseevents
            pause
            mousedown
            mouseup
            mousemove
            mousedrag
            mousedragtimed
            mouseclick
            mousedoubleclick
        End Enum

        Public Sub addTask(mouseMoveTask As MouseTask)
            Me.Add(mouseMoveTask)
        End Sub
        Public Sub addTask(mouseMoveTask As MouseTask, insertindex As Integer)
            Me.Insert(insertindex, mouseMoveTask)
        End Sub

        Public Sub addTask(mouseev As mouseevents, mousebutton As MouseButtons, frompoint As Point, topoint As Point, locationpt As Point, waitdelay As Integer)
            Me.Add(New MouseTask(mouseev, mousebutton, frompoint, topoint, locationpt, waitdelay))
        End Sub
        Public Sub addTask(mouseev As mouseevents, mousebutton As MouseButtons, frompoint As Point, topoint As Point, locationpt As Point, waitdelay As Integer, insertat As Integer)
            addTask(New MouseTask(mouseev, mousebutton, frompoint, topoint, locationpt, waitdelay), insertat)
        End Sub
        Public Sub removeTask(mouseMoveTask As MouseTask)
            Me.Remove(mouseMoveTask)
        End Sub
        Public Sub removeTask(mouseMoveTask As Integer)
            Me.RemoveAt(mouseMoveTask)
        End Sub



        Public Sub saveMacro(location As String)
            Using scrx As New IO.FileStream(location, IO.FileMode.Create)
                Dim bsx As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
                Dim csx As New Xml.Serialization.XmlSerializer(Me.GetType)

                bsx.Serialize(scrx, Me)
                scrx.Flush()
                scrx.Close()

            End Using
        End Sub

        <Serializable> Public Class MouseTask

            Public Property MouseEvent As mouseevents
            Public Property Button As MouseButtons
            Public Property FromPt As Point
            Public Property ToPt As Point
            Public Property location As Point
            Public Property delay As Integer

            Public Sub New(mouseev As mouseevents, mousebutton As MouseButtons, frompoint As Point, topoint As Point, locationpt As Point, waitdelay As Integer)
                MouseEvent = mouseev
                Button = mousebutton
                FromPt = frompoint
                ToPt = topoint
                location = locationpt
                delay = waitdelay
            End Sub

            Public Sub New()

            End Sub


        End Class
    End Class
End Class




