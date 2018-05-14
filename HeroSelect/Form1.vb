Imports System.Runtime.InteropServices

Public Class Form1

    Private WithEvents keyPress As New KeyboardHook

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim heroList As New List(Of String)(New String() {
                                            "Ana",
                                            "Bastion",
                                            "Brigitte",
                                            "Doomfist",
                                            "DVA",
                                            "Genji",
                                            "Hanzo",
                                            "Junkrat",
                                            "Lucio",
                                            "McCree",
                                            "Mei",
                                            "Mercy",
                                            "Moira",
                                            "Orisa",
                                            "Pharah",
                                            "Reaper",
                                            "Reinhardt",
                                            "Roadhog",
                                            "Soldier76",
                                            "Sombra",
                                            "Symmetra",
                                            "Torbjorn",
                                            "Tracer",
                                            "Widowmaker",
                                            "Winston",
                                            "Zarya",
                                            "Zenyatta"})

        Dim txtSelectedHero = New Label
        txtSelectedHero.AutoSize = False
        txtSelectedHero.Size = New Drawing.Size(876, 50)
        txtSelectedHero.Location = New Point(0, 620)
        txtSelectedHero.Font = New Drawing.Font("BigNoodleTooOblique", 36)
        txtSelectedHero.ForeColor = Color.White
        txtSelectedHero.TextAlign = ContentAlignment.MiddleCenter
        txtSelectedHero.Name = "txtSelectedHero"
        Me.Controls.Add(txtSelectedHero)

        Dim heroHeader = New Label
        heroHeader.AutoSize = False
        heroHeader.Size = New Drawing.Size(876, 30)
        heroHeader.Location = New Point(0, 585)
        heroHeader.Font = New Drawing.Font("BigNoodleTooOblique", 20)
        heroHeader.ForeColor = Color.White
        heroHeader.TextAlign = ContentAlignment.TopCenter
        heroHeader.Text = "Selected Hero:"
        heroHeader.Name = "SelectedHero"
        Me.Controls.Add(heroHeader)

        Dim buttons As Button = Nothing
        For i As Integer = 0 To 26
            buttons = New Button()
            buttons.Size = New Drawing.Size(90, 155)
            Dim fileLocation As String
            fileLocation = IO.Path.Combine(My.Application.Info.DirectoryPath, "HeroImgs\" & heroList(i) & ".png")
            buttons.Image = Image.FromFile(fileLocation)
            Dim xSpacer = 95
            If i < 9 Then
                buttons.Location = New Point(5 + (i * xSpacer), 110)
            ElseIf i < 18 Then
                buttons.Location = New Point(5 + ((i - 9) * xSpacer), 270)
            Else
                buttons.Location = New Point(5 + ((i - 18) * xSpacer), 430)
            End If
            Dim buttonName = "btn" & heroList(i)
            buttons.Name = buttonName
            buttons.FlatStyle = FlatStyle.Flat
            Me.Controls.Add(buttons)
            'AddHandler buttons.MouseEnter, AddressOf enterButton
            'AddHandler buttons.MouseLeave, AddressOf leaveButton
            AddHandler buttons.MouseClick, AddressOf clickButton

        Next
    End Sub

    Private Sub enterButton(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.FlatAppearance.BorderColor = Color.White

        setHeroText(sender)

    End Sub

    Private Sub leaveButton(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.FlatAppearance.BorderColor = Color.Black
    End Sub

    Private Sub clickButton(ByVal sender As Object, ByVal e As System.EventArgs)
        setHeroText(sender)
    End Sub

    Private Sub setHeroText(sender)
        Dim btn As Button = DirectCast(sender, Button)
        Dim btnName As String = btn.Name
        btnName = btnName.Remove(0, 3)

        Dim selectedHero() As Control
        selectedHero = Me.Controls.Find("txtSelectedHero", True)
        selectedHero(0).Text = btnName
    End Sub

    Public Class GlobalVariables
        Public Shared runValue As Boolean = False
    End Class

    Private Sub f4KeyDown(ByVal Key As Keys) Handles keyPress.KeyDown
        If Key.ToString = "F4" Then
            If GlobalVariables.runValue = True Then
                GlobalVariables.runValue = False
            Else
                GlobalVariables.runValue = True
            End If
            startSelectHero()
        End If

    End Sub

    Function GetPixelColor(ByVal Location As Point)
        Dim b As New Bitmap(1, 1)
        Dim g As Graphics = Graphics.FromImage(b)

        g.CopyFromScreen(Location, Point.Empty, New Size(1, 1))

        'Windows.Forms.Cursor.Position = New Point(Location)

        Return b.GetPixel(0, 0)

    End Function

    Function CompareColors(ByVal col As Color)
        Dim compareColor As New Color()
        compareColor = Color.FromArgb(255, 0, 0, 0)


        If col.Equals(compareColor) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub startSelectHero()
        While GlobalVariables.runValue = True
            Dim offenseIconLocation As New Point(112, 831)
            Dim defenceIconLocation As New Point(641, 836)
            Dim tankIconLocation As New Point(1044, 831)
            Dim supportIconLocation As New Point(1446, 832)

            Console.WriteLine(GetPixelColor(offenseIconLocation))
            Console.WriteLine(GetPixelColor(defenceIconLocation))
            Console.WriteLine(GetPixelColor(tankIconLocation))
            Console.WriteLine(GetPixelColor(supportIconLocation))

            Console.WriteLine(CompareColors(GetPixelColor(offenseIconLocation)))
            Console.WriteLine(CompareColors(GetPixelColor(defenceIconLocation)))
            Console.WriteLine(CompareColors(GetPixelColor(tankIconLocation)))
            Console.WriteLine(CompareColors(GetPixelColor(supportIconLocation)))

            Dim heroCoords As Dictionary(Of String, String) = New Dictionary(Of String, String) From
            {
                {"Doomfist", "117:897"},
                {"Genji", "181: 897"},
                {"McCree", "245:897"},
                {"Pharah", "309:897"},
                {"Reaper", "371:897"},
                {"Soldier76", "432:897"},
                {"Sombra", "499:897"},
                {"Tracer", "558:897"},
                {"Bastion", "645:897"},
                {"Hanzo", "711:897"},
                {"Junkrat", "773:897"},
                {"Mei", "834:897"},
                {"Tobjorn", "896:897"},
                {"Widowmaker", "960:897"},
                {"D.Va", "1053:897"},
                {"Orisa", "1117:897"},
                {"Reinhardt", "1178:897"},
                {"Roadhog", "1240:897"},
                {"Winston", "1301:897"},
                {"Zarya", "1363:897"},
                {"Ana", "1447: 897"},
                {"Brigitte", "1512:897"},
                {"Lucio", "1575:897"},
                {"Mercy", "1639:897"},
                {"Moira", "1701:897"},
                {"Symmetra", "1769:897"},
                {"Zenyatta", "1830:897"}
            }

            If CompareColors(GetPixelColor(offenseIconLocation)) And CompareColors(GetPixelColor(defenceIconLocation)) And CompareColors(GetPixelColor(tankIconLocation)) And CompareColors(GetPixelColor(supportIconLocation)) Then
                Dim selectedHero() As Control
                Dim selectedValue As String
                selectedHero = Me.Controls.Find("txtSelectedHero", True)
                selectedValue = selectedHero(0).Text

                If heroCoords.ContainsKey(selectedValue) Then
                    Dim coordStr As String = heroCoords.Item(selectedValue)
                    Dim coords As String() = coordStr.Split(New Char() {":"})

                    Dim selectedLocation As New Point(coords(0), coords(1))

                    Windows.Forms.Cursor.Position = New Point(selectedLocation)

                    mouse_event(&H2, 0, 0, 0, 0)
                    mouse_event(&H4, 0, 0, 0, 0)

                    mouse_event(&H2, 0, 0, 0, 0)
                    mouse_event(&H4, 0, 0, 0, 0)

                    Console.WriteLine("Clicked")

                    GlobalVariables.runValue = False

                End If
            End If

        End While

    End Sub

    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)

    Private Structure heroSelectCoords
        Public heroName As String
        Public xValue As Integer
        Public yValue As Integer
    End Structure

End Class

Public Class KeyboardHook

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As KBDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Overloads Shared Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Private Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    <Flags()>
    Private Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum

    Public Shared Event KeyDown(ByVal Key As Keys)
    Public Shared Event KeyUp(ByVal Key As Keys)

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const HC_ACTION As Integer = 0
    Private Const WM_KEYDOWN = &H100
    Private Const WM_KEYUP = &H101
    Private Const WM_SYSKEYDOWN = &H104
    Private Const WM_SYSKEYUP = &H105

    Private Delegate Function KBDLLHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

    Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardProc)
    Private HHookID As IntPtr = IntPtr.Zero

    Private Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        If (nCode = HC_ACTION) Then
            Dim struct As KBDLLHOOKSTRUCT
            Select Case wParam
                Case WM_KEYDOWN, WM_SYSKEYDOWN
                    RaiseEvent KeyDown(CType(CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT).vkCode, Keys))
                Case WM_KEYUP, WM_SYSKEYUP
                    RaiseEvent KeyUp(CType(CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT).vkCode, Keys))
            End Select
        End If
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function

    Public Sub New()
        HHookID = SetWindowsHookEx(WH_KEYBOARD_LL, KBDLLHookProcDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
        If HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set keyboard hook")
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(HHookID)
        End If
        MyBase.Finalize()
    End Sub

End Class
