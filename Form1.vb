Imports System.Runtime.InteropServices
Imports System.Text

Public Class Form1 
    <DllImport("user32.dll", EntryPoint:="GetWindowThreadProcessId")> _
    Private Shared Function GetWindowThreadProcessId(<InAttribute()> ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="GetForegroundWindow")> Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", EntryPoint:="GetWindowTextW")> _
    Private Shared Function GetWindowTextW(<InAttribute()> ByVal hWnd As IntPtr, <OutAttribute(), MarshalAs(UnmanagedType.LPWStr)> ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function

    Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As IntPtr) As Integer
    Private Declare Function SetFocus Lib "user32.dll" (ByVal hWnd As IntPtr) As Integer

    Public nombreForm As String = ""

    <DllImport("user32.dll")> _
    Shared Function GetAsyncKeyState(ByVal vKey As System.Windows.Forms.Keys) As Short
    End Function

    Private Declare Function GetAsynkeyState Lib "user32" (ByVal vKey As Integer) As Short

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try 


                    If GetAsyncKeyState(Keys.Return) <> 0 Then
                        Console.WriteLine("ENTER")
                    End If

                    Dim hWnd As IntPtr = GetForegroundWindow()
                    If Not hWnd.Equals(IntPtr.Zero) Then
                        Dim lgth As Integer = GetWindowTextLength(hWnd)
                        Dim wTitle As New System.Text.StringBuilder("", lgth + 1)
                        If lgth > 0 Then
                            GetWindowTextW(hWnd, wTitle, wTitle.Capacity)
                        End If

                        Dim wProcID As Integer = Nothing
                        GetWindowThreadProcessId(hWnd, wProcID)

                        Dim Proc As Process = Process.GetProcessById(wProcID)
                        Dim wFileName As String = ""
                        Try
                            wFileName = Proc.MainModule.FileName
                            Console.WriteLine(wFileName.ToString)
                            Console.WriteLine(GetAsynkeyState(65).ToString)

                        Catch ex As Exception
                            wFileName = ""
                        End Try
                        If wTitle.ToString.Length > 0 And Not wTitle.ToString.Equals("TECLAS ESPECIALES") Then
                            nombreForm = wTitle.ToString
                            'formSelect = wProcID
                            Console.WriteLine(nombreForm)
                            'Console.WriteLine(formSelect)
                        End If
                    End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
            nombreForm = ""
        End Try


    End Sub

    Private Function capturakey()
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("<")
                    Console.WriteLine("<")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
        Return 0
    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("<")
                    Console.WriteLine("<")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send(">")
                    Console.WriteLine(">")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("{{}")
                    Console.WriteLine("{")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("{}}")
                    Console.WriteLine("}")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("[")
                    Console.WriteLine("[")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("]")
                    Console.WriteLine("]")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("=")
                    Console.WriteLine("=")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("!")
                    Console.WriteLine("!")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("'")
                    Console.WriteLine("'")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send(";")
                    Console.WriteLine(";")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("#")
                    Console.WriteLine("#")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("$")
                    Console.WriteLine("$")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("{%}")
                    Console.WriteLine("%")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("&")
                    Console.WriteLine("&")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("|")
                    Console.WriteLine("|")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("@")
                    Console.WriteLine("@")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send(":")
                    Console.WriteLine(":")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("+")
                    Console.WriteLine("+")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub


    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("-")
                    Console.WriteLine("-")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("*")
                    Console.WriteLine("*")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("/")
                    Console.WriteLine("/")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("\")
                    Console.WriteLine("\")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("""")
                    Console.WriteLine("""")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("_")
                    Console.WriteLine("_")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("¿")
                    Console.WriteLine("¿")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("?")
                    Console.WriteLine("?")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send(".")
                    Console.WriteLine(".")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("°")
                    Console.WriteLine("°")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    'SendKeys.Send("^") 
                    Dim wsh, MyKey
                    wsh = CreateObject("Wscript.Shell")
                    MyKey = "{^}"
                    wsh.SendKeys(MyKey)

                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("Ñ")
                    Console.WriteLine("Ñ")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send("ñ")
                    Console.WriteLine("ñ")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try 
    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        Try
            If nombreForm.Length > 0 Then
                Dim selecthwnd As IntPtr = FindWindow(Nothing, nombreForm)
                If Not selecthwnd = 0 Then
                    Console.WriteLine(selecthwnd.ToString)
                    SetForegroundWindow(selecthwnd)
                    SetFocus(selecthwnd)
                    SendKeys.Send(",")
                    Console.WriteLine(",")
                    Beep()
                End If
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try
    End Sub
End Class
