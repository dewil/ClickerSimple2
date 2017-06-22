Imports System.Runtime.InteropServices
Imports BrowseEmAll.Gecko
Imports BrowseEmAll.Gecko.Core
Imports BrowseEmAll.Gecko.Core.Events
Imports BrowseEmAll.Gecko.Winforms
Imports System.Threading
Imports System.Net
Imports System.Web
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Xml

Public Class Form1
    Const MOUSEEVENTF_LEFTDOWN As Integer = 2
    Const MOUSEEVENTF_LEFTUP As Integer = 4
    Const MOUSEEVENTF_RIGHTDOWN As Integer = 8
    Const MOUSEEVENTF_RIGHT_UP As Integer = 16
    'input type constant
    Const INPUT_MOUSE As Integer = 0
    Private clickLocation As Point
    Dim firefoxpath As String = ""
    Dim brsw As GeckoWebBrowser
    Private m_tabControl As TabControl
    Private counter As Long = 0
    Private th As Thread
    Dim zagl As Boolean = False
    Dim pageloaded As Boolean = False

    Public settings_file As String = "settings.xml"
    Public GLOBAL_URL As String = ""
    Public GLOBAL_TOKEN As String = ""
    Public GLOBAL_TIMEOUT As String = (1000 * 60 * 3).ToString
    Public GLOBAL_ENTER_TIMEOUT As String = (1000 * 5).ToString

    Public working As Boolean = False

    'for mouse
    <DllImport("User32.dll", SetLastError:=True)>
    Public Shared Function SendInput(nInputs As Integer, ByRef pInputs As INPUT, cbSize As Integer) As Integer
    End Function

    ' Get a handle to an application window.
    <DllImport("USER32.DLL", CharSet:=CharSet.Unicode)>
    Public Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function

    ' Activate an application window.
    <DllImport("USER32.DLL")>
    Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    'text cursor coordinate
    <DllImport("user32")>
    Private Shared Function GetCaretPos(ByRef p As Point) As Integer
    End Function
    <DllImport("user32")>
    Private Shared Function SetCaretPos(x As Integer, y As Integer) As Integer
    End Function

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As IntPtr, ByVal wParam As IntPtr, ByVal lParam As String) As Long
    Private Const WM_SETTEXT = &HC


    Private Sub deactivate_mouse()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        Button1.Enabled = False
        Button2.Enabled = False
    End Sub

    Private Sub activate_mouse()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        Button1.Enabled = True
        Button2.Enabled = True
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'If isexpered() Then
        '    MsgBox("Истек срок использования")
        '    Application.Exit()
        'End If

        CheckBox1.Checked = False

        deactivate_mouse()

        firefoxpath = Application.StartupPath + "\firefox"

        Xpcom.Initialize(firefoxpath)

        ' Dim xulrunnerPath As String = XULRunnerLocator.GetXULRunnerLocation()

        'Xpcom.Initialize(xulrunnerPath)

        GeckoPreferences.User("browser.xul.error_pages.enabled") = True

        GeckoPreferences.User("gfx.font_rendering.graphite.enabled") = True

        GeckoPreferences.User("full-screen-api.enabled") = True

        m_tabControl = New TabControl()
        m_tabControl.Dock = DockStyle.Fill

        brsw = AddBrowser()

        Panel2.Controls.Add(m_tabControl)



        AddHandler brsw.Navigated, AddressOf test

        Dim thr As Thread = New Thread(AddressOf get_Set)
        thr.Start()


        ' Timer1.Enabled = True
        Timer1.Interval = GLOBAL_TIMEOUT
        SetLabel1("Таймер отключен.")
        working = False

    End Sub

    Private Function isexpered() As Boolean
        Dim ret As Boolean = False
        If Date.Now >= #6/23/2017# Then 'до 23 июня 2017
            ret = True
        Else
            ret = False
        End If
        Return ret
    End Function

    Protected Function AddBrowser() As GeckoWebBrowser
        Dim tabPage = New TabPage()
        'tabPage.Text = "blank"

        Dim browser = New GeckoWebBrowser()

        ' browser.Navigate("http://www.google.com")
        browser.Dock = DockStyle.Fill
        'tabPage.DockPadding.Top = 25
        tabPage.Dock = DockStyle.Fill

        ' AddToolbarAndBrowserToTab(tabPage, browser)

        m_tabControl.TabPages.Add(tabPage)
        tabPage.Show()
        m_tabControl.SelectedTab = tabPage




        Return browser
    End Function

    Private Sub test()

        th = New Thread(AddressOf mouse_clicker)
        th.Start()

    End Sub


    Private Sub mouse_clicker()

        'Application.DoEvents()

        ' Thread.Sleep(5000)
        Task.Delay(Integer.Parse(GLOBAL_ENTER_TIMEOUT)).Wait()

        If CheckBox1.Checked Then
            '-------------------------------for mouse------------------
            Dim x As Integer = 0
            Dim y As Integer = 0

            Try
                x = CInt(TextBox1.Text.ToString)
            Catch ex As Exception
                MsgBox("Настройте координаты клика.")
                GoTo exit_
            End Try
            Try
                y = CInt(TextBox2.Text.ToString)
            Catch ex As Exception
                MsgBox("Настройте координаты клика.")
                GoTo exit_
            End Try

            clickLocation = New Point(x, y)
            Cursor.Position = clickLocation
            Dim i As New INPUT()
            i.type = INPUT_MOUSE
            i.mi.dx = clickLocation.X '0
            'clickLocation.X;
            i.mi.dy = clickLocation.Y '0
            ' clickLocation.Y;
            i.mi.dwFlags = MOUSEEVENTF_LEFTDOWN
            i.mi.dwExtraInfo = IntPtr.Zero
            i.mi.mouseData = 0
            i.mi.time = 0
            'send the input
            SendInput(1, i, Marshal.SizeOf(i))
            'set the INPUT for mouse up and send
            i.mi.dwFlags = MOUSEEVENTF_LEFTUP
            SendInput(1, i, Marshal.SizeOf(i))
            '-----------------------------------------------------------
        Else

            '------------------клавиша Enter--------------------------------------------------
            ' Dim calculatorHandle As IntPtr = FindWindow("CalcFrame", "Калькулятор")
            Dim calculatorHandle As IntPtr = FindWindow(Nothing, "WhatsApp")
            If calculatorHandle = IntPtr.Zero Then
                MessageBox.Show("WhatsApp is not running.")
                GoTo exit_
            Else
                'MsgBox(calculatorHandle.ToString)
                SetForegroundWindow(calculatorHandle)
                SendKeys.SendWait("{Enter}")
            End If
            '--------------------------------------------------------------------------------
        End If

        '  SetLabel1(counter.ToString)

        'SetLabel1(brsw.Url().ToString)
        'Application.DoEvents()
exit_:
        ' th.Abort()

    End Sub

    Public Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Public Structure INPUT
        Public type As UInteger
        Public mi As MOUSEINPUT
    End Structure

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button2.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim mousePos As Point = System.Windows.Forms.Control.MousePosition
        TextBox1.Text = mousePos.X.ToString
        TextBox2.Text = mousePos.Y.ToString

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        If Button3.Text = "Старт" Then
            Button3.Text = "Стоп"
            Timer1.Enabled = True
            Timer1.Start()
            SetLabel1("Таймер включен (" + (GLOBAL_TIMEOUT / 1000 / 60).ToString + " мин.)")
            If Not working Then
                start_()
            End If
        Else
            Timer1.Stop()
            Timer1.Enabled = False
            Button3.Text = "Старт"
            SetLabel1("Таймер отключен.")
            working = False
        End If
        'start_()


        'brsw.Navigate("https://api.whatsapp.com/send?phone=79111014020&text=%D0%9A%20%D1%81%D0%BE%D0%B6%D0%B0%D0%BB%D0%B5%D0%BD%D0%B8%D1%8E%20%D1%87%D0%B5%D0%BA%20%D1%83%D0%B6%D0%B5%20%D0%B1%D1%8B%D0%BB%20%D1%80%D0%B0%D0%BD%D0%B5%D0%B5%20%D0%B7%D0%B0%D0%B3%D1%80%D1%83%D0%B6%D0%B5%D0%BD%2C%20%D0%BF%D0%BE%D0%BF%D1%80%D0%BE%D0%B1%D1%83%D0%B9%D1%82%D0%B5%20%D0%B7%D0%B0%D0%B3%D1%80%D1%83%D0%B7%D0%B8%D1%82%D1%8C%20%D0%B5%D1%89%D0%B5%20%D0%BE%D0%B4%D0%B8%D0%BD.")



        'clip_test()

        'brsw.Navigate("https://api.whatsapp.com/send?phone=79113863298&text=Я%20заитересован%20в%20покупке%20вашего%20авто")
        'browser.Show()
        ' MsgBox(firefoxpath)
        'browser.Navigate("http://google.com")
        ' WebBrowser1.Navigate("https://api.whatsapp.com/send?phone=79113863298&text=Я%20заитересован%20в%20покупке%20вашего%20авто")
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            activate_mouse()
        Else
            deactivate_mouse()
        End If
    End Sub
    Private Sub clip_test()
        Dim x As Integer = 0
        Dim y As Integer = 0

        Try
            x = CInt(TextBox1.Text.ToString)
        Catch ex As Exception
            MsgBox("Настройте координаты клика.")
            GoTo exit_
        End Try
        Try
            y = CInt(TextBox2.Text.ToString)
        Catch ex As Exception
            MsgBox("Настройте координаты клика.")
            GoTo exit_
        End Try

        Clipboard.SetText("+79113863298", TextDataFormat.UnicodeText)

        Dim calculatorHandle As IntPtr = FindWindow(Nothing, "Viber +79113863298")
        'Dim calculatorHandle As IntPtr = FindWindow(Nothing, "WhatsApp")
        If calculatorHandle = IntPtr.Zero Then
            MessageBox.Show("WhatsApp is not running.")
            GoTo exit_
        Else
            'MsgBox(calculatorHandle.ToString)
            SetForegroundWindow(calculatorHandle)

        End If

        clickLocation = New Point(x, y)
        Cursor.Position = clickLocation
        Dim i As New INPUT()
        i.type = INPUT_MOUSE
        i.mi.dx = clickLocation.X '0
        'clickLocation.X;
        i.mi.dy = clickLocation.Y '0
        ' clickLocation.Y;
        i.mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        i.mi.dwExtraInfo = IntPtr.Zero
        i.mi.mouseData = 0
        i.mi.time = 0
        'send the input
        SendInput(1, i, Marshal.SizeOf(i))
        'set the INPUT for mouse up and send
        i.mi.dwFlags = MOUSEEVENTF_LEFTUP
        SendInput(1, i, Marshal.SizeOf(i))

        'Thread.Sleep(1000)

        SendMessage(calculatorHandle, WM_SETTEXT, 0, Clipboard.GetText)


        'Thread.Sleep(1000)
        ' Me.Hide()
        'SendKeys.Send("+79113863298")
        'SendKeys.SendWait("^v")
        ' Dim calculatorHandle As IntPtr = FindWindow("CalcFrame", "Калькулятор")

exit_:
    End Sub

    Private Sub start_()
        Dim i As Task(Of Integer) = get_query()
    End Sub

    Private Async Function get_query() As Task(Of Integer)

        working = True

        SetLabel1("Обработка начата в " + Format(Date.Now, "dd.MM.yyyy HH:mm:ss"))

        Dim client_token As String = GLOBAL_TOKEN
        Dim url_site As String = GLOBAL_URL

        Dim request As WebRequest = WebRequest.Create(url_site + "messages")

        request.Headers.Add("X-Client-Token", client_token)
        request.Headers.Add("X-Accept-Version", "v5")

        Dim response As WebResponse

        Try
            response = request.GetResponse()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            GoTo exit_
        End Try


        Dim dataStream As Stream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim reader As New StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = reader.ReadToEnd()

        reader.Close()
        response.Close()


        Dim ids As List(Of String) = New List(Of String)
        Dim result = JsonConvert.DeserializeObject(Of ArrayList)(responseFromServer)
        If result.Count = 0 Then GoTo exit_
        Dim token As JToken

        Dim msg As String = "В очереди : " + result.Count.ToString + " сообщений." + vbCrLf

        SetLabel1(msg)
        Application.DoEvents()
        Dim cnt As Integer = 0
        Dim first As Integer = -1
begin_:

        Dim res As Integer = 0
        For res = 0 To result.Count - 1



            token = JObject.Parse(result.Item(res).ToString)
            Dim receipt_id As String = token.SelectToken("receipt_id")
            Dim channel As String = token.SelectToken("channel")
            Dim mobile As String = token.SelectToken("mobile")
            Dim message As String = token.SelectToken("message")

            If channel.ToUpper = "WHATSAPP" Then

                If Not mobile = "0" Then
                    cnt += 1

                    If Not mobile.Length = 11 Then
                        'по умолчанию прибавляем код страны Russia 7
                        mobile = "7" + mobile
                    End If

                    SetLabel1("Обработка whatsapp № " + cnt.ToString)
                    Application.DoEvents()
                    brsw.Navigate("https://api.whatsapp.com/send?phone=" + mobile.ToString + "&text=" + Uri.EscapeDataString(message)) ' message.Replace(" ", "%20") 'Uri.EscapeDataString(message))
                    ' Application.DoEvents()
                    ' Task.Delay(10000).Wait()
                    Await Task.Delay(10000)


                End If

                ids.Add(receipt_id)

            End If


        Next



        Dim z As Integer = Await delete_messages(ids, client_token, url_site)




exit_:
        working = False
        Return 0
    End Function

    Private Async Function delete_messages(ids As List(Of String), client_token As String, url_site As String) As Task(Of Integer)
        working = True
        Dim i As Integer = 0
        For i = 0 To ids.Count - 1
            Dim param = ids.Item(i).ToString
            Dim rqst As WebRequest = WebRequest.Create(url_site + "message/check?receipt_id=" + param)
            rqst.Headers.Add("X-Client-Token", client_token)
            rqst.Headers.Add("X-Accept-Version", "v5")
            Dim rsps As WebResponse
            Try
                rsps = rqst.GetResponse()
            Catch ex As Exception
                SetLabel1(param + " - ошибка запроса на удаление: " + ex.Message.ToString)
                GoTo nxt_
            End Try
            Dim dataStream As Stream = rsps.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            ' Read the content.
            Dim responseFromServer As String = reader.ReadToEnd()

            reader.Close()
            rsps.Close()

            SetLabel1(i.ToString + " Запрос на удаление id = " + param.ToString + " ... " + responseFromServer)
nxt_:
            Await Task.Delay(1000)
        Next

        SetLabel1("Обработка завершена в " + Format(Date.Now, "dd.MM.yyyy HH:mm:ss") + " следующая обработка в " + Format(DateAdd(DateInterval.Minute, GLOBAL_TIMEOUT / 60 / 1000, Date.Now), "dd.MM.yyyy HH:mm:ss"))
        Application.DoEvents()

        working = False
        Return 0
    End Function


    Private Delegate Sub SetLabel1Delegate(ByVal str As String)
    Private Sub SetLabel1Callback(ByVal str As String)
        Dim txt As String = str 'Format(Date.Now, "dd.MM.yyyy HH:mm:ss") + " " + str
        Label4.Text = txt
    End Sub
    Private Sub SetLabel1(ByVal str As String)
        Dim txt As String = str
        If Label4.InvokeRequired Then
            Dim Label1Delegate As New SetLabel1Delegate(AddressOf SetLabel1Callback)
            Label4.Invoke(Label1Delegate, New [Object]() {txt})
        Else
            Label4.Text = txt

        End If
    End Sub

    Private Sub get_Set()
        Dim x As Boolean = get_settings()
    End Sub

    Private Function get_settings() As Boolean
        Dim ret As Boolean = True
        Dim timeout As String = "3"
        Dim timeout_enter As String = "5"
        If Dir(settings_file) = "" Then
            MsgBox("Отсутствует файл " + settings_file)
            Return False
        End If

        Try

            Dim xmlDoc As New XmlDocument

            xmlDoc.Load(settings_file)

            SetTextBox(xmlDoc.SelectSingleNode("//URL").InnerText, TextBox3)
            GLOBAL_URL = xmlDoc.SelectSingleNode("//URL").InnerText
            SetTextBox(xmlDoc.SelectSingleNode("//Token").InnerText, TextBox4)
            GLOBAL_TOKEN = xmlDoc.SelectSingleNode("//Token").InnerText
            SetTextBox(xmlDoc.SelectSingleNode("//Timeout").InnerText, TextBox5)
            SetTextBox(xmlDoc.SelectSingleNode("//Waitenter").InnerText, TextBox6)

        Catch ex As Exception
            MsgBox("Чтение файла " + settings_file + " : " + ex.Message.ToString)
            Return False
        End Try

        If TextBox5.Text = "" Or TextBox5.Text = "0" Then
            SetTextBox("3", TextBox5)
        End If

        timeout = Integer.Parse(TextBox5.Text.ToString)
        GLOBAL_TIMEOUT = (timeout * 1000 * 60).ToString

        timeout_enter = Integer.Parse(TextBox6.Text.ToString)
        GLOBAL_ENTER_TIMEOUT = (timeout_enter * 1000).ToString

        Return ret
    End Function

    Private Sub save_settings()

        If TextBox3.Text = "" Then
            MsgBox("Нет URL")
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MsgBox("Нет токена")
            Exit Sub
        End If

        Dim timeout As Integer = 0
        Try
            timeout = Integer.Parse(TextBox5.Text.ToString)
        Catch ex As Exception
            SetTextBox("3", TextBox5)
            MsgBox("Установлен таймаут по умолчанию")
            timeout = 3
        End Try

        Dim timeout_enter As Integer = 0
        Try
            timeout_enter = Integer.Parse(TextBox6.Text.ToString)
        Catch ex As Exception
            SetTextBox("5", TextBox6)
            MsgBox("Установлена задержка перед отправкой WhatsApp по умолчанию")
            timeout_enter = 5
        End Try

        Dim Writer As XmlTextWriter
        Dim File As System.IO.StreamWriter

        If Not Dir(settings_file) = "" Then
            Kill(settings_file)
        End If



        Try
            File = New StreamWriter(settings_file, False, System.Text.Encoding.Unicode)
            Writer = New XmlTextWriter(File)
            Writer.Formatting = Xml.Formatting.Indented

            Writer.WriteStartDocument()
            Writer.WriteStartElement("root")

            Writer.WriteElementString("URL", TextBox3.Text)
            GLOBAL_URL = TextBox3.Text
            Writer.WriteElementString("Token", TextBox4.Text)
            GLOBAL_TOKEN = TextBox4.Text
            Writer.WriteElementString("Timeout", TextBox5.Text)
            GLOBAL_TIMEOUT = (timeout * 1000 * 60).ToString
            Writer.WriteElementString("Waitenter", TextBox6.Text)
            GLOBAL_ENTER_TIMEOUT = (timeout_enter * 1000).ToString

            Writer.WriteEndElement()
            Writer.WriteEndDocument()

            Writer.Close()

        Catch ex As Exception
            MsgBox("Запись " + settings_file + " : " + ex.Message.ToString)
            Exit Sub
        End Try

        MsgBox("Настройки сохранены.")

    End Sub


    Private Delegate Sub SetTextBoxDelegate(ByVal str As String, tb As TextBox)
    Private Sub SetTextBoxCallback(ByVal str As String, tb As TextBox)
        Dim txt As String = str 'Format(Date.Now, "dd.MM.yyyy HH:mm:ss") + " " + str
        tb.Text = txt
    End Sub
    Private Sub SetTextBox(ByVal str As String, tb As TextBox)
        Dim txt As String = str
        If tb.InvokeRequired Then
            Dim SetTextBoxDelegate As New SetTextBoxDelegate(AddressOf SetTextBoxCallback)
            tb.Invoke(SetTextBoxDelegate, New [Object]() {txt, tb})
        Else
            tb.Text = txt

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Not working Then
            start_()
        End If
    End Sub

    Private Class thread_param
        Private _str As String
        Public Property str As String
            Get
                Return _str
            End Get
            Set(value As String)
                _str = value
            End Set
        End Property

    End Class

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        save_settings()

    End Sub
End Class
