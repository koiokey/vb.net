Imports System.ComponentModel
Imports System.Diagnostics.Eventing.Reader
Imports System.DirectoryServices.ActiveDirectory
Imports System.Drawing.Text
Imports System.Media
Imports System.Threading
Imports Microsoft.VisualBasic.Devices

Public Class Small_Game
    Private mouseX As Integer = 0
    Private KCode As Integer = 0
    Private Stay As Integer = 0
    Private Stay_A As Integer = 0

    Dim Scoresound As New SoundPlayer
    Dim Gameoversound As New SoundPlayer
    Dim BackgroundSound As New SoundPlayer

    Dim CoinSound As New SOUND
    Dim BombSound As New SOUND
    '避免共用相同ID
    Private intSoundC = 0
    Private intSoundB = 9999
    Sub BackgroundMusic()
        BackgroundSound.PlayLooping()
    End Sub
    '計數器
    Function Time_count() As Integer
        Static time = 33
        time -= 1
        Return time
    End Function
    Function Check_count() As Integer
        Static check = 2
        check -= 1
        Return check
    End Function
    Function Heart_count() As Integer
        Static heart = 3
        heart -= 1
        Return heart
    End Function
    '^^^計數器

    Function Drop_coin(ByRef pbox As PictureBox, speed As Integer) As Integer
        If pbox.Top <= -pbox.Width Then
            pbox.Location = New Point(CInt(Int(900 * Rnd())) + 100)

        ElseIf pbox.Top >= 800 Then
            pbox.Location = New Point(500, -pbox.Width)
            Return 0
        End If
        '碰撞箱
        If ((PictureBox1.Top - pbox.Top <= pbox.Width) AndAlso PictureBox1.Top - pbox.Top >= -PictureBox1.Height) AndAlso
            (PictureBox1.Left - pbox.Left <= (pbox.Width - 50) AndAlso pbox.Left - PictureBox1.Left <= pbox.Width + 50) Then
            pbox.Location = New Point(500, -pbox.Width)
            Label2.Text = "Score:" & CInt(Label3.Text) + 1000
            Label3.Text = (CInt(Label3.Text) + 1000).ToString
            '音效
            With CoinSound
                CoinSound.Name = "SOUND" & intSoundC
                CoinSound.Play(1, False)
            End With
            intSoundC += 1
            Return 0
        End If

        pbox.Top += speed
        Return 0
    End Function
    Async Function Drop_bomb(pbox As PictureBox, speed As Integer) As Task(Of Integer)
        If pbox.Top <= -pbox.Width Then
            pbox.Location = New Point(CInt(Int(900 * Rnd())) + 100)

        ElseIf pbox.Top >= 800 Then
            pbox.Location = New Point(500, -pbox.Width)
            Return 0
        End If
        '碰撞箱
        If (PictureBox1.Top - pbox.Top <= pbox.Width AndAlso PictureBox1.Top - pbox.Top >= -PictureBox1.Height) AndAlso
            (PictureBox1.Left - pbox.Left <= (pbox.Width - 50) AndAlso pbox.Left - PictureBox1.Left <= pbox.Width + 50) Then
            pbox.Location = New Point(500, -pbox.Width)
            '音效
            With BombSound
                .Name = "SOUND" & intSoundB
                .Play(2, False)
            End With
            intSoundB += 1
            '計算血量
            Dim heart = Heart_count()
            Label4.Text = "x" & (heart).ToString
            'heart -= 1
            If heart = 0 Then
                '所有timer停止
                Timer1.Stop()
                Timer2.Stop()
                Timer3.Stop()
                Timer4.Stop()
                Timer5.Stop()
                Timer6.Stop()
                Timer7.Stop()
                Timer8.Stop()
                Timer9.Stop()
                Timer10.Stop()
                Timer11.Stop()
                '顯示Gameover
                PictureBox13.Visible = True
                PictureBox13.BackgroundImage = My.Resources.ResourceManager.GetObject("gameover")
                PictureBox13.BringToFront()
                My.Computer.Audio.Stop()
                Gameoversound.Play()
                Await Task.Delay(3000)
                'Thread.Sleep(3000)
                'Global SMMoney設定爲0
                ShareMoney.SMMoney = 0
                Me.Close()
            End If
            Return 0
        End If

        pbox.Top += speed
        Return 0
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '隨機數
        Randomize()

        PictureBox1.Left = 500 - PictureBox1.Width / 2 '人物位置
        '無用物件隱藏
        Label5.Visible = False

        '設定背景顔色
        For i = 1 To 11
            Me.Controls.Item("PictureBox" & i).BackColor = Nothing
        Next
        PictureBox2.BackgroundImage = Nothing
        PictureBox1.Hide()
        '移動金幣炸彈到form外
        For i = 3 To 11
            Me.Controls.Item("PictureBox" & i).Top = -PictureBox3.Width
        Next
        '隱藏Label
        Label1.Text = Nothing
        Label1.TextAlign = ContentAlignment.MiddleRight
        Label2.Visible = False
        Label4.Visible = False
        '音效初始化
        Scoresound.Stream = My.Resources.ResourceManager.GetStream("win")
        BackgroundSound.Stream = My.Resources.ResourceManager.GetStream("backgroundmusic")
        Gameoversound.Stream = My.Resources.ResourceManager.GetStream("gameoversound")
        Timer1.Start()

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'A D 左 右 鍵控制方向
        '阻止玩家在倒數前操作產生bug
        If Timer1.Enabled = True Then
            Exit Sub
        End If
        KCode = e.KeyCode
        Stay = 1
        KeyTimer.Start()
    End Sub
    Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp
        '用滑鼠控制人物
        '阻止玩家在倒數前操作產生bug
        If Timer1.Enabled = True Then
            Exit Sub
        End If
        mouseX = e.X
        Stay_A = 1
        MoveTimer.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Static i = 3
        If i = 3 Then
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetStream("first_sound"), AudioPlayMode.Background)
            PictureBox2.BackgroundImage = My.Resources.ResourceManager.GetObject("start3")
        ElseIf i = 2 Then
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetStream("first_sound"), AudioPlayMode.Background)
            PictureBox2.BackgroundImage = My.Resources.ResourceManager.GetObject("start2")
        ElseIf i = 1 Then
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetStream("first_sound"), AudioPlayMode.Background)
            PictureBox2.BackgroundImage = My.Resources.ResourceManager.GetObject("start1")
        ElseIf i = 0 Then
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetStream("pianoA"), AudioPlayMode.Background)
            PictureBox2.BackgroundImage = My.Resources.ResourceManager.GetObject("start")
        End If

        If i = -1 Then '倒數結束后的初始值
            '圖片載入
            PictureBox1.BackgroundImage = My.Resources.ResourceManager.GetObject("runner")
            For j = 3 To 7
                Me.Controls.Item("PictureBox" & j).BackgroundImage = My.Resources.ResourceManager.GetObject("coin")
            Next
            For j = 8 To 11
                Me.Controls.Item("PictureBox" & j).BackgroundImage = My.Resources.ResourceManager.GetObject("bomb")
            Next
            PictureBox12.BackgroundImage = My.Resources.ResourceManager.GetObject("heart")

            PictureBox2.Hide()
            PictureBox1.Show()
            Timer2.Start()
            Timer3.Start()
            Timer4.Start()
            Timer5.Start()
            Timer6.Start()
            Timer7.Start()
            Timer8.Start()
            Timer9.Start()
            Timer10.Start()
            Timer11.Start()
            '開始計算分數
            Label2.Visible = True
            Label2.Text = "Score:" & 0.ToString
            '顯示生命值
            PictureBox12.Visible = True
            Label4.Visible = True
            '結束倒數
            Timer1.Stop()
        End If
        i -= 1
    End Sub

    Private Async Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        '背景音樂
        Dim check = Check_count()
        If check > 0 Then
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("life-of-a-wandering-wizard-15549"),
                                AudioPlayMode.BackgroundLoop)
        End If
        Dim i = Time_count() '有兩秒延遲
        If i = 12 Then
            Label1.ForeColor = Color.Red
        End If
        Label1.Text = i - 2
        i -= 1
        If i = 0 Then
            Thread.Sleep(1000)
            Timer1.Stop()
            Timer2.Stop()
            Timer3.Stop()
            Timer4.Stop()
            Timer5.Stop()
            Timer6.Stop()
            Timer7.Stop()
            Timer8.Stop()
            Timer9.Stop()
            Timer10.Stop()
            Timer11.Stop()
            '顯示Score
            PictureBox14.BackgroundImage = My.Resources.ResourceManager.GetObject("score")
            PictureBox14.BringToFront()
            PictureBox14.Visible = True
            Label6.Visible = True
            '隱藏圖片
            PictureBox1.Hide()
            PictureBox2.Hide()
            PictureBox3.Hide()
            PictureBox4.Hide()
            PictureBox5.Hide()
            PictureBox6.Hide()
            PictureBox7.Hide()
            PictureBox8.Hide()
            PictureBox9.Hide()
            PictureBox10.Hide()
            PictureBox11.Hide()
            PictureBox12.Hide()
            Label1.Hide()
            Label2.Hide()
            Label3.Hide()
            Label4.Hide()
            Label6.Text = Label3.Text
            My.Computer.Audio.Stop()
            Scoresound.Play()
            Await Task.Delay(6000)
            'Global SMMoney 設定爲 Label3
            ShareMoney.SMMoney = Label3.Text
            Me.Close()
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Dim pbox As PictureBox = PictureBox3
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 3)
            Timer3.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox3.Visible = False
            Drop_coin(PictureBox3, speed)
            Exit Sub
        End If
        PictureBox3.Visible = True
        Timer3.Interval = 10
        Drop_coin(PictureBox3, speed)
    End Sub
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        Timer4.Interval = 10
        Dim pbox As PictureBox = PictureBox4
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 3)
            Timer4.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox4.Visible = False
            Drop_coin(PictureBox4, speed)
            Exit Sub
        End If
        PictureBox4.Visible = True
        Timer4.Interval = 10
        Drop_coin(PictureBox4, speed)
    End Sub
    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        Timer5.Interval = 10
        Dim pbox As PictureBox = PictureBox5
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 3)
            Timer5.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox5.Visible = False
            Drop_coin(PictureBox5, speed)
            Exit Sub
        End If
        PictureBox5.Visible = True
        Timer5.Interval = 10
        Drop_coin(PictureBox5, speed)
    End Sub

    Private Sub Timer6_Tick(sender As Object, e As EventArgs) Handles Timer6.Tick
        Timer6.Interval = 10
        Dim pbox As PictureBox = PictureBox6
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 3)
            Timer6.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox6.Visible = False
            Drop_coin(PictureBox6, speed)
            Exit Sub
        End If
        PictureBox6.Visible = True
        Timer6.Interval = 10
        Drop_coin(PictureBox6, speed)
    End Sub

    Private Sub Timer8_Tick(sender As Object, e As EventArgs) Handles Timer8.Tick
        Timer8.Interval = 10
        Dim pbox As PictureBox = PictureBox8
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 5)
            Timer8.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox8.Visible = False
            Drop_bomb(PictureBox8, speed)
            Exit Sub
        End If
        PictureBox8.Visible = True
        Timer8.Interval = 10
        Drop_bomb(PictureBox8, speed)
    End Sub
    Private Sub Timer9_Tick(sender As Object, e As EventArgs) Handles Timer9.Tick
        Timer9.Interval = 10
        Dim pbox As PictureBox = PictureBox9
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 5)
            Timer9.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox9.Visible = False
            Drop_bomb(PictureBox9, speed)
            Exit Sub
        End If
        PictureBox9.Visible = True
        Timer9.Interval = 10
        Drop_bomb(PictureBox9, speed)
    End Sub
    Private Sub Timer10_Tick(sender As Object, e As EventArgs) Handles Timer10.Tick
        Timer10.Interval = 10
        Dim pbox As PictureBox = PictureBox10
        Static speed
        If pbox.Top <= -pbox.Width Then
            speed = CInt(Int(10 * Rnd()) + 5)
            Timer10.Interval = CInt(Int(3000 * Rnd())) + 1000
            PictureBox10.Visible = False
            Drop_bomb(PictureBox10, speed)
            Exit Sub
        End If
        PictureBox10.Visible = True
        Timer10.Interval = 10
        Drop_bomb(PictureBox10, speed)
    End Sub

    Private Sub MoveTimer_Tick(sender As Object, e As EventArgs) Handles MoveTimer.Tick
        '避免同時點擊和鍵盤
        KeyTimer.Stop()
        '圖片初始化
        PictureBox1.BackgroundImage = Nothing
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

        If mouseX - (PictureBox1.Left + PictureBox1.Width / 2) > 15 Then
            PictureBox1.Left += 15
            If Stay_A = 1 Then
                PictureBox1.Image = My.Resources.ResourceManager.GetObject("runner")
                Stay_A = 0
            End If
        ElseIf mouseX - (PictureBox1.Left + PictureBox1.Width / 2) < -15 Then
            PictureBox1.Left -= 15
            If Stay_A = 1 Then
                PictureBox1.Image = My.Resources.ResourceManager.GetObject("runner_reverse")
                Stay_A = 0
            End If
        Else
            '有可能刪除
            'PictureBox1.BackgroundImage = My.Resources.ResourceManager.GetObject("runner")
            'PictureBox1.Image = Nothing
            MoveTimer.Stop()
        End If


    End Sub

    Private Sub KeyTimer_Tick(sender As Object, e As EventArgs) Handles KeyTimer.Tick
        '避免同時點擊和鍵盤
        MoveTimer.Stop()
        '圖片初始化
        PictureBox1.BackgroundImage = Nothing
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

        If (KCode = Keys.Left Or KCode = Keys.A) AndAlso PictureBox1.Left >= 0 Then
            PictureBox1.Left -= 10
            If Stay = 1 Then
                PictureBox1.Image = My.Resources.ResourceManager.GetObject("runner_reverse")
                Stay = 0
            End If
        ElseIf (KCode = Keys.Right Or KCode = Keys.D) AndAlso PictureBox1.Left <= Me.Width - PictureBox1.Width Then
            PictureBox1.Left += 10
            If Stay = 1 Then
                PictureBox1.Image = My.Resources.ResourceManager.GetObject("runner")
                Stay = 0
            End If

            'PictureBox1.BackgroundImage = My.Resources.ResourceManager.GetObject("runner")
            'PictureBox1.Image = Nothing
        End If
    End Sub
End Class
