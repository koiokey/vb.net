'產生form之間公用數值
Imports System.Drawing

Module ShareMoney
    Public SMMoney As Integer
End Module

Public Class Form1

#Region "全域變數"
    Dim pickCrCount As Integer = 0
    Public Shared mainCrPick(4) As Integer
    '總玩家人數
    Public Total_player As Integer = 0
    'Loading some dic for button clicked orelse sound
    Dim Bclicked As New SOUND
    Dim Dclicked As New SOUND
#Region "首頁"
    Dim OpenBegin As New Button
    Dim LeaveGane As New Button
    Dim Adjustment As New Button
#End Region
#Region "選角畫面"
    Dim clearCr As New Button
    Dim gochoiseGame As New Button
    Dim backBigin As New Button

    Dim pickCr1 As New PictureBox
    Dim pickCr2 As New PictureBox
    Dim pickCr3 As New PictureBox
    Dim pickCr4 As New PictureBox
    Dim pickCr5 As New PictureBox

    Dim PlayerPick1 As New PictureBox
    Dim PlayerPick2 As New PictureBox
    Dim PlayerPick3 As New PictureBox
    Dim PlayerPick4 As New PictureBox
    Dim PlayerPick5 As New PictureBox
#End Region
#Region "選擇條件"
    Dim backCr As New Button
    Dim goMainGame As New Button

    Dim leftEndDay As New Button
    Dim leftBeginCion As New Button
    Dim leftEndCion As New Button
    Dim rightEndDay As New Button
    Dim rightBeginCion As New Button
    Dim rightEndCoin As New Button
    Dim EndCoin As New Button
    Dim EndDay As New Button
    Dim BeginCion As New Button

    Public Shared EndCoinNum As Integer
    Public Shared EndDayNum As Integer
    Public Shared BeginCionNum As Integer
#End Region
    '抓客戶電腦畫面
    Dim screenRectangle As Rectangle = Screen.PrimaryScreen.WorkingArea
    Dim x As Integer = (screenRectangle.Width - Me.Width) \ 2
    Dim y As Integer = (screenRectangle.Height - Me.Height) \ 2
    Dim x1 As Integer = (screenRectangle.Width - Me.Width) \ 3
    Dim y1 As Integer = (screenRectangle.Height - Me.Height)

#End Region
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Loading Background music
        My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("Wallpaper"),
                               AudioPlayMode.BackgroundLoop)

        Me.Icon = My.Resources.ResourceManager.GetObject("icon")
        Me.Text = "地產大亨：臺灣之旅"
        EndCoinNum = 80000
        BeginCionNum = 15000
        EndDayNum = 60
        hidePickCr()
        hidechoose()
        '檢測屏幕分辨率， 錯誤則退出
        If Screen.PrimaryScreen.Bounds.Width <> 1920 OrElse
        Screen.PrimaryScreen.Bounds.Height <> 1080 Then
            MsgBox("請在設定中切換分辨率為1920x1080!" &
                   vbCrLf & "如果還是錯誤請將比例調整為100%!", , "Screen Error")
            Application.Exit()
        End If
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        '載入地圖物件
        Generate_map_elements()
        Hidegame()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("beginScience")
        Me.BackgroundImageLayout = ImageLayout.Stretch
#Region "動態生成 首頁"
        'OpenBegin
        AddHandler OpenBegin.Click, AddressOf OpenBegin_Click
        Me.Controls.Add(OpenBegin)
        Me.OpenBegin.Name = "OpenBegin"
        Me.OpenBegin.Size = New System.Drawing.Size(281, 37)
        Me.OpenBegin.TabIndex = 1
        Me.OpenBegin.Text = "開始遊戲"
        Me.OpenBegin.UseVisualStyleBackColor = True
        '
        'LeaveGane
        AddHandler LeaveGane.Click, AddressOf LeaveGane_Click
        Me.Controls.Add(LeaveGane)
        Me.LeaveGane.Name = "LeaveGane"
        Me.LeaveGane.Size = New System.Drawing.Size(281, 33)
        Me.LeaveGane.TabIndex = 2
        Me.LeaveGane.Text = "離開遊戲"
        Me.LeaveGane.UseVisualStyleBackColor = True
        '
        'Adjustment
        AddHandler Adjustment.Click, AddressOf Adjustment_Click
        Me.Controls.Add(Adjustment)
        Me.Adjustment.Name = "Adjustment"
        Me.Adjustment.Size = New System.Drawing.Size(281, 37)
        Me.Adjustment.TabIndex = 3
        Me.Adjustment.Text = "選項"
        Me.Adjustment.UseVisualStyleBackColor = True
        '按鈕位置
        OpenBegin.Location = New Point(x, y + 150)
        LeaveGane.Location = New Point(x, y + 200)
        Adjustment.Location = New Point(x, y + 250)
#End Region
#Region "動態生成 選角畫面"
#Region "選角按鈕 上傳圖片"
        With pickCr1
            .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_1")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With
        With pickCr2
            .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_2")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With
        With pickCr3
            .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_3")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With
        With pickCr4
            .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_4")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With
        With pickCr5
            .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_5")
            .BackgroundImageLayout = ImageLayout.Stretch
        End With
#End Region
        'clearCr
        AddHandler clearCr.Click, AddressOf clearCr_Click
        Me.Controls.Add(clearCr)
        Me.clearCr.Location = New System.Drawing.Point(x1 + 750, y1 + 75)
        Me.clearCr.Name = "clearCr"
        Me.clearCr.Text = "清除角色"
        Me.clearCr.Size = New System.Drawing.Size(100, 40)
        Me.clearCr.TabIndex = 13
        Me.clearCr.UseVisualStyleBackColor = True
        '
        'gochoiseGame
        AddHandler gochoiseGame.Click, AddressOf gochoiseGame_Click
        Me.Controls.Add(gochoiseGame)
        gochoiseGame.BackColor = Color.Transparent
        gochoiseGame.FlatStyle = FlatStyle.Popup
        Me.gochoiseGame.Location = New System.Drawing.Point(x1 + 1150, y1 - 215)
        Me.gochoiseGame.Name = "gochoiseGame"
        Me.gochoiseGame.Size = New System.Drawing.Size(215, 130)
        Me.gochoiseGame.TabIndex = 14
        Me.gochoiseGame.UseVisualStyleBackColor = True
        '
        'backBigin
        AddHandler backBigin.Click, AddressOf backBigin_Click
        Me.Controls.Add(backBigin)
        backBigin.BackColor = Color.Transparent
        backBigin.FlatStyle = FlatStyle.Popup
        Me.backBigin.Location = New System.Drawing.Point(x1 - 425, y1 - 215)
        Me.backBigin.Name = "backBigin"
        Me.backBigin.Size = New System.Drawing.Size(130, 130)
        Me.backBigin.TabIndex = 15
        Me.backBigin.UseVisualStyleBackColor = True
        '
        AddHandler pickCr1.Click, AddressOf pickCr1_Click
        AddHandler pickCr1.MouseEnter, AddressOf pickCr1_Enter
        AddHandler pickCr1.MouseLeave, AddressOf pickCr1_leave
        Me.Controls.Add(pickCr1)
        Me.pickCr1.BackColor = Color.Transparent
        Me.pickCr1.Location = New System.Drawing.Point(x1, y1 + 50)
        Me.pickCr1.Name = "pickCr1"
        Me.pickCr1.Size = New System.Drawing.Size(100, 100)
        Me.pickCr1.TabIndex = 8
        Me.pickCr1.TabStop = False

        AddHandler pickCr2.Click, AddressOf pickCr2_Click
        AddHandler pickCr2.MouseEnter, AddressOf pickCr2_Enter
        AddHandler pickCr2.MouseLeave, AddressOf pickCr2_leave
        Me.Controls.Add(pickCr2)
        Me.pickCr2.BackColor = Color.Transparent
        Me.pickCr2.Location = New System.Drawing.Point(x1 + 295, y1 + 50)
        Me.pickCr2.Name = "pickCr2"
        Me.pickCr2.Size = New System.Drawing.Size(100, 100)
        Me.pickCr2.TabIndex = 9
        Me.pickCr2.TabStop = False

        AddHandler pickCr3.Click, AddressOf pickCr3_Click
        AddHandler pickCr3.MouseEnter, AddressOf pickCr3_Enter
        AddHandler pickCr3.MouseLeave, AddressOf pickCr3_leave
        Me.Controls.Add(pickCr3)
        Me.pickCr3.BackColor = Color.Transparent
        Me.pickCr3.Location = New System.Drawing.Point(x1 + 575, y1 + 50)
        Me.pickCr3.Name = "PickCr3"
        Me.pickCr3.Size = New System.Drawing.Size(100, 100)
        Me.pickCr3.TabIndex = 10
        Me.pickCr3.TabStop = False

        AddHandler pickCr4.Click, AddressOf pickCr4_Click
        AddHandler pickCr4.MouseEnter, AddressOf pickCr4_Enter
        AddHandler pickCr4.MouseLeave, AddressOf pickCr4_leave
        Me.Controls.Add(pickCr4)
        Me.pickCr4.BackColor = Color.Transparent
        Me.pickCr4.Location = New System.Drawing.Point(x1 + 147, y1 + 150)
        Me.pickCr4.Name = "pickCr4"
        Me.pickCr4.Size = New System.Drawing.Size(100, 100)
        Me.pickCr4.TabIndex = 11
        Me.pickCr4.TabStop = False

        AddHandler pickCr5.Click, AddressOf pickCr5_Click
        AddHandler pickCr5.MouseEnter, AddressOf pickCr5_Enter
        AddHandler pickCr5.MouseLeave, AddressOf pickCr5_leave
        Me.pickCr5.BackColor = Color.Transparent
        Me.Controls.Add(pickCr5)
        Me.pickCr5.Location = New System.Drawing.Point(x1 + 442, y1 + 150)
        Me.pickCr5.Name = "pickCr5"
        Me.pickCr5.Size = New System.Drawing.Size(100, 100)
        Me.pickCr5.TabIndex = 12
        Me.pickCr5.TabStop = False

        'PlayerPick1
        Me.Controls.Add(PlayerPick1)
        Me.PlayerPick2.BackColor = Color.Transparent
        Me.PlayerPick1.BackColor = System.Drawing.SystemColors.Control
        Me.PlayerPick1.Location = New System.Drawing.Point(x1 - PlayerPick1.Width * 1.7, y1 - PlayerPick1.Height * 5.4)
        Me.PlayerPick1.Name = "PlayerPick1"
        Me.PlayerPick1.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick1.TabIndex = 4
        Me.PlayerPick1.TabStop = False
        '

        'PlayerPick2
        Me.Controls.Add(PlayerPick2)
        Me.PlayerPick1.BackColor = Color.Transparent
        Me.PlayerPick2.Location = New System.Drawing.Point(x1 + PlayerPick2.Width * 1.7, y1 - PlayerPick2.Height * 5.4)
        Me.PlayerPick2.Name = "PlayerPick2"
        Me.PlayerPick2.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick2.TabIndex = 5
        Me.PlayerPick2.TabStop = False
        '
        'PlayerPick3
        Controls.Add(PlayerPick3)
        Me.PlayerPick3.BackColor = Color.Transparent
        Me.PlayerPick3.Location = New System.Drawing.Point(x1 + PlayerPick3.Width * 5.3, y1 - PlayerPick3.Height * 5.4)
        Me.PlayerPick3.Name = "PlayerPick3"
        Me.PlayerPick3.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick3.TabIndex = 6
        Me.PlayerPick3.TabStop = False
        '
        'PlayerPick4
        Me.Controls.Add(PlayerPick4)
        Me.PlayerPick4.BackColor = Color.Transparent
        Me.PlayerPick4.Location = New System.Drawing.Point(x1 + PlayerPick4.Width * 8.8, y1 - PlayerPick4.Height * 5.4)
        Me.PlayerPick4.Name = "PlayerPick4"
        Me.PlayerPick4.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick4.TabIndex = 7
        Me.PlayerPick4.TabStop = False
        '
#End Region
#Region "動態生成 選擇獲勝條件"

        AddHandler goMainGame.Click, AddressOf goMainGame_Click
        Me.Controls.Add(goMainGame)
        goMainGame.BackColor = Color.Transparent
        goMainGame.FlatStyle = FlatStyle.Popup
        goMainGame.Location = New System.Drawing.Point(x + goMainGame.Width * 9.9, y - goMainGame.Height * 14.4)
        goMainGame.Size = New System.Drawing.Size(330, 190)
        goMainGame.TabIndex = 19
        goMainGame.UseVisualStyleBackColor = True

        AddHandler backCr.Click, AddressOf backCr_Click
        Me.Controls.Add(backCr)
        backCr.BackColor = Color.Transparent
        backCr.FlatStyle = FlatStyle.Popup
        backCr.Location = New System.Drawing.Point(x - backCr.Width * 9, y + backCr.Height * 17.8)
        backCr.Size = New System.Drawing.Size(150, 150)
        backCr.TabIndex = 19
        backCr.UseVisualStyleBackColor = True

        AddHandler leftBeginCion.Click, AddressOf leftBeginCion_Click
        Me.Controls.Add(leftBeginCion)
        leftBeginCion.BackColor = Color.Transparent
        leftBeginCion.FlatStyle = FlatStyle.Popup
        leftBeginCion.Location = New System.Drawing.Point(x - leftBeginCion.Width, y - leftBeginCion.Height * 1.9)
        leftBeginCion.Size = New System.Drawing.Size(280, 90)
        leftBeginCion.TabIndex = 19
        leftBeginCion.UseVisualStyleBackColor = True

        AddHandler leftEndDay.Click, AddressOf leftEndDay_Click
        Me.Controls.Add(leftEndDay)
        leftEndDay.BackColor = Color.Transparent
        leftEndDay.FlatStyle = FlatStyle.Popup
        leftEndDay.Location = New System.Drawing.Point(x - leftEndDay.Width, y + leftEndDay.Height * 5.2)
        leftEndDay.Size = New System.Drawing.Size(280, 90)
        leftEndDay.TabIndex = 19
        leftEndDay.UseVisualStyleBackColor = True

        AddHandler leftEndCion.Click, AddressOf leftEndCion_Click
        Me.Controls.Add(leftEndCion)
        leftEndCion.BackColor = Color.Transparent
        leftEndCion.FlatStyle = FlatStyle.Popup
        leftEndCion.Location = New System.Drawing.Point(x - leftEndCion.Width, y + leftEndCion.Height * 13.1)
        leftEndCion.Size = New System.Drawing.Size(280, 90)
        leftEndCion.TabIndex = 19
        leftEndCion.UseVisualStyleBackColor = True

        AddHandler rightBeginCion.Click, AddressOf rightBeginCion_Click
        Me.Controls.Add(rightBeginCion)
        rightBeginCion.BackColor = Color.Transparent
        rightBeginCion.FlatStyle = FlatStyle.Popup
        rightBeginCion.Location = New System.Drawing.Point(x + rightBeginCion.Width * 8.5, y - rightBeginCion.Height * 1.9)
        rightBeginCion.Size = New System.Drawing.Size(280, 90)
        rightBeginCion.TabIndex = 19
        rightBeginCion.UseVisualStyleBackColor = True

        AddHandler rightEndDay.Click, AddressOf rightEndDay_Click
        Me.Controls.Add(rightEndDay)
        rightEndDay.BackColor = Color.Transparent
        rightEndDay.FlatStyle = FlatStyle.Popup
        rightEndDay.Location = New System.Drawing.Point(x + rightEndDay.Width * 8.5, y + rightEndDay.Height * 5.1)
        rightEndDay.Size = New System.Drawing.Size(280, 90)
        rightEndDay.TabIndex = 19
        rightEndDay.UseVisualStyleBackColor = True



        AddHandler rightEndCoin.Click, AddressOf rightEndCoin_Click
        Me.Controls.Add(rightEndCoin)
        rightEndCoin.BackColor = Color.Transparent
        rightEndCoin.FlatStyle = FlatStyle.Popup
        rightEndCoin.Location = New System.Drawing.Point(x + rightEndCoin.Width * 8.5, y + rightEndCoin.Height * 13.1)
        rightEndCoin.Size = New System.Drawing.Size(280, 90)
        rightEndCoin.TabIndex = 19
        rightEndCoin.UseVisualStyleBackColor = True

        AddHandler BeginCion.Click, AddressOf BeginCion_Click
        Me.Controls.Add(BeginCion)
        BeginCion.BackColor = Color.Transparent
        BeginCion.FlatStyle = FlatStyle.Popup
        Me.BeginCion.Location = New System.Drawing.Point(x + BeginCion.Width * 3.7, y - BeginCion.Height * 1.9)
        Me.BeginCion.Name = "BeginCion"
        Me.BeginCion.Size = New System.Drawing.Size(280, 90)
        Me.BeginCion.TabIndex = 24
        Me.BeginCion.TabStop = False

        AddHandler EndDay.Click, AddressOf EndDay_Click
        Me.Controls.Add(EndDay)
        EndDay.BackColor = Color.Transparent
        EndDay.FlatStyle = FlatStyle.Popup
        EndDay.Location = New System.Drawing.Point(x + EndDay.Width * 3.7, y + EndDay.Height * 5.1)
        EndDay.Size = New System.Drawing.Size(280, 90)
        EndDay.TabIndex = 19
        EndDay.TabStop = False

        AddHandler EndCoin.Click, AddressOf EndCoin_Click
        Me.Controls.Add(EndCoin)
        EndCoin.BackColor = Color.Transparent
        EndCoin.FlatStyle = FlatStyle.Popup
        Me.EndCoin.Location = New System.Drawing.Point(x + EndCoin.Width * 3.7, y + EndCoin.Height * 13.1)
        Me.EndCoin.Name = "BeginCion"
        Me.EndCoin.Size = New System.Drawing.Size(280, 90)
        Me.EndCoin.TabIndex = 24
        Me.EndCoin.TabStop = False

        BeginCion.Enabled = False
        EndCoin.Enabled = False
        EndDay.Enabled = False
        BeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        EndCoin.BackColor = Color.FromArgb(128, Color.Gray)
        EndDay.BackColor = Color.FromArgb(128, Color.Gray)
#End Region
    End Sub
    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.WindowState = FormWindowState.Maximized ' 將視窗最大化
        Me.Location = New Point(0, 0) ' 設置視窗位置
    End Sub
#Region "首頁按鈕"
    Private Sub OpenBegin_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 1000
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        '主頁面隱藏
        hideMain()
        showPickCr()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("8a5cbcd740688e2f")
        If pickCrCount >= 2 Then
            gochoiseGame.Show()
        End If

    End Sub

    Private Sub LeaveGane_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub Adjustment_Click(sender As Object, e As EventArgs)

    End Sub
#End Region
#Region "pickCr事件"

#Region "點選'選擇腳色'事件"
    Private Sub pickCr1_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 2000
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        pickCr1.Enabled = False
        mainCrPick(pickCrCount) = 1
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr2_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 2100
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        pickCr2.Enabled = False
        mainCrPick(pickCrCount) = 2
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr3_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 2200
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        pickCr3.Enabled = False
        mainCrPick(pickCrCount) = 3
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr4_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 2300
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        pickCr4.Enabled = False
        mainCrPick(pickCrCount) = 4
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr5_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 2400
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        pickCr5.Enabled = False
        mainCrPick(pickCrCount) = 5
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

#End Region

#Region "滑鼠進選角的畫面互動"
    '滑進'
    Private Sub pickCr1_Enter(sender As Object, e As System.EventArgs)
        Select Case pickCrCount
            Case 0
                With PlayerPick1
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_1")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 1
                With PlayerPick2
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_1")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 2
                With PlayerPick3
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_1")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 3
                With PlayerPick4
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_1")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
        End Select
    End Sub
    Private Sub pickCr2_Enter(sender As Object, e As System.EventArgs)
        Select Case pickCrCount
            Case 0
                With PlayerPick1
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_2")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 1
                With PlayerPick2
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_2")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 2
                With PlayerPick3
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_2")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 3
                With PlayerPick4
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_2")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
        End Select
    End Sub
    Private Sub pickCr3_Enter(sender As Object, e As System.EventArgs)
        Select Case pickCrCount
            Case 0
                With PlayerPick1
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_3")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 1
                With PlayerPick2
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_3")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 2
                With PlayerPick3
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_3")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 3
                With PlayerPick4
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_3")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
        End Select
    End Sub
    Private Sub pickCr4_Enter(sender As Object, e As System.EventArgs)
        Select Case pickCrCount
            Case 0
                With PlayerPick1
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_4")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 1
                With PlayerPick2
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_4")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 2
                With PlayerPick3
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_4")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 3
                With PlayerPick4
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_4")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
        End Select
    End Sub
    Private Sub pickCr5_Enter(sender As Object, e As System.EventArgs)
        Select Case pickCrCount
            Case 0
                With PlayerPick1
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_5")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 1
                With PlayerPick2
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_5")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 2
                With PlayerPick3
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_5")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
            Case 3
                With PlayerPick4
                    .BackgroundImage = My.Resources.ResourceManager.GetObject("bishop_5")
                    .BackgroundImageLayout = ImageLayout.Stretch
                End With
        End Select
    End Sub
    '滑出'
    Private Sub pickCr1_leave(sender As Object, e As System.EventArgs)
        If pickCr1.Enabled = True Then
            Select Case pickCrCount
                Case 0
                    With PlayerPick1
                        .BackgroundImage = Nothing
                    End With
                Case 1
                    With PlayerPick2
                        .BackgroundImage = Nothing
                    End With
                Case 2
                    With PlayerPick3
                        .BackgroundImage = Nothing
                    End With
                Case 3
                    With PlayerPick4
                        .BackgroundImage = Nothing
                    End With
            End Select
        End If
    End Sub
    Private Sub pickCr2_leave(sender As Object, e As System.EventArgs)
        If pickCr2.Enabled = True Then
            Select Case pickCrCount
                Case 0
                    With PlayerPick1
                        .BackgroundImage = Nothing
                    End With
                Case 1
                    With PlayerPick2
                        .BackgroundImage = Nothing
                    End With
                Case 2
                    With PlayerPick3
                        .BackgroundImage = Nothing
                    End With
                Case 3
                    With PlayerPick4
                        .BackgroundImage = Nothing
                    End With
            End Select
        End If
    End Sub
    Private Sub pickCr3_leave(sender As Object, e As System.EventArgs)
        If pickCr3.Enabled = True Then
            Select Case pickCrCount
                Case 0
                    With PlayerPick1
                        .BackgroundImage = Nothing
                    End With
                Case 1
                    With PlayerPick2
                        .BackgroundImage = Nothing
                    End With
                Case 2
                    With PlayerPick3
                        .BackgroundImage = Nothing
                    End With
                Case 3
                    With PlayerPick4
                        .BackgroundImage = Nothing
                    End With
            End Select
        End If
    End Sub
    Private Sub pickCr4_leave(sender As Object, e As System.EventArgs)
        If pickCr4.Enabled = True Then
            Select Case pickCrCount
                Case 0
                    With PlayerPick1
                        .BackgroundImage = Nothing
                    End With
                Case 1
                    With PlayerPick2
                        .BackgroundImage = Nothing
                    End With
                Case 2
                    With PlayerPick3
                        .BackgroundImage = Nothing
                    End With
                Case 3
                    With PlayerPick4
                        .BackgroundImage = Nothing
                    End With
            End Select
        End If
    End Sub
    Private Sub pickCr5_leave(sender As Object, e As System.EventArgs)
        If pickCr5.Enabled = True Then
            Select Case pickCrCount
                Case 0
                    With PlayerPick1
                        .BackgroundImage = Nothing
                    End With
                Case 1
                    With PlayerPick2
                        .BackgroundImage = Nothing
                    End With
                Case 2
                    With PlayerPick3
                        .BackgroundImage = Nothing
                    End With
                Case 3
                    With PlayerPick4
                        .BackgroundImage = Nothing
                    End With
            End Select
        End If
    End Sub
#End Region

#Region "清除選擇腳色"
    Private Sub clearCr_Click(sender As Object, e As EventArgs)
        If pickCrCount >= 0 Then
            If pickCrCount = 2 Then
                gochoiseGame.Hide()
            End If
            If pickCrCount = 5 Then
                pickCrCount -= 1
                Select Case mainCrPick(pickCrCount)
                    Case 1
                        pickCr1.Enabled = True
                    Case 2
                        pickCr2.Enabled = True
                    Case 3
                        pickCr3.Enabled = True
                    Case 4
                        pickCr4.Enabled = True
                    Case 5
                        pickCr5.Enabled = True
                End Select
                pickCrCount -= 1
            Else
                If pickCrCount <> 0 Then
                    pickCrCount -= 1
                End If
            End If

            Select Case mainCrPick(pickCrCount)
                Case 1
                    pickCr1.Enabled = True
                Case 2
                    pickCr2.Enabled = True
                Case 3
                    pickCr3.Enabled = True
                Case 4
                    pickCr4.Enabled = True
                Case 5
                    pickCr5.Enabled = True
            End Select

            Select Case pickCrCount
                Case 0
                    With PlayerPick1
                        .BackgroundImage = Nothing
                    End With
                Case 1
                    With PlayerPick2
                        .BackgroundImage = Nothing
                    End With
                Case 2
                    With PlayerPick3
                        .BackgroundImage = Nothing
                    End With
                Case 3
                    With PlayerPick4
                        .BackgroundImage = Nothing
                    End With
            End Select
        End If
    End Sub
#End Region

#Region "選擇遊戲獲勝條件顯示&&回主頁面"
    Private Sub gochoiseGame_Click(sender As Object, e As EventArgs)
        showchoose()
        hidePickCr()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("0d8cbab0aeb8fccf")
        If pickCrCount = 5 Then
            Total_player = 4
        Else
            Total_player = pickCrCount
        End If
    End Sub
    '回主頁面'
    Private Sub backBigin_Click(sender As Object, e As EventArgs)
        hidePickCr()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("beginScience")
        showMain()
        rePickCr()
    End Sub
#End Region

#End Region

#Region "隱藏 顯示 重製畫面"
    '隱藏葉面'
    Private Sub hidePickCr()
        PlayerPick1.Hide()
        PlayerPick2.Hide()
        PlayerPick3.Hide()
        PlayerPick4.Hide()
        pickCr1.Hide()
        pickCr2.Hide()
        pickCr4.Hide()
        pickCr5.Hide()
        pickCr3.Hide()
        clearCr.Hide()
        backBigin.Hide()
        gochoiseGame.Hide()
    End Sub
    Private Sub hideMain()
        OpenBegin.Hide()
        LeaveGane.Hide()
        Adjustment.Hide()
    End Sub
    Private Sub hidechoose()
        goMainGame.Hide()
        backCr.Hide()
        leftEndDay.Hide()
        leftBeginCion.Hide()
        leftEndCion.Hide()
        EndDay.Hide()
        BeginCion.Hide()
        EndCoin.Hide()
        rightEndDay.Hide()
        rightBeginCion.Hide()
        rightEndCoin.Hide()
    End Sub

    Sub Hidegame()
        map.Hide()
        Dice1.Hide()
        Dice2.Hide()
        point.Hide()
        a.Hide()
        b.Hide()
        c.Hide()
        d.Hide()
        card.Hide()
        space.Hide()
        purchase.Hide()
        house1.Hide()
        house2.Hide()
        landmark.Hide()
        Dicebutton.Hide()
        play_card1.Hide()
        play_card2.Hide()
        play_card3.Hide()
        play_card4.Hide()
        play_money1.Hide()
        play_money2.Hide()
        play_money3.Hide()
        play_money4.Hide()
        play_order1.Hide()
        play_order2.Hide()
        play_order3.Hide()
        play_order4.Hide()
        play_jail_card1.Hide()
        play_jail_card2.Hide()
        play_jail_card3.Hide()
        play_jail_card4.Hide()
        play_land1.Hide()
        play_land2.Hide()
        play_land3.Hide()
        play_land4.Hide()
        Round_label.Hide()
    End Sub


    '顯示頁面'
    Private Sub showPickCr()
        PlayerPick1.Show()
        PlayerPick2.Show()
        PlayerPick3.Show()
        PlayerPick4.Show()
        pickCr1.Show()
        pickCr2.Show()
        pickCr4.Show()
        pickCr5.Show()
        pickCr3.Show()
        backBigin.Show()
        clearCr.Show()
    End Sub
    Private Sub showMain()
        OpenBegin.Show()
        LeaveGane.Show()
        Adjustment.Show()
    End Sub
    Private Sub showchoose()
        goMainGame.Show()
        backCr.Show()
        clearCr.Show()
        leftEndDay.Show()
        leftBeginCion.Show()
        leftEndCion.Show()
        EndDay.Show()
        BeginCion.Show()
        EndCoin.Show()
        rightEndDay.Show()
        rightBeginCion.Show()
        rightEndCoin.Show()
    End Sub

    Sub Showgame()
        map.Show()
        Dice1.Show()
        Dice2.Show()
        point.Show()
        Round_label.Show()

        If pickCrCount = 1 Then
            a.Show()
            b.Show()
            play_card1.Show()
            play_card2.Show()
            play_money1.Show()
            play_money2.Show()
            play_order1.Show()
            play_order2.Show()
            play_jail_card1.Show()
            play_jail_card2.Show()
            play_land1.Show()
            play_land2.Show()
        ElseIf pickCrCount = 2 Then
            a.Show()
            b.Show()
            play_card1.Show()
            play_card2.Show()
            play_money1.Show()
            play_money2.Show()
            play_order1.Show()
            play_order2.Show()
            play_jail_card1.Show()
            play_jail_card2.Show()
            play_land1.Show()
            play_land2.Show()
        ElseIf pickCrCount = 3 Then
            a.Show()
            b.Show()
            c.Show()
            play_card1.Show()
            play_card2.Show()
            play_card3.Show()
            play_money1.Show()
            play_money2.Show()
            play_money3.Show()
            play_order1.Show()
            play_order2.Show()
            play_order3.Show()
            play_jail_card1.Show()
            play_jail_card2.Show()
            play_jail_card3.Show()
            play_land1.Show()
            play_land2.Show()
            play_land3.Show()
        Else
            a.Show()
            b.Show()
            c.Show()
            d.Show()
            play_card1.Show()
            play_card2.Show()
            play_card3.Show()
            play_card4.Show()
            play_money1.Show()
            play_money2.Show()
            play_money3.Show()
            play_money4.Show()
            play_order1.Show()
            play_order2.Show()
            play_order3.Show()
            play_order4.Show()
            play_jail_card1.Show()
            play_jail_card2.Show()
            play_jail_card3.Show()
            play_jail_card4.Show()
            play_land1.Show()
            play_land2.Show()
            play_land3.Show()
            play_land4.Show()
        End If
        card.Show()

        Dicebutton.Show()
    End Sub
    '重製頁面'
    Private Sub rePickCr()
        pickCrCount = 0
        pickCr1.Enabled = True
        pickCr2.Enabled = True
        pickCr3.Enabled = True
        pickCr4.Enabled = True
        pickCr5.Enabled = True
        With PlayerPick1
            .BackgroundImage = Nothing
        End With
        With PlayerPick2
            .BackgroundImage = Nothing
        End With
        With PlayerPick3
            .BackgroundImage = Nothing
        End With
        With PlayerPick4
            .BackgroundImage = Nothing
        End With
    End Sub


#End Region

#Region "選擇獲勝條件事件"
    Private Sub goMainGame_Click(sender As Object, e As EventArgs)
        'Stop begin music
        My.Computer.Audio.Stop()

        'For button click sound
        Static count = 4000
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        hidechoose()
        Clear_information()
        Showgame()
        a.Money = BeginCionNum '要改 目前測試用
        b.Money = BeginCionNum
        c.Money = BeginCionNum
        d.Money = BeginCionNum
        Player_Update()
        '角色設定
        Select Case Total_player
            Case 2
                a.Survival_statestatus = 1
                b.Survival_statestatus = 1
            Case 3
                a.Survival_statestatus = 1
                b.Survival_statestatus = 1
                c.Survival_statestatus = 1
            Case 4
                a.Survival_statestatus = 1
                b.Survival_statestatus = 1
                c.Survival_statestatus = 1
                d.Survival_statestatus = 1
        End Select

        If mainCrPick(0) = 1 Then
            a.Image = My.Resources.ResourceManager.GetObject("bishop_1")
            play_card1.Image = My.Resources.ResourceManager.GetObject("playercard_red")
        ElseIf mainCrPick(0) = 2 Then
            a.Image = My.Resources.ResourceManager.GetObject("bishop_2")
            play_card1.Image = My.Resources.ResourceManager.GetObject("playercard_yellow")
        ElseIf mainCrPick(0) = 3 Then
            a.Image = My.Resources.ResourceManager.GetObject("bishop_3")
            play_card1.Image = My.Resources.ResourceManager.GetObject("playercard_green")
        ElseIf mainCrPick(0) = 4 Then
            a.Image = My.Resources.ResourceManager.GetObject("bishop_4")
            play_card1.Image = My.Resources.ResourceManager.GetObject("playercard_blue")
        ElseIf mainCrPick(0) = 5 Then
            a.Image = My.Resources.ResourceManager.GetObject("bishop_5")
            play_card1.Image = My.Resources.ResourceManager.GetObject("playercard_pink")
        End If

        If mainCrPick(1) = 1 Then
            b.Image = My.Resources.ResourceManager.GetObject("bishop_1")
            play_card2.Image = My.Resources.ResourceManager.GetObject("playercard_red")
        ElseIf mainCrPick(1) = 2 Then
            b.Image = My.Resources.ResourceManager.GetObject("bishop_2")
            play_card2.Image = My.Resources.ResourceManager.GetObject("playercard_yellow")
        ElseIf mainCrPick(1) = 3 Then
            b.Image = My.Resources.ResourceManager.GetObject("bishop_3")
            play_card2.Image = My.Resources.ResourceManager.GetObject("playercard_green")
        ElseIf mainCrPick(1) = 4 Then
            b.Image = My.Resources.ResourceManager.GetObject("bishop_4")
            play_card2.Image = My.Resources.ResourceManager.GetObject("playercard_blue")
        ElseIf mainCrPick(1) = 5 Then
            b.Image = My.Resources.ResourceManager.GetObject("bishop_5")
            play_card2.Image = My.Resources.ResourceManager.GetObject("playercard_pink")
        End If

        If mainCrPick(2) = 1 Then
            c.Image = My.Resources.ResourceManager.GetObject("bishop_1")
            play_card3.Image = My.Resources.ResourceManager.GetObject("playercard_red")
        ElseIf mainCrPick(2) = 2 Then
            c.Image = My.Resources.ResourceManager.GetObject("bishop_2")
            play_card3.Image = My.Resources.ResourceManager.GetObject("playercard_yellow")
        ElseIf mainCrPick(2) = 3 Then
            c.Image = My.Resources.ResourceManager.GetObject("bishop_3")
            play_card3.Image = My.Resources.ResourceManager.GetObject("playercard_green")
        ElseIf mainCrPick(2) = 4 Then
            c.Image = My.Resources.ResourceManager.GetObject("bishop_4")
            play_card3.Image = My.Resources.ResourceManager.GetObject("playercard_blue")
        ElseIf mainCrPick(2) = 5 Then
            c.Image = My.Resources.ResourceManager.GetObject("bishop_5")
            play_card3.Image = My.Resources.ResourceManager.GetObject("playercard_pink")
        End If

        If mainCrPick(3) = 1 Then
            d.Image = My.Resources.ResourceManager.GetObject("bishop_1")
            play_card4.Image = My.Resources.ResourceManager.GetObject("playercard_red")
        ElseIf mainCrPick(3) = 2 Then
            d.Image = My.Resources.ResourceManager.GetObject("bishop_2")
            play_card4.Image = My.Resources.ResourceManager.GetObject("playercard_yellow")
        ElseIf mainCrPick(3) = 3 Then
            d.Image = My.Resources.ResourceManager.GetObject("bishop_3")
            play_card4.Image = My.Resources.ResourceManager.GetObject("playercard_green")
        ElseIf mainCrPick(3) = 4 Then
            d.Image = My.Resources.ResourceManager.GetObject("bishop_4")
            play_card4.Image = My.Resources.ResourceManager.GetObject("playercard_blue")
        ElseIf mainCrPick(3) = 5 Then
            d.Image = My.Resources.ResourceManager.GetObject("bishop_5")
            play_card4.Image = My.Resources.ResourceManager.GetObject("playercard_pink")
        End If
        Me.BackgroundImage = Nothing
    End Sub
    Private Sub backCr_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4100
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        hidechoose()
        showPickCr()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("8a5cbcd740688e2f")
        EndCoinNum = 100000
        EndDayNum = 60
        BeginCionNum = 15000
        If pickCrCount >= 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub leftBeginCion_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4200
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        BeginCionNum = 10000
        leftBeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        BeginCion.BackColor = Color.Transparent
        rightBeginCion.BackColor = Color.Transparent
        leftBeginCion.Enabled = False
        BeginCion.Enabled = True
        rightBeginCion.Enabled = True
    End Sub

    Private Sub BeginCion_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4300
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        BeginCionNum = 15000
        leftBeginCion.BackColor = Color.Transparent
        BeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        rightBeginCion.BackColor = Color.Transparent
        leftBeginCion.Enabled = True
        BeginCion.Enabled = False
        rightBeginCion.Enabled = True
    End Sub
    Private Sub rightBeginCion_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4400
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        BeginCionNum = 20000
        leftBeginCion.BackColor = Color.Transparent
        BeginCion.BackColor = Color.Transparent
        rightBeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        leftBeginCion.Enabled = True
        BeginCion.Enabled = True
        rightBeginCion.Enabled = False
    End Sub
    Private Sub rightEndDay_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4500
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        EndDayNum = 100
        rightEndDay.BackColor = Color.FromArgb(128, Color.Gray)
        EndDay.BackColor = Color.Transparent
        leftEndDay.BackColor = Color.Transparent
        rightEndDay.Enabled = False
        EndDay.Enabled = True
        leftEndDay.Enabled = True
    End Sub
    Private Sub EndDay_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4600
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        EndDayNum = 80
        rightEndDay.BackColor = Color.Transparent
        EndDay.BackColor = Color.FromArgb(128, Color.Gray)
        leftEndDay.BackColor = Color.Transparent
        rightEndDay.Enabled = True
        EndDay.Enabled = False
        leftEndDay.Enabled = True
    End Sub
    Private Sub leftEndDay_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4700
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        EndDayNum = 60
        rightEndDay.BackColor = Color.Transparent
        EndDay.BackColor = Color.Transparent
        leftEndDay.BackColor = Color.FromArgb(128, Color.Gray)
        rightEndDay.Enabled = True
        EndDay.Enabled = True
        leftEndDay.Enabled = False
    End Sub

    Private Sub rightEndCoin_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4800
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        EndCoinNum = 100000
        rightEndCoin.BackColor = Color.FromArgb(128, Color.Gray)
        EndCoin.BackColor = Color.Transparent
        leftEndCion.BackColor = Color.Transparent
        rightEndCoin.Enabled = False
        EndCoin.Enabled = True
        leftEndCion.Enabled = True
    End Sub

    Private Sub EndCoin_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 4900
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        EndCoinNum = 80000
        rightEndCoin.BackColor = Color.Transparent
        EndCoin.BackColor = Color.FromArgb(128, Color.Gray)
        leftEndCion.BackColor = Color.Transparent
        rightEndCoin.Enabled = True
        EndCoin.Enabled = False
        leftEndCion.Enabled = True
    End Sub
    Private Sub leftEndCion_Click(sender As Object, e As EventArgs)
        'For button click sound
        Static count = 5000
        count += 1
        With Bclicked
            .Name = "SOUND" & count
            .Play(3, False)
        End With

        EndCoinNum = 50000
        rightEndCoin.BackColor = Color.Transparent
        EndCoin.BackColor = Color.Transparent
        leftEndCion.BackColor = Color.FromArgb(128, Color.Gray)
        rightEndCoin.Enabled = True
        EndCoin.Enabled = True
        leftEndCion.Enabled = False
    End Sub

#End Region

#Region "產生地圖物件"
    Dim map As PictureBox
    Dim Dice1 As PictureBox
    Dim Dice2 As PictureBox
    Dim point As TextBox
    Dim card As PictureBox
    Dim space As PictureBox
    Dim purchase As PictureBox
    Dim house1 As PictureBox
    Dim house2 As PictureBox
    Dim landmark As PictureBox
    '角色資訊卡
    Dim play_card1 As PictureBox
    Dim play_card2 As PictureBox
    Dim play_card3 As PictureBox
    Dim play_card4 As PictureBox
    '角色資訊
    Dim play_money1 As Label
    Dim play_money2 As Label
    Dim play_money3 As Label
    Dim play_money4 As Label

    Dim play_order1 As Label
    Dim play_order2 As Label
    Dim play_order3 As Label
    Dim play_order4 As Label

    Dim play_jail_card1 As Label
    Dim play_jail_card2 As Label
    Dim play_jail_card3 As Label
    Dim play_jail_card4 As Label

    Public play_land1 As Label
    Public play_land2 As Label
    Public play_land3 As Label
    Public play_land4 As Label

    Dim Round_label As Label

    Friend WithEvents Dicebutton As Button
    Private pictureBoxesList As New List(Of PictureBox) ' 儲存pictureBox 的陣列
    '座標系統陣列
    Public coordinates(32, 1) As Integer
    '建立玩家物件(名稱, 起始位置, 初始金額)
    Public a As New Player("Player 1", 0, BeginCionNum)
    Public b As New Player("Player 2", 0, BeginCionNum)
    Public c As New Player("Player 3", 0, BeginCionNum)
    Public d As New Player("Player 4", 0, BeginCionNum)
    Sub Generate_map_elements()
        Me.ClientSize = New Size(1600, 1000)
        '初始化地圖物件
        map = New PictureBox()
        map.Location = New Point(0, 0)
        map.BackgroundImage = My.Resources.ResourceManager.GetObject("newmap")
        map.BackgroundImageLayout = ImageLayout.Zoom
        map.SizeMode = PictureBoxSizeMode.CenterImage
        map.Dock = DockStyle.Fill
        map.TabIndex = 0
        Me.Controls.Add(map)

        '角色
        Me.Controls.Add(a)

        a.SizeMode = PictureBoxSizeMode.Zoom
        a.BackColor = Color.Transparent
        a.Location = New Point(Me.ClientSize.Width * 0.33, Me.ClientSize.Height * 0.05)
        a.BringToFront()

        Me.Controls.Add(b)
        b.BackColor = Color.Transparent
        'b.Image = My.Resources.ResourceManager.GetObject("bishop_2")
        b.SizeMode = PictureBoxSizeMode.Zoom
        b.Location = New Point(Me.ClientSize.Width * 0.33, Me.ClientSize.Height * 0.05)
        b.BringToFront()

        Me.Controls.Add(c)
        'c.Image = My.Resources.ResourceManager.GetObject("bishop_3")
        c.SizeMode = PictureBoxSizeMode.Zoom
        c.BackColor = Color.Transparent
        c.Location = New Point(Me.ClientSize.Width * 0.33, Me.ClientSize.Height * 0.05)
        c.BringToFront()

        Me.Controls.Add(d)
        'd.Image = My.Resources.ResourceManager.GetObject("bishop_4")
        d.SizeMode = PictureBoxSizeMode.Zoom
        d.BackColor = Color.Transparent
        d.Location = New Point(Me.ClientSize.Width * 0.33, Me.ClientSize.Height * 0.05)
        d.BringToFront()

        '角色去背
        a.Parent = map
        b.Parent = map
        c.Parent = map
        d.Parent = map
        '角色資訊卡(包含金錢、順位、以及監獄豁免卡張數)
        play_card1 = New PictureBox()
        Me.Controls.Add(play_card1)
        play_card1.Location = New Point(Me.ClientSize.Width * 0, Me.ClientSize.Height * 0)
        play_card1.Size = New Size(450, 185)
        play_card1.SizeMode = PictureBoxSizeMode.CenterImage
        play_card1.BringToFront()

        play_money1 = New Label()
        Me.Controls.Add(play_money1)
        play_money1.BackColor = Color.Transparent
        play_money1.Parent = play_card1
        play_money1.Location = New Point(Me.ClientSize.Width * 0.18, Me.ClientSize.Height * 0.05)
        play_money1.Font = New Font("標楷體", 16, FontStyle.Bold)

        play_order1 = New Label()
        Me.Controls.Add(play_order1)
        play_order1.BackColor = Color.Transparent
        play_order1.Parent = play_card1
        play_order1.Location = New Point(Me.ClientSize.Width * 0.1, Me.ClientSize.Height * 0.14)
        play_order1.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_order1.Text = "1"

        play_jail_card1 = New Label()
        Me.Controls.Add(play_jail_card1)
        play_jail_card1.BackColor = Color.Transparent
        play_jail_card1.Parent = play_card1
        play_jail_card1.Location = New Point(Me.ClientSize.Width * 0.21, Me.ClientSize.Height * 0.125)
        play_jail_card1.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_jail_card1.Text = 0

        play_land1 = New Label()
        Me.Controls.Add(play_land1)
        play_land1.Location = New Point(Me.ClientSize.Width * 0.02, Me.ClientSize.Height * 0 + play_card1.Height + 10)
        play_land1.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_land1.AutoSize = False
        play_land1.Width = 420
        play_land1.Height = 200
        play_land1.Text = "玩家一持有土地:"
        play_land1.BringToFront()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        play_card2 = New PictureBox()
        Me.Controls.Add(play_card2)
        play_card2.Location = New Point(Me.ClientSize.Width * 0, Me.ClientSize.Height * 0.25)
        play_card2.Size = New Size(450, 185)
        play_card2.SizeMode = PictureBoxSizeMode.CenterImage
        play_card2.BringToFront()

        play_money2 = New Label()
        Me.Controls.Add(play_money2)
        play_money2.BackColor = Color.Transparent
        play_money2.Parent = play_card2
        play_money2.Location = New Point(Me.ClientSize.Width * 0.18, Me.ClientSize.Height * 0.05)
        play_money2.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_money2.BringToFront()

        play_order2 = New Label()
        Me.Controls.Add(play_order2)
        play_order2.BackColor = Color.Transparent
        play_order2.Parent = play_card2
        play_order2.Location = New Point(Me.ClientSize.Width * 0.1, Me.ClientSize.Height * 0.14)
        play_order2.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_order2.Text = "2"
        play_money2.BringToFront()

        play_jail_card2 = New Label()
        Me.Controls.Add(play_jail_card2)
        play_jail_card2.BackColor = Color.Transparent
        play_jail_card2.Parent = play_card2
        play_jail_card2.Location = New Point(Me.ClientSize.Width * 0.21, Me.ClientSize.Height * 0.125)
        play_jail_card2.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_jail_card2.Text = 0
        play_jail_card2.BringToFront()

        play_land2 = New Label()
        Me.Controls.Add(play_land2)
        play_land2.Location = New Point(Me.ClientSize.Width * 0.02, Me.ClientSize.Height * 0.25 + play_card2.Height + 10)
        play_land2.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_land2.AutoSize = False
        play_land2.Width = 420
        play_land2.Height = 200
        play_land2.Text = "玩家二持有土地:"
        play_land2.BringToFront()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        play_card3 = New PictureBox()
        Me.Controls.Add(play_card3)
        play_card3.Location = New Point(Me.ClientSize.Width * 0, Me.ClientSize.Height * 0.5)
        play_card3.Size = New Size(450, 185)
        play_card3.SizeMode = PictureBoxSizeMode.CenterImage
        play_card3.BringToFront()

        play_money3 = New Label()
        Me.Controls.Add(play_money3)
        play_money3.BackColor = Color.Transparent
        play_money3.Parent = play_card3
        play_money3.Location = New Point(Me.ClientSize.Width * 0.18, Me.ClientSize.Height * 0.05)
        play_money3.Font = New Font("標楷體", 16, FontStyle.Bold)

        play_order3 = New Label()
        Me.Controls.Add(play_order3)
        play_order3.BackColor = Color.Transparent
        play_order3.Parent = play_card3
        play_order3.Location = New Point(Me.ClientSize.Width * 0.1, Me.ClientSize.Height * 0.14)
        play_order3.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_order3.Text = "3"

        play_jail_card3 = New Label()
        Me.Controls.Add(play_jail_card3)
        play_jail_card3.BackColor = Color.Transparent
        play_jail_card3.Parent = play_card3
        play_jail_card3.Location = New Point(Me.ClientSize.Width * 0.21, Me.ClientSize.Height * 0.125)
        play_jail_card3.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_jail_card3.Text = 0

        play_land3 = New Label()
        Me.Controls.Add(play_land3)
        play_land3.Location = New Point(Me.ClientSize.Width * 0.02, Me.ClientSize.Height * 0.5 + play_card3.Height + 10)
        play_land3.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_land3.AutoSize = False
        play_land3.Width = 420
        play_land3.Height = 200
        play_land3.Text = "玩家三持有土地:"
        play_land3.BringToFront()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        play_card4 = New PictureBox()
        Me.Controls.Add(play_card4)
        play_card4.Location = New Point(Me.ClientSize.Width * 0, Me.ClientSize.Height * 0.75)
        play_card4.Size = New Size(450, 185)
        play_card4.SizeMode = PictureBoxSizeMode.CenterImage
        play_card4.BringToFront()

        play_money4 = New Label()
        Me.Controls.Add(play_money4)
        play_money4.BackColor = Color.Transparent
        play_money4.Parent = play_card4
        play_money4.Location = New Point(Me.ClientSize.Width * 0.18, Me.ClientSize.Height * 0.05)
        play_money4.Font = New Font("標楷體", 16, FontStyle.Bold)

        play_order4 = New Label()
        Me.Controls.Add(play_order4)
        play_order4.BackColor = Color.Transparent
        play_order4.Parent = play_card4
        play_order4.Location = New Point(Me.ClientSize.Width * 0.1, Me.ClientSize.Height * 0.14)
        play_order4.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_order4.Text = "4"

        play_jail_card4 = New Label()
        Me.Controls.Add(play_jail_card4)
        play_jail_card4.BackColor = Color.Transparent
        play_jail_card4.Parent = play_card4
        play_jail_card4.Location = New Point(Me.ClientSize.Width * 0.21, Me.ClientSize.Height * 0.125)
        play_jail_card4.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_jail_card4.Text = 0

        play_land4 = New Label()
        Me.Controls.Add(play_land4)
        play_land4.Location = New Point(Me.ClientSize.Width * 0.02, Me.ClientSize.Height * 0.75 + play_card4.Height + 10)
        play_land4.Font = New Font("標楷體", 16, FontStyle.Bold)
        play_land4.AutoSize = False
        play_land4.Width = 420
        play_land4.Height = 200
        play_land4.Text = "玩家四持有土地:"
        play_land4.BringToFront()

        Round_label = New Label()
        Me.Controls.Add(Round_label)
        Round_label.BackColor = Color.Transparent
        Round_label.Parent = map
        Round_label.Font = New Font("標楷體", 16, FontStyle.Bold)
        Round_label.Location = New Point(Me.ClientSize.Width * 0.7, Me.ClientSize.Height * 0.2)
        Round_label.Text = "回合數 ：" & " " & Round
        Round_label.AutoSize = True
        Round_label.BringToFront()
        '骰子1
        Dice1 = New PictureBox()
        Dice1.Size = New Size(90, 90)
        Dice1.Location = New Point(Me.ClientSize.Width * 0.5, Me.ClientSize.Height * 0.35)
        Dice1.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(Dice1)
        Dice1.Parent = map
        Dice1.BackColor = Color.Transparent
        Dice1.BringToFront()
        '骰子2
        Dice2 = New PictureBox()
        Dice2.Size = New Size(90, 90)
        Dice2.Location = New Point(Me.ClientSize.Width * 0.65, Me.ClientSize.Height * 0.35)
        Dice2.SizeMode = PictureBoxSizeMode.Zoom
        Me.Controls.Add(Dice2)
        Dice2.Parent = map
        Dice2.BackColor = Color.Transparent
        Dice2.BringToFront()
        '點數和
        point = New TextBox()
        point.Size = New Size(60, 40)
        point.Location = New Point(Me.ClientSize.Width * 0.58, Me.ClientSize.Height * 0.5)
        point.Enabled = False
        point.TextAlign = HorizontalAlignment.Center
        point.Font = New Font("標楷體", 16, FontStyle.Bold)
        point.TextAlign = HorizontalAlignment.Center
        point.BringToFront()
        Me.Controls.Add(point)
        point.Parent = map
        point.BringToFront()
        '擲骰按鈕
        Dicebutton = New Button()
        Dicebutton.Size = New Size(80, 45)
        Dicebutton.Location = New Point(Me.ClientSize.Width * 0.575, Me.ClientSize.Height * 0.6)
        Dicebutton.Text = "擲骰"
        Dicebutton.TabStop = False
        Me.Controls.Add(Dicebutton)
        Dicebutton.BringToFront()
        '土地
        Kinmen.Initialize()
        Mazu.Initialize()
        Liuchiu.Initialize()
        Koto_island.Initialize()
        Lyudao.Initialize()
        Penghu.Initialize()
        Taitung.Initialize()
        Hualien.Initialize()
        Pingtung.Initialize()
        Kaohsiung.Initialize()
        Tainan.Initialize()
        Chiayi.Initialize()
        Nantou.Initialize()
        Yunlin.Initialize()
        Changhua.Initialize()
        Taichung.Initialize()
        Miaoli.Initialize()
        Yilan.Initialize()
        Hsinchu.Initialize()
        Taoyuan.Initialize()
        New_Taipei.Initialize()
        Taipei.Initialize()
        Keelung.Initialize()
        'PictureBox
        card = New PictureBox()
        With card
            .Size = New Size(400, 600)
            .Location = New Point(Me.ClientSize.Width * 0.93, Me.ClientSize.Height * 0.2)
            .SizeMode = PictureBoxSizeMode.CenterImage
        End With
        Me.Controls.Add(card)
        card.BringToFront()

        space = New PictureBox()
        With space
            .Size = New Size(50, 50)
            .Location = New Point(Me.ClientSize.Width * 0.93, Me.ClientSize.Height * 0.13)
            .SizeMode = PictureBoxSizeMode.StretchImage
        End With
        space.Image = My.Resources.ResourceManager.GetObject("space")
        Me.Controls.Add(space)
        space.BringToFront()

        purchase = New PictureBox()
        With purchase
            .Size = New Size(50, 50)
            .Location = New Point(Me.ClientSize.Width * 0.93, Me.ClientSize.Height * 0.13)
            .SizeMode = PictureBoxSizeMode.StretchImage
        End With
        purchase.Image = My.Resources.ResourceManager.GetObject("purchase")
        Me.Controls.Add(purchase)
        purchase.BringToFront()

        house1 = New PictureBox()
        With house1
            .Size = New Size(50, 50)
            .Location = New Point(Me.ClientSize.Width * 0.97, Me.ClientSize.Height * 0.13)
            .SizeMode = PictureBoxSizeMode.StretchImage
        End With
        house1.Image = My.Resources.ResourceManager.GetObject("house")
        Me.Controls.Add(house1)
        house1.BringToFront()

        house2 = New PictureBox()
        With house2
            .Size = New Size(50, 50)
            .Location = New Point(Me.ClientSize.Width * 1.01, Me.ClientSize.Height * 0.13)
            .SizeMode = PictureBoxSizeMode.StretchImage
        End With
        house2.Image = My.Resources.ResourceManager.GetObject("house")
        Me.Controls.Add(house2)
        house2.BringToFront()

        landmark = New PictureBox()
        With landmark
            .Size = New Size(50, 50)
            .Location = New Point(Me.ClientSize.Width * 1.14, Me.ClientSize.Height * 0.13)
            .SizeMode = PictureBoxSizeMode.StretchImage
        End With
        landmark.Image = My.Resources.ResourceManager.GetObject("landmark")
        Me.Controls.Add(landmark)
        landmark.BringToFront()
        '地價卡圖片
        For i As Integer = 1 To 11
            Dim newPictureBox As New PictureBox()
            With newPictureBox
                .Size = New Size(95, 130)
                .BackColor = Color.Transparent
            End With
            Me.Controls.Add(newPictureBox)
            newPictureBox.Parent = map
            pictureBoxesList.Add(newPictureBox)
            newPictureBox.BringToFront()
            AddHandler newPictureBox.Click, AddressOf PictureBox_Click
        Next

        For i As Integer = 1 To 12
            Dim newPictureBox As New PictureBox()
            With newPictureBox
                .Size = New Size(130, 95)
                .BackColor = Color.Transparent
            End With
            Me.Controls.Add(newPictureBox)
            pictureBoxesList.Add(newPictureBox)
            newPictureBox.Parent = map
            newPictureBox.BringToFront()
            AddHandler newPictureBox.Click, AddressOf PictureBox_Click
        Next

        pictureBoxesList(0).Location = New Point(Me.ClientSize.Width * 0.393, Me.ClientSize.Height * 0.035)
        pictureBoxesList(1).Location = New Point(Me.ClientSize.Width * 0.455, Me.ClientSize.Height * 0.035)
        pictureBoxesList(2).Location = New Point(Me.ClientSize.Width * 0.57, Me.ClientSize.Height * 0.035)
        pictureBoxesList(3).Location = New Point(Me.ClientSize.Width * 0.686, Me.ClientSize.Height * 0.035)
        pictureBoxesList(4).Location = New Point(Me.ClientSize.Width * 0.747, Me.ClientSize.Height * 0.035)

        pictureBoxesList(5).Location = New Point(Me.ClientSize.Width * 0.747, Me.ClientSize.Height * 0.834)
        pictureBoxesList(6).Location = New Point(Me.ClientSize.Width * 0.686, Me.ClientSize.Height * 0.834)
        pictureBoxesList(7).Location = New Point(Me.ClientSize.Width * 0.63, Me.ClientSize.Height * 0.834)
        pictureBoxesList(8).Location = New Point(Me.ClientSize.Width * 0.57, Me.ClientSize.Height * 0.834)
        pictureBoxesList(9).Location = New Point(Me.ClientSize.Width * 0.455, Me.ClientSize.Height * 0.834)
        pictureBoxesList(10).Location = New Point(Me.ClientSize.Width * 0.393, Me.ClientSize.Height * 0.834)

        pictureBoxesList(11).Location = New Point(Me.ClientSize.Width * 0.81, Me.ClientSize.Height * 0.168)
        pictureBoxesList(12).Location = New Point(Me.ClientSize.Width * 0.81, Me.ClientSize.Height * 0.263)
        pictureBoxesList(13).Location = New Point(Me.ClientSize.Width * 0.81, Me.ClientSize.Height * 0.358)
        pictureBoxesList(14).Location = New Point(Me.ClientSize.Width * 0.81, Me.ClientSize.Height * 0.453)
        pictureBoxesList(15).Location = New Point(Me.ClientSize.Width * 0.81, Me.ClientSize.Height * 0.635)
        pictureBoxesList(16).Location = New Point(Me.ClientSize.Width * 0.81, Me.ClientSize.Height * 0.733)

        pictureBoxesList(17).Location = New Point(Me.ClientSize.Width * 0.31, Me.ClientSize.Height * 0.168)
        pictureBoxesList(18).Location = New Point(Me.ClientSize.Width * 0.31, Me.ClientSize.Height * 0.263)
        pictureBoxesList(19).Location = New Point(Me.ClientSize.Width * 0.31, Me.ClientSize.Height * 0.45)
        pictureBoxesList(20).Location = New Point(Me.ClientSize.Width * 0.31, Me.ClientSize.Height * 0.54)
        pictureBoxesList(21).Location = New Point(Me.ClientSize.Width * 0.31, Me.ClientSize.Height * 0.635)
        pictureBoxesList(22).Location = New Point(Me.ClientSize.Width * 0.31, Me.ClientSize.Height * 0.733)

        DiceTimer1.Interval = 100 ' 調整計時器的間隔，控制動畫速度
        DiceTimer2.Interval = 125 ' 第二顆骰子的計時器間隔

        coordinates(0, 0) = Me.ClientSize.Width * 0.33 : coordinates(1, 0) = Me.ClientSize.Width * 0.41 : coordinates(2, 0) = Me.ClientSize.Width * 0.47 : coordinates(3, 0) = Me.ClientSize.Width * 0.53 : coordinates(4, 0) = Me.ClientSize.Width * 0.59 : coordinates(5, 0) = Me.ClientSize.Width * 0.645 : coordinates(6, 0) = Me.ClientSize.Width * 0.705 : coordinates(7, 0) = Me.ClientSize.Width * 0.765 : coordinates(8, 0) = Me.ClientSize.Width * 0.85
        coordinates(9, 1) = Me.ClientSize.Height * 0.18 : coordinates(10, 1) = Me.ClientSize.Height * 0.27 : coordinates(11, 1) = Me.ClientSize.Height * 0.37 : coordinates(12, 1) = Me.ClientSize.Height * 0.46 : coordinates(13, 1) = Me.ClientSize.Height * 0.56 : coordinates(14, 1) = Me.ClientSize.Height * 0.65 : coordinates(15, 1) = Me.ClientSize.Height * 0.74 : coordinates(16, 1) = Me.ClientSize.Height * 0.89
        coordinates(17, 0) = Me.ClientSize.Width * 0.765 : coordinates(18, 0) = Me.ClientSize.Width * 0.705 : coordinates(19, 0) = Me.ClientSize.Width * 0.645 : coordinates(20, 0) = Me.ClientSize.Width * 0.59 : coordinates(21, 0) = Me.ClientSize.Width * 0.53 : coordinates(22, 0) = Me.ClientSize.Width * 0.47 : coordinates(23, 0) = Me.ClientSize.Width * 0.41 : coordinates(24, 0) = Me.ClientSize.Width * 0.33
        coordinates(25, 1) = Me.ClientSize.Height * 0.74 : coordinates(26, 1) = Me.ClientSize.Height * 0.65 : coordinates(27, 1) = Me.ClientSize.Height * 0.56 : coordinates(28, 1) = Me.ClientSize.Height * 0.46 : coordinates(29, 1) = Me.ClientSize.Height * 0.37 : coordinates(30, 1) = Me.ClientSize.Height * 0.27 : coordinates(31, 1) = Me.ClientSize.Height * 0.18
        For y As Integer = 0 To 8
            coordinates(y, 1) = Me.ClientSize.Height * 0.05
        Next

        For x As Integer = 9 To 16
            coordinates(x, 0) = Me.ClientSize.Width * 0.85
        Next

        For y As Integer = 17 To 24
            coordinates(y, 1) = Me.ClientSize.Height * 0.89
        Next

        For x As Integer = 25 To 31
            coordinates(x, 0) = Me.ClientSize.Width * 0.33
        Next

    End Sub
    Private Sub PictureBox_Click(sender As Object, e As EventArgs)
        Dim clickedPictureBox As PictureBox = CType(sender, PictureBox)
        Dim index As Integer = pictureBoxesList.IndexOf(clickedPictureBox)

        Select Case index
            Case 0 '金門'
                card.Image = My.Resources.ResourceManager.GetObject("Kinmen")
                Select Case Kinmen.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 1
                card.Image = My.Resources.ResourceManager.GetObject("Mazu")
                Select Case Mazu.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 2
                card.Image = My.Resources.ResourceManager.GetObject("Liuchiu")
                Select Case Liuchiu.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 3
                card.Image = My.Resources.ResourceManager.GetObject("Koto_island")
                Select Case Koto_island.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 4
                card.Image = My.Resources.ResourceManager.GetObject("Lyudao")
                Select Case Lyudao.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 5
                card.Image = My.Resources.ResourceManager.GetObject("Chiayi")
                Select Case Chiayi.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 6
                card.Image = My.Resources.ResourceManager.GetObject("Nantou")
                Select Case Nantou.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 7
                card.Image = My.Resources.ResourceManager.GetObject("Yunlin")
                Select Case Yunlin.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 8
                card.Image = My.Resources.ResourceManager.GetObject("Changhua")
                Select Case Changhua.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 9
                card.Image = My.Resources.ResourceManager.GetObject("Taichung")
                Select Case Taichung.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 10
                card.Image = My.Resources.ResourceManager.GetObject("Miaoli")
                Select Case Miaoli.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 11
                card.Image = My.Resources.ResourceManager.GetObject("Penghu")
                Select Case Penghu.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 12
                card.Image = My.Resources.ResourceManager.GetObject("Taitung")
                Select Case Taitung.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 13
                card.Image = My.Resources.ResourceManager.GetObject("Hualien")
                Select Case Hualien.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 14
                card.Image = My.Resources.ResourceManager.GetObject("Pingtung")
                Select Case Pingtung.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 15
                card.Image = My.Resources.ResourceManager.GetObject("Kaohsiung")
                Select Case Kaohsiung.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 16
                card.Image = My.Resources.ResourceManager.GetObject("Tainan")
                Select Case Tainan.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 17
                card.Image = My.Resources.ResourceManager.GetObject("Keelung")
                Select Case Keelung.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 18
                card.Image = My.Resources.ResourceManager.GetObject("Taipei")
                Select Case Taipei.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 19
                card.Image = My.Resources.ResourceManager.GetObject("New_Taipei")
                Select Case New_Taipei.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 20
                card.Image = My.Resources.ResourceManager.GetObject("Taoyuan")
                Select Case Taoyuan.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 21
                card.Image = My.Resources.ResourceManager.GetObject("Hsinchu")
                Select Case Hsinchu.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select
            Case 22
                card.Image = My.Resources.ResourceManager.GetObject("Yilan")
                Select Case Yilan.Level
                    Case -1
                        space.Visible = True
                        purchase.Visible = False
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 0
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = False
                    Case 1
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = False
                        landmark.Visible = False
                    Case 2
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = True
                        house2.Visible = True
                        landmark.Visible = False
                    Case 3
                        space.Visible = False
                        purchase.Visible = True
                        house1.Visible = False
                        house2.Visible = False
                        landmark.Visible = True
                End Select

        End Select
    End Sub

#End Region

#Region "擲骰子+人物移動"
    Private random As New Random()
    Private Const numberOfFrames As Integer = 10 '動畫中的幀數
    Private rolling1 As Boolean = False
    Private rolling2 As Boolean = False
    '骰子的六個面的點數
    Private diceFaces() As Integer = {1, 2, 3, 4, 5, 6}
    Private currentFrame1 As Integer = 0
    Private currentFrame2 As Integer = 2
    '總回合數
    Public Round As Integer = 1
    '目前是誰的回合
    Public Shared Current_player As Integer = 0

    '淘汰玩家人數
    Public Gameover_player As Integer = 0
    '破產玩家
    Dim Gameover_order As Integer
    Private Sub DiceButton_Click(sender As Object, e As EventArgs) Handles Dicebutton.Click '擲骰子
        'Dice Sound
        Static DCount = 10000
        DCount += 1
        With Dclicked
            .Name = "SOUND" & DCount
            .Play(4, False)
        End With

        Current_player += 1
        If Current_player > Total_player Then
            Current_player = 1
            Round += 1
            Round_label.Text = "總回合數：" & " " & Round
        End If

        For i As Integer = 0 To 3
            Select Case Current_player
                Case 1
                    If a.Survival_statestatus = 0 Then
                        Current_player += 1
                    Else
                        Exit For
                    End If
                Case 2
                    If b.Survival_statestatus = 0 Then
                        Current_player += 1
                    Else
                        Exit For
                    End If
                Case 3
                    If c.Survival_statestatus = 0 Then
                        Current_player += 1
                    Else
                        Exit For
                    End If
                Case 4
                    If d.Survival_statestatus = 0 Then
                        Current_player = 1
                    Else
                        Exit For
                    End If
            End Select
        Next
        RollDiceUsingImages()
    End Sub

    Private Sub RollDiceUsingImages()
        If Not rolling1 AndAlso Not rolling2 Then
            rolling1 = True
            rolling2 = True
            Dicebutton.Enabled = False
            point.Text = ""
            DiceTimer1.Start()
            DiceTimer2.Start()
            Dice1.BorderStyle = BorderStyle.Fixed3D
            Dice2.BorderStyle = BorderStyle.Fixed3D
            currentFrame1 = 0
            currentFrame2 = 0
        End If

    End Sub
    Private Sub DiceTimer1_Tick(sender As Object, e As EventArgs) Handles DiceTimer1.Tick
        If rolling1 Then
            Dim randomFace = diceFaces(random.Next(0, 6))
            ' 根據點數設定 Dice1 的圖片
            Select Case randomFace
                Case 1
                    Dice1.Image = My.Resources.ResourceManager.GetObject("1") ' Dice1 是代表點數 1 的圖片
                    Dice1.Tag = "1"
                Case 2
                    Dice1.Image = My.Resources.ResourceManager.GetObject("2") ' Dice2 是代表點數 2 的圖片
                    Dice1.Tag = "2"
                Case 3
                    Dice1.Image = My.Resources.ResourceManager.GetObject("3") ' Dice3 是代表點數 3 的圖片
                    Dice1.Tag = "3"
                Case 4
                    Dice1.Image = My.Resources.ResourceManager.GetObject("4")
                    Dice1.Tag = "4"
                Case 5
                    Dice1.Image = My.Resources.ResourceManager.GetObject("5")
                    Dice1.Tag = "5"
                Case 6
                    Dice1.Image = My.Resources.ResourceManager.GetObject("6")
                    Dice1.Tag = "6"
            End Select

            currentFrame1 += 1
            If currentFrame1 >= numberOfFrames Then
                DiceTimer1.Stop() ' 停止動畫
                rolling1 = False
            End If
        End If
    End Sub

    Private Sub DiceTimer2_Tick(sender As Object, e As EventArgs) Handles DiceTimer2.Tick
        If rolling2 Then
            Dim randomFace = diceFaces(random.Next(0, 6))
            '根據點數設定 Dice2 的圖片
            Select Case randomFace
                Case 1
                    Dice2.Image = My.Resources.ResourceManager.GetObject("1") ' Dice1 是代表點數 1 的圖片
                    Dice2.Tag = "1"
                Case 2
                    Dice2.Image = My.Resources.ResourceManager.GetObject("2") ' Dice2 是代表點數 2 的圖片
                    Dice2.Tag = "2"
                Case 3
                    Dice2.Image = My.Resources.ResourceManager.GetObject("3") ' Dice3 是代表點數 3 的圖片
                    Dice2.Tag = "3"
                Case 4
                    Dice2.Image = My.Resources.ResourceManager.GetObject("4")
                    Dice2.Tag = "4"
                Case 5
                    Dice2.Image = My.Resources.ResourceManager.GetObject("5")
                    Dice2.Tag = "5"
                Case 6
                    Dice2.Image = My.Resources.ResourceManager.GetObject("6")
                    Dice2.Tag = "6"
            End Select

            currentFrame2 += 1
            If currentFrame2 >= numberOfFrames Then
                DiceTimer2.Stop() ' 停止動畫
                point.Text = CType(Dice1.Tag, Integer) + CType(Dice2.Tag, Integer)
                rolling2 = False
                Select Case Current_player
                    Case 1
                        If a.Jail_status <= 0 Then
                            a.Move(point.Text) '角色開始移動
                        ElseIf Dice1.Tag = Dice2.Tag Then ' 擲出相同骰子 
                            a.Jail_status = 0
                            a.Move(point.Text)
                        Else
                            a.Jail_Move()
                        End If
                    Case 2
                        If b.Jail_status <= 0 Then
                            b.Move(point.Text) '角色開始移動
                        ElseIf Dice1.Tag = Dice2.Tag Then ' 擲出相同骰子 
                            b.Jail_status = 0
                            b.Move(point.Text)
                        Else
                            b.Jail_Move()
                        End If

                    Case 3
                        If c.Jail_status <= 0 Then
                            c.Move(point.Text) '角色開始移動
                        ElseIf Dice1.Tag = Dice2.Tag Then ' 擲出相同骰子 
                            c.Jail_status = 0
                            c.Move(point.Text)
                        Else
                            c.Jail_Move()
                        End If

                    Case 4
                        If d.Jail_status <= 0 Then
                            d.Move(point.Text) '角色開始移動
                        ElseIf Dice1.Tag = Dice2.Tag Then ' 擲出相同骰子 
                            d.Jail_status = 0
                            d.Move(point.Text)
                        Else
                            d.Jail_Move()
                        End If
                End Select

            End If
        End If
    End Sub

#End Region

#Region "更新狀態"
    Public Sub Player_Update()
        play_money1.Text = a.Money
        play_money2.Text = b.Money
        play_money3.Text = c.Money
        play_money4.Text = d.Money
        If a.Jail_status < 0 Then
            play_jail_card1.Text = -a.Jail_status
        Else
            play_jail_card1.Text = 0
        End If

        If b.Jail_status < 0 Then
            play_jail_card2.Text = -b.Jail_status
        Else
            play_jail_card2.Text = 0
        End If

        If c.Jail_status < 0 Then
            play_jail_card3.Text = -c.Jail_status
        Else
            play_jail_card3.Text = 0
        End If

        If d.Jail_status < 0 Then
            play_jail_card4.Text = -d.Jail_status
        Else
            play_jail_card4.Text = 0
        End If
        ' b c d 同理
    End Sub
#End Region

#Region "宣告破產"
    Sub Bankrupt()
        If a.Check() Then
            Dim Ans = MsgBox("玩家一破產了，玩家一並不適合投資房地產，或許寫程式才是你的致富之道", , "破產！")
            a.Hide()
            a.Survival_statestatus = 0
            play_land1.Text = "該玩家已破產"
            Gameover_player += 1
            Gameover_order = 1 '破產玩家是誰
            Deletehouse()
            'Failed music
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("losesound"),
                                    AudioPlayMode.Background)
        ElseIf b.Check() Then
            Dim Ans = MsgBox("玩家二破產了，玩家二並不適合投資房地產，或許寫程式才是你的致富之道", , "破產！")
            b.Hide()
            b.Survival_statestatus = 0
            play_land2.Text = "該玩家已破產"
            Gameover_player += 1
            Gameover_order = 2
            Deletehouse()
            'Failed music
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("losesound"),
                                    AudioPlayMode.Background)
        ElseIf c.Check() Then
            Dim Ans = MsgBox("玩家三破產了，玩家三並不適合投資房地產，或許寫程式才是你的致富之道", , "破產！")
            c.Hide()
            c.Survival_statestatus = 0
            play_land3.Text = "該玩家已破產"
            Gameover_player += 1
            Gameover_order = 3
            Deletehouse()
            'Failed music
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("losesound"),
                                    AudioPlayMode.Background)
        ElseIf d.Check() Then
            Dim Ans = MsgBox("玩家四破產了，玩家四並不適合投資房地產，或許寫程式才是你的致富之道", , "破產！")
            d.Hide()
            d.Survival_statestatus = 0
            play_land4.Text = "該玩家已破產"
            Gameover_player += 1
            Gameover_order = 4
            Deletehouse()
            'Failed music
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("losesound"),
                                    AudioPlayMode.Background)
        End If
    End Sub
#End Region

#Region "遊戲結束"
    Sub GameOver()
        Dim winner As String '勝利者
        Dim max As Integer '最高金額
        Dim max_player As Integer '最高金額的玩家
        If Round > EndDayNum Then
            max = a.Money
            max_player = 1
            If b.Money > max Then
                max = b.Money
                max_player = 2
            End If
            If c.Money > max Then
                max = c.Money
                max_player = 3
            End If
            If d.Money > max Then
                max = d.Money
                max_player = 4
            End If

            Select Case max_player
                Case 1
                    winner = a.Name
                Case 2
                    winner = b.Name
                Case 3
                    winner = c.Name
                Case 4
                    winner = d.Name
            End Select
        ElseIf Gameover_player = Total_player - 1 Then

            If a.Survival_statestatus = 1 Then
                winner = a.Name
            End If
            If b.Survival_statestatus = 1 Then
                winner = b.Name
            End If
            If c.Survival_statestatus = 1 Then
                winner = c.Name
            End If
            If d.Survival_statestatus = 1 Then
                winner = d.Name
            End If
        Else
            If a.Money >= EndCoinNum Then
                winner = a.Name
            End If
            If b.Money >= EndCoinNum Then
                winner = b.Name
            End If
            If c.Money >= EndCoinNum Then
                winner = c.Name
            End If
            If d.Money >= EndCoinNum Then
                winner = d.Name
            End If
        End If

        MsgBox("遊戲結束，勝利者為" & winner, , "GameOver")
        Dim Ans = MsgBox("是否重啟新局", MsgBoxStyle.YesNo, "重起新局")
        If Ans = MsgBoxResult.Yes Then
            Hidegame()
            Clear_information()
            Me.BackgroundImage = My.Resources.ResourceManager.GetObject("beginScience")
            showMain()
            'Loading Background music
            My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("Wallpaper"),
                               AudioPlayMode.BackgroundLoop)
            Player_Update()
        Else
            Me.Close()
        End If
    End Sub

#End Region

#Region "重建玩家資訊"
    Sub Clear_information()
        Current_player = 0
        Dice1.Image = Nothing
        Dice2.Image = Nothing
        point.Text = ""
        Gameover_player = 0
        a.Survival_statestatus = 1 : b.Survival_statestatus = 1 : c.Survival_statestatus = 1 : d.Survival_statestatus = 1
        play_land1.Text = "玩家一持有土地:" : play_land2.Text = "玩家二持有土地:" : play_land3.Text = "玩家三持有土地:" : play_land4.Text = "玩家四持有土地:"
        a.Jail_status = 0 : b.Jail_status = 0 : c.Jail_status = 0 : d.Jail_status = 0
        a.Money = 0 : b.Money = 0 : c.Money = 0 : d.Money = 0
        a.Position = 0 : b.Position = 0 : c.Position = 0 : d.Position = 0
        a.Location = New Point(coordinates(0, 0), coordinates(0, 1)) : b.Location = New Point(coordinates(0, 0), coordinates(0, 1)) : c.Location = New Point(coordinates(0, 0), coordinates(0, 1)) : d.Location = New Point(coordinates(0, 0), coordinates(0, 1))
        Kinmen.Initialize()
        Mazu.Initialize()
        Liuchiu.Initialize()
        Koto_island.Initialize()
        Lyudao.Initialize()
        Penghu.Initialize()
        Taitung.Initialize()
        Hualien.Initialize()
        Pingtung.Initialize()
        Kaohsiung.Initialize()
        Tainan.Initialize()
        Chiayi.Initialize()
        Nantou.Initialize()
        Yunlin.Initialize()
        Changhua.Initialize()
        Taichung.Initialize()
        Miaoli.Initialize()
        Yilan.Initialize()
        Hsinchu.Initialize()
        Taoyuan.Initialize()
        New_Taipei.Initialize()
        Taipei.Initialize()
        Keelung.Initialize()
    End Sub
#End Region

#Region "清除房地產"
    Public Sub Deletehouse()
        If (Kinmen.Owner = Gameover_order) Then
            Kinmen.Owner = 0
            Kinmen.Level = -1
        End If
        If (Mazu.Owner = Gameover_order) Then
            Mazu.Owner = 0
            Mazu.Level = -1
        End If
        If (Liuchiu.Owner = Gameover_order) Then
            Liuchiu.Owner = 0
            Liuchiu.Level = -1
        End If
        If (Koto_island.Owner = Gameover_order) Then
            Koto_island.Owner = 0
            Koto_island.Level = -1
        End If
        If (Lyudao.Owner = Gameover_order) Then
            Lyudao.Owner = 0
            Lyudao.Level = -1
        End If
        If (Penghu.Owner = Gameover_order) Then
            Penghu.Owner = 0
            Penghu.Level = -1
        End If
        If (Taitung.Owner = Gameover_order) Then
            Taitung.Owner = 0
            Taitung.Level = -1
        End If
        If (Hualien.Owner = Gameover_order) Then
            Hualien.Owner = 0
            Hualien.Level = -1
        End If
        If (Pingtung.Owner = Gameover_order) Then
            Pingtung.Owner = 0
            Pingtung.Level = -1
        End If
        If (Kaohsiung.Owner = Gameover_order) Then
            Kaohsiung.Owner = 0
            Kaohsiung.Level = -1
        End If
        If (Tainan.Owner = Gameover_order) Then
            Tainan.Owner = 0
            Tainan.Level = -1
        End If
        If (Chiayi.Owner = Gameover_order) Then
            Chiayi.Owner = 0
            Chiayi.Level = -1
        End If
        If (Nantou.Owner = Gameover_order) Then
            Nantou.Owner = 0
            Nantou.Level = -1
        End If
        If (Yunlin.Owner = Gameover_order) Then
            Yunlin.Owner = 0
            Yunlin.Level = -1
        End If
        If (Changhua.Owner = Gameover_order) Then
            Changhua.Owner = 0
            Changhua.Level = -1
        End If
        If (Taichung.Owner = Gameover_order) Then
            Taichung.Owner = 0
            Taichung.Level = -1
        End If
        If (Miaoli.Owner = Gameover_order) Then
            Miaoli.Owner = 0
            Miaoli.Level = -1
        End If
        If (Yilan.Owner = Gameover_order) Then
            Yilan.Owner = 0
            Yilan.Level = -1
        End If
        If (Hsinchu.Owner = Gameover_order) Then
            Hsinchu.Owner = 0
            Hsinchu.Level = -1
        End If
        If (Taoyuan.Owner = Gameover_order) Then
            Taoyuan.Owner = 0
            Taoyuan.Level = -1
        End If
        If (New_Taipei.Owner = Gameover_order) Then
            New_Taipei.Owner = 0
            New_Taipei.Level = -1
        End If
        If (Taipei.Owner = Gameover_order) Then
            Taipei.Owner = 0
            Taipei.Level = -1
        End If
        If (Keelung.Owner = Gameover_order) Then
            Keelung.Owner = 0
            Keelung.Level = -1
        End If
    End Sub
#End Region

End Class

#Region "建立玩家類別"
Public Class Player
    Inherits PictureBox
    Public Property Position As Integer
    Public Property Money As Integer
    Public Property Jail_status As Integer = 0 '是否在監獄 負數代表擁有監獄豁免權
    Public Property Survival_statestatus As Integer = 0
    Private moveTimer As Timer
    Public Sub New(ByVal playerName As String, ByVal startingPosition As Integer, ByVal startingMoney As Integer)
        '初始化player物件
        Me.Name = playerName
        Me.Position = startingPosition
        Me.Money = startingMoney
        Me.Size = New Size(41, 72)
        '初始化timer物件
        moveTimer = New Timer()
        moveTimer.Interval = 25 ' 設置計時器間隔（毫秒）
        AddHandler moveTimer.Tick, AddressOf MoveTimer_Tick
    End Sub

    Dim direction As Integer = 1

#Region "一般Move方法"
    Public Shadows Sub Move(ByVal point As Integer)
        Position += point
        If Position >= 32 Then
            Position -= 32
        End If
        Me.BringToFront()
        moveTimer.Start()
    End Sub

    Public Sub MoveTimer_Tick(sender As Object, e As EventArgs)
        '在計時器觸發時執行的操作
        If Me.Left = Form1.coordinates(0, 0) And Me.Top = Form1.coordinates(0, 1) Then
            direction = 1
            Money += 2000
            Form1.Player_Update()
        ElseIf Me.Left = Form1.coordinates(8, 0) And Me.Top <= Form1.coordinates(8, 1) Then
            direction = 2
        ElseIf Me.Left = Form1.coordinates(16, 0) And Me.Top = Form1.coordinates(16, 1) Then
            direction = 3
        ElseIf Me.Left = Form1.coordinates(24, 0) And Me.Top = Form1.coordinates(24, 1) Then
            direction = 4
        End If

        Select Case direction
            Case 1
                If Me.Left < Form1.coordinates(Position, 0) Or Me.Top <> Form1.coordinates(Position, 1) Then
                    ' 如果還沒到達目標位置，持續增加 Me.Left 的值
                    Form1.Dicebutton.Enabled = False
                    Me.Left += Form1.ClientSize.Width * 0.005
                    If Me.Left > Form1.coordinates(8, 0) Then
                        Me.Left = Form1.coordinates(8, 0)
                    End If
                Else
                    '如果超出範圍Left-1修正
                    If Me.Left > Form1.coordinates(Position, 0) Then
                        Me.Left -= 1
                    Else
                        ' 如果到達目標位置，停止計時器 並開始做出對應的動作
                        Form1.Dicebutton.Enabled = True
                        moveTimer.Stop()
                        Action(Position)
                    End If
                End If

            Case 2
                If Me.Top < Form1.coordinates(Position, 1) Or Me.Left <> Form1.coordinates(Position, 0) Then
                    ' 如果還沒到達目標位置，持續增加 Me.Top 的值 
                    Form1.Dicebutton.Enabled = False
                    Me.Top += Form1.ClientSize.Height * 0.01
                    If Me.Top > Form1.coordinates(16, 1) Then
                        Me.Top = Form1.coordinates(16, 1)
                    End If
                Else
                    If Me.Top > Form1.coordinates(Position, 1) Then
                        Me.Top -= 1
                    Else
                        ' 如果到達目標位置，停止計時器 並開始做出對應的動作
                        Form1.Dicebutton.Enabled = True
                        moveTimer.Stop()
                        Action(Position)
                    End If
                End If
            Case 3
                If Me.Left > Form1.coordinates(Position, 0) Or Me.Top <> Form1.coordinates(Position, 1) Then
                    ' 如果還沒到達目標位置，持續減少 Me.Left 的值
                    Form1.Dicebutton.Enabled = False
                    Me.Left -= Form1.ClientSize.Width * 0.005
                    If Me.Left < Form1.coordinates(24, 0) Then
                        Me.Left = Form1.coordinates(24, 0)
                    End If
                Else
                    If Me.Left < Form1.coordinates(Position, 0) Then
                        Me.Left += 1
                    Else
                        ' 如果到達目標位置，停止計時器 並開始做出對應的動作
                        Form1.Dicebutton.Enabled = True
                        moveTimer.Stop()
                        Action(Position)
                    End If
                End If
            Case 4
                If Me.Top > Form1.coordinates(Position, 1) Or Me.Left <> Form1.coordinates(Position, 0) Then
                    ' 如果還沒到達目標位置，持續減少 Me.Top 的值
                    Form1.Dicebutton.Enabled = False
                    Me.Top -= Form1.ClientSize.Height * 0.01
                    If Me.Top < Form1.coordinates(0, 1) Then
                        Me.Top = Form1.coordinates(0, 1)
                    End If
                Else
                    If Me.Top > Form1.coordinates(Position, 1) Then
                        Me.Top += 1
                    Else
                        ' 如果到達目標位置，停止計時器 並開始做出對應的動作
                        Form1.Dicebutton.Enabled = True
                        moveTimer.Stop()
                        Action(Position)
                    End If
                End If

        End Select
    End Sub
#End Region

#Region "監獄Move方法"
    Sub Jail_Move()
        Jail_status -= 1
        MsgBox("未擲出相同點數的骰子，還剩下" & Jail_status & "回合即可出獄")
        Form1.Dicebutton.Enabled = True
    End Sub
#End Region

#Region "Action方法"
    Public Sub Action(ByVal position As Integer) '參數為目的地的位置
        'Money Temp
        Dim tempMoney = Me.Money

        Select Case position
            Case 0 '不用執行動作(在Move時已經處理完畢)
            Case 1 '金門
                '判斷土地誰持有
                '尚未被任何人持有
                If Kinmen.Owner = 0 Then
                    Dim reply = MsgBox("你來到了金門，此空地價格為" & Kinmen.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Kinmen.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Kinmen.Buy
                            Kinmen.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "金門"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "金門"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "金門"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "金門"
                            End Select
                            Kinmen.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Kinmen.Owner = Form1.Current_player Then
                    Select Case Kinmen.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Kinmen.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Kinmen.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Kinmen.Build
                                    Kinmen.Level = 1
                                    'Form1.Land_Level(0) = 1
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Kinmen.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Kinmen.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Kinmen.Build
                                    Kinmen.Level = 2
                                    'Form1.Land_Level(0) = 2
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Kinmen.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Kinmen.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Kinmen.SpecialBuild
                                    Kinmen.Level = 3
                                    'Form1.Land_Level(0) = 3
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Kinmen.Owner & "所持有，需付過路費" & Kinmen.fee(Kinmen.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Kinmen.fee(Kinmen.Level)
                        Select Case Kinmen.Owner
                            Case 1
                                Form1.a.Money += Kinmen.fee(Kinmen.Level)
                            Case 2
                                Form1.b.Money += Kinmen.fee(Kinmen.Level)
                            Case 3
                                Form1.c.Money += Kinmen.fee(Kinmen.Level)
                            Case 4
                                Form1.d.Money += Kinmen.fee(Kinmen.Level)
                        End Select
                    End If
                End If
            Case 2 '馬祖
                If Mazu.Owner = 0 Then
                    Dim reply = MsgBox("你來到了馬祖，此空地價格為" & Mazu.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Mazu.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Mazu.Buy
                            Mazu.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "馬祖"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "馬祖"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "馬祖"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "馬祖"
                            End Select
                            Mazu.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Mazu.Owner = Form1.Current_player Then
                    Select Case Mazu.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Mazu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Mazu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Mazu.Build
                                    Mazu.Level = 1
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Mazu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Mazu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Mazu.Build
                                    Mazu.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Mazu.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Mazu.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Mazu.SpecialBuild
                                    Mazu.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Mazu.Owner & "所持有，需付過路費" & Mazu.fee(Mazu.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Mazu.fee(Mazu.Level)
                        Select Case Mazu.Owner
                            Case 1
                                Form1.a.Money += Mazu.fee(Mazu.Level)
                            Case 2
                                Form1.b.Money += Mazu.fee(Mazu.Level)
                            Case 3
                                Form1.c.Money += Mazu.fee(Mazu.Level)
                            Case 4
                                Form1.d.Money += Mazu.fee(Mazu.Level)
                        End Select
                    End If
                End If
            Case 3 '所得稅
                MsgBox("繳交所得稅2000元", , "所得稅")
                Me.Money -= 2000
            Case 4 '小琉球
                If Liuchiu.Owner = 0 Then
                    Dim reply = MsgBox("你來到了小琉球，此空地價格為" & Liuchiu.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Liuchiu.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Liuchiu.Buy
                            Liuchiu.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "小琉球"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "小琉球"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "小琉球"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "小琉球"
                            End Select
                            Liuchiu.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Liuchiu.Owner = Form1.Current_player Then
                    Select Case Liuchiu.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Liuchiu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Liuchiu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Liuchiu.Build
                                    Liuchiu.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Liuchiu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Liuchiu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Liuchiu.Build
                                    Liuchiu.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Liuchiu.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Liuchiu.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Liuchiu.SpecialBuild
                                    Liuchiu.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Liuchiu.Owner & "所持有，需付過路費" & Liuchiu.fee(Liuchiu.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Liuchiu.fee(Liuchiu.Level)
                        Select Case Liuchiu.Owner
                            Case 1
                                Form1.a.Money += Liuchiu.fee(Liuchiu.Level)
                            Case 2
                                Form1.b.Money += Liuchiu.fee(Liuchiu.Level)
                            Case 3
                                Form1.c.Money += Liuchiu.fee(Liuchiu.Level)
                            Case 4
                                Form1.d.Money += Liuchiu.fee(Liuchiu.Level)
                        End Select
                    End If
                End If
            Case 5 '事件
                Destiny(position)
            Case 6 '蘭嶼
                If Koto_island.Owner = 0 Then
                    Dim reply = MsgBox("你來到了蘭嶼，此空地價格為" & Koto_island.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Koto_island.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Koto_island.Buy
                            Koto_island.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "蘭嶼"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "蘭嶼"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "蘭嶼"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "蘭嶼"
                            End Select
                            Koto_island.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Koto_island.Owner = Form1.Current_player Then
                    Select Case Koto_island.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Koto_island.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Koto_island.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Koto_island.Build
                                    Koto_island.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Koto_island.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Koto_island.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Koto_island.Build
                                    Koto_island.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Koto_island.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Koto_island.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Koto_island.SpecialBuild
                                    Koto_island.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Koto_island.Owner & "所持有，需付過路費" & Koto_island.fee(Koto_island.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Koto_island.fee(Koto_island.Level)
                        Select Case Koto_island.Owner
                            Case 1
                                Form1.a.Money += Koto_island.fee(Koto_island.Level)
                            Case 2
                                Form1.b.Money += Koto_island.fee(Koto_island.Level)
                            Case 3
                                Form1.c.Money += Koto_island.fee(Koto_island.Level)
                            Case 4
                                Form1.d.Money += Koto_island.fee(Koto_island.Level)
                        End Select
                    End If
                End If
            Case 7 '綠島
                If Lyudao.Owner = 0 Then
                    Dim reply = MsgBox("你來到了綠島，此空地價格為" & Lyudao.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Lyudao.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Lyudao.Buy
                            Lyudao.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "綠島"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "綠島"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "綠島"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "綠島"
                            End Select
                            Lyudao.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Lyudao.Owner = Form1.Current_player Then
                    Select Case Lyudao.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Lyudao.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Lyudao.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Lyudao.Build
                                    Lyudao.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Lyudao.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Lyudao.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Lyudao.Build
                                    Lyudao.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Lyudao.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Lyudao.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Lyudao.SpecialBuild
                                    Lyudao.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Lyudao.Owner & "所持有，需付過路費" & Lyudao.fee(Lyudao.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Lyudao.fee(Lyudao.Level)
                        Select Case Lyudao.Owner
                            Case 1
                                Form1.a.Money += Lyudao.fee(Lyudao.Level)
                            Case 2
                                Form1.b.Money += Lyudao.fee(Lyudao.Level)
                            Case 3
                                Form1.c.Money += Lyudao.fee(Lyudao.Level)
                            Case 4
                                Form1.d.Money += Lyudao.fee(Lyudao.Level)
                        End Select
                    End If
                End If
            Case 8 '不用執行動作
            Case 9 '澎湖
                If Penghu.Owner = 0 Then
                    Dim reply = MsgBox("你來到了澎湖，此空地價格為" & Penghu.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Penghu.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Penghu.Buy
                            Penghu.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "澎湖"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "澎湖"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "澎湖"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "澎湖"
                            End Select
                            Penghu.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Penghu.Owner = Form1.Current_player Then
                    Select Case Penghu.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Penghu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Penghu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Penghu.Build
                                    Penghu.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Penghu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Penghu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Penghu.Build
                                    Penghu.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Penghu.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Penghu.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Penghu.SpecialBuild
                                    Penghu.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Penghu.Owner & "所持有，需付過路費" & Penghu.fee(Penghu.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Penghu.fee(Penghu.Level)
                        Select Case Penghu.Owner
                            Case 1
                                Form1.a.Money += Penghu.fee(Penghu.Level)
                            Case 2
                                Form1.b.Money += Penghu.fee(Penghu.Level)
                            Case 3
                                Form1.c.Money += Penghu.fee(Penghu.Level)
                            Case 4
                                Form1.d.Money += Penghu.fee(Penghu.Level)
                        End Select
                    End If
                End If
            Case 10 '台東
                If Taitung.Owner = 0 Then
                    Dim reply = MsgBox("你來到了台東，此空地價格為" & Taitung.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Taitung.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Taitung.Buy
                            Taitung.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "台東"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "台東"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "台東"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "台東"
                            End Select
                            Taitung.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Taitung.Owner = Form1.Current_player Then
                    Select Case Taitung.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Taitung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taitung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taitung.Build
                                    Taitung.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Taitung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taitung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taitung.Build
                                    Taitung.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Taitung.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taitung.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taitung.SpecialBuild
                                    Taitung.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Taitung.Owner & "所持有，需付過路費" & Taitung.fee(Taitung.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Taitung.fee(Taitung.Level)
                        Select Case Taitung.Owner
                            Case 1
                                Form1.a.Money += Taitung.fee(Taitung.Level)
                            Case 2
                                Form1.b.Money += Taitung.fee(Taitung.Level)
                            Case 3
                                Form1.c.Money += Taitung.fee(Taitung.Level)
                            Case 4
                                Form1.d.Money += Taitung.fee(Taitung.Level)
                        End Select
                    End If
                End If
            Case 11 '花蓮
                If Hualien.Owner = 0 Then
                    Dim reply = MsgBox("你來到了花蓮，此空地價格為" & Hualien.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Hualien.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Hualien.Buy
                            Hualien.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "花蓮"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "花蓮"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "花蓮"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "花蓮"
                            End Select
                            Hualien.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Hualien.Owner = Form1.Current_player Then
                    Select Case Hualien.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Hualien.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Hualien.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Hualien.Build
                                    Hualien.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Hualien.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Hualien.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Hualien.Build
                                    Hualien.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Hualien.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Hualien.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Hualien.SpecialBuild
                                    Hualien.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Hualien.Owner & "所持有，需付過路費" & Hualien.fee(Hualien.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Hualien.fee(Hualien.Level)
                        Select Case Hualien.Owner
                            Case 1
                                Form1.a.Money += Hualien.fee(Hualien.Level)
                            Case 2
                                Form1.b.Money += Hualien.fee(Hualien.Level)
                            Case 3
                                Form1.c.Money += Hualien.fee(Hualien.Level)
                            Case 4
                                Form1.d.Money += Hualien.fee(Hualien.Level)
                        End Select
                    End If
                End If
            Case 12 '屏東
                If Pingtung.Owner = 0 Then
                    Dim reply = MsgBox("你來到了屏東，此空地價格為" & Pingtung.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Pingtung.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Pingtung.Buy
                            Pingtung.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "屏東"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "屏東"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "屏東"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "屏東"
                            End Select
                            Pingtung.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Pingtung.Owner = Form1.Current_player Then
                    Select Case Pingtung.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Pingtung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Pingtung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Pingtung.Build
                                    Pingtung.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Pingtung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Pingtung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Pingtung.Build
                                    Pingtung.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Pingtung.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Pingtung.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Pingtung.SpecialBuild
                                    Pingtung.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Pingtung.Owner & "所持有，需付過路費" & Pingtung.fee(Pingtung.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Pingtung.fee(Pingtung.Level)
                        Select Case Pingtung.Owner
                            Case 1
                                Form1.a.Money += Pingtung.fee(Pingtung.Level)
                            Case 2
                                Form1.b.Money += Pingtung.fee(Pingtung.Level)
                            Case 3
                                Form1.c.Money += Pingtung.fee(Pingtung.Level)
                            Case 4
                                Form1.d.Money += Pingtung.fee(Pingtung.Level)
                        End Select
                    End If
                End If
            Case 13 '事件
                Destiny(position)
            Case 14 '高雄
                If Kaohsiung.Owner = 0 Then
                    Dim reply = MsgBox("你來到了高雄，此空地價格為" & Kaohsiung.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Kaohsiung.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Kaohsiung.Buy
                            Kaohsiung.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "高雄"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "高雄"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "高雄"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "高雄"
                            End Select
                            Kaohsiung.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Kaohsiung.Owner = Form1.Current_player Then
                    Select Case Kaohsiung.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Kaohsiung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Kaohsiung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Kaohsiung.Build
                                    Kaohsiung.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Kaohsiung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Kaohsiung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Kaohsiung.Build
                                    Kaohsiung.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Kaohsiung.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Kaohsiung.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Kaohsiung.SpecialBuild
                                    Kaohsiung.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Kaohsiung.Owner & "所持有，需付過路費" & Kaohsiung.fee(Kaohsiung.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Kaohsiung.fee(Kaohsiung.Level)
                        Select Case Kaohsiung.Owner
                            Case 1
                                Form1.a.Money += Kaohsiung.fee(Kaohsiung.Level)
                            Case 2
                                Form1.b.Money += Kaohsiung.fee(Kaohsiung.Level)
                            Case 3
                                Form1.c.Money += Kaohsiung.fee(Kaohsiung.Level)
                            Case 4
                                Form1.d.Money += Kaohsiung.fee(Kaohsiung.Level)
                        End Select
                    End If
                End If
            Case 15 '台南
                If Tainan.Owner = 0 Then
                    Dim reply = MsgBox("你來到了台南，此空地價格為" & Tainan.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Tainan.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Tainan.Buy
                            Tainan.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "台南"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "台南"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "台南"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "台南"
                            End Select
                            Tainan.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Tainan.Owner = Form1.Current_player Then
                    Select Case Tainan.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Tainan.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Tainan.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Tainan.Build
                                    Tainan.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Tainan.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Tainan.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Tainan.Build
                                    Tainan.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Tainan.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Tainan.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Tainan.SpecialBuild
                                    Tainan.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Tainan.Owner & "所持有，需付過路費" & Tainan.fee(Tainan.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Tainan.fee(Tainan.Level)
                        Select Case Tainan.Owner
                            Case 1
                                Form1.a.Money += Tainan.fee(Tainan.Level)
                            Case 2
                                Form1.b.Money += Tainan.fee(Tainan.Level)
                            Case 3
                                Form1.c.Money += Tainan.fee(Tainan.Level)
                            Case 4
                                Form1.d.Money += Tainan.fee(Tainan.Level)
                        End Select
                    End If
                End If
            Case 16 '小遊戲
                Form1.Visible = False
                Dim SMGame As New Small_Game
                SMGame.ShowDialog() '暫停主程式，執行小游戲
                Money += ShareMoney.SMMoney
                Form1.Visible = True
                Form1.Size = New Size(1920, 1080)
                If ShareMoney.SMMoney <> 0 Then
                    MsgBox(Me.Name & "獲得" & ShareMoney.SMMoney & "!",, "Good result")
                Else
                    MsgBox("再加油！",, "Bad result")
                End If

            Case 17 '嘉義
                If Chiayi.Owner = 0 Then
                    Dim reply = MsgBox("你來到了嘉義，此空地價格為" & Chiayi.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Chiayi.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Chiayi.Buy
                            Chiayi.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "嘉義"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "嘉義"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "嘉義"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "嘉義"
                            End Select
                            Chiayi.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Chiayi.Owner = Form1.Current_player Then
                    Select Case Chiayi.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Chiayi.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Chiayi.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Chiayi.Build
                                    Chiayi.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Chiayi.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Chiayi.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Chiayi.Build
                                    Chiayi.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Chiayi.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Chiayi.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Chiayi.SpecialBuild
                                    Chiayi.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Chiayi.Owner & "所持有，需付過路費" & Chiayi.fee(Chiayi.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Chiayi.fee(Chiayi.Level)
                        Select Case Chiayi.Owner
                            Case 1
                                Form1.a.Money += Chiayi.fee(Chiayi.Level)
                            Case 2
                                Form1.b.Money += Chiayi.fee(Chiayi.Level)
                            Case 3
                                Form1.c.Money += Chiayi.fee(Chiayi.Level)
                            Case 4
                                Form1.d.Money += Chiayi.fee(Chiayi.Level)
                        End Select
                    End If
                End If
            Case 18 '南投
                If Nantou.Owner = 0 Then
                    Dim reply = MsgBox("你來到了南投，此空地價格為" & Nantou.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Nantou.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Nantou.Buy
                            Nantou.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "南投"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "南投"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "南投"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "南投"
                            End Select
                            Nantou.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Nantou.Owner = Form1.Current_player Then
                    Select Case Nantou.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Nantou.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Nantou.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Nantou.Build
                                    Nantou.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Nantou.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Nantou.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Nantou.Build
                                    Nantou.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Nantou.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Nantou.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Nantou.SpecialBuild
                                    Nantou.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Nantou.Owner & "所持有，需付過路費" & Nantou.fee(Nantou.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Nantou.fee(Nantou.Level)
                        Select Case Nantou.Owner
                            Case 1
                                Form1.a.Money += Nantou.fee(Nantou.Level)
                            Case 2
                                Form1.b.Money += Nantou.fee(Nantou.Level)
                            Case 3
                                Form1.c.Money += Nantou.fee(Nantou.Level)
                            Case 4
                                Form1.d.Money += Nantou.fee(Nantou.Level)
                        End Select
                    End If
                End If
            Case 19 '雲林
                If Yunlin.Owner = 0 Then
                    Dim reply = MsgBox("你來到了雲林，此空地價格為" & Yunlin.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Yunlin.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Yunlin.Buy
                            Yunlin.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "雲林"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "雲林"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "雲林"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "雲林"
                            End Select
                            Yunlin.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Yunlin.Owner = Form1.Current_player Then
                    Select Case Yunlin.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Yunlin.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Yunlin.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Yunlin.Build
                                    Yunlin.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Yunlin.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Yunlin.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Yunlin.Build
                                    Yunlin.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Yunlin.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Yunlin.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Yunlin.SpecialBuild
                                    Yunlin.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Yunlin.Owner & "所持有，需付過路費" & Yunlin.fee(Yunlin.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Yunlin.fee(Yunlin.Level)
                        Select Case Yunlin.Owner
                            Case 1
                                Form1.a.Money += Yunlin.fee(Yunlin.Level)
                            Case 2
                                Form1.b.Money += Yunlin.fee(Yunlin.Level)
                            Case 3
                                Form1.c.Money += Yunlin.fee(Yunlin.Level)
                            Case 4
                                Form1.d.Money += Yunlin.fee(Yunlin.Level)
                        End Select
                    End If
                End If
            Case 20 '彰化
                If Changhua.Owner = 0 Then
                    Dim reply = MsgBox("你來到了彰化，此空地價格為" & Changhua.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Changhua.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Changhua.Buy
                            Changhua.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "彰化"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "彰化"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "彰化"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "彰化"
                            End Select
                            Changhua.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Changhua.Owner = Form1.Current_player Then
                    Select Case Changhua.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Changhua.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Changhua.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.Build
                                    Changhua.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Changhua.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Changhua.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.Build
                                    Changhua.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Changhua.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Changhua.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Changhua.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Changhua.Owner & "所持有，需付過路費" & Changhua.fee(Changhua.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Changhua.fee(Changhua.Level)
                        Select Case Changhua.Owner
                            Case 1
                                Form1.a.Money += Changhua.fee(Changhua.Level)
                            Case 2
                                Form1.b.Money += Changhua.fee(Changhua.Level)
                            Case 3
                                Form1.c.Money += Changhua.fee(Changhua.Level)
                            Case 4
                                Form1.d.Money += Changhua.fee(Changhua.Level)
                        End Select
                    End If
                End If
            Case 21 '事件
                Destiny(position)
            Case 22 '台中
                If Taichung.Owner = 0 Then
                    Dim reply = MsgBox("你來到了台中，此空地價格為" & Taichung.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Taichung.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Taichung.Buy
                            Taichung.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "台中"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "台中"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "台中"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "台中"
                            End Select
                            Taichung.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Taichung.Owner = Form1.Current_player Then
                    Select Case Taichung.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Taichung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taichung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taichung.Build
                                    Taichung.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Taichung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taichung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taichung.Build
                                    Taichung.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Taichung.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taichung.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Taichung.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Taichung.Owner & "所持有，需付過路費" & Taichung.fee(Taichung.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Taichung.fee(Taichung.Level)
                        Select Case Taichung.Owner
                            Case 1
                                Form1.a.Money += Taichung.fee(Taichung.Level)
                            Case 2
                                Form1.b.Money += Taichung.fee(Taichung.Level)
                            Case 3
                                Form1.c.Money += Taichung.fee(Taichung.Level)
                            Case 4
                                Form1.d.Money += Taichung.fee(Taichung.Level)
                        End Select
                    End If
                End If
            Case 23 '苗栗
                If Miaoli.Owner = 0 Then
                    Dim reply = MsgBox("你來到了苗栗，此空地價格為" & Miaoli.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Miaoli.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Miaoli.Buy
                            Miaoli.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "苗栗"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "苗栗"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "苗栗"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "苗栗"
                            End Select
                            Miaoli.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Miaoli.Owner = Form1.Current_player Then
                    Select Case Miaoli.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Miaoli.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Miaoli.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Miaoli.Build
                                    Miaoli.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Miaoli.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Miaoli.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Miaoli.Build
                                    Miaoli.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Miaoli.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Miaoli.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Miaoli.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Miaoli.Owner & "所持有，需付過路費" & Miaoli.fee(Miaoli.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Miaoli.fee(Miaoli.Level)
                        Select Case Miaoli.Owner
                            Case 1
                                Form1.a.Money += Miaoli.fee(Miaoli.Level)
                            Case 2
                                Form1.b.Money += Miaoli.fee(Miaoli.Level)
                            Case 3
                                Form1.c.Money += Miaoli.fee(Miaoli.Level)
                            Case 4
                                Form1.d.Money += Miaoli.fee(Miaoli.Level)
                        End Select
                    End If
                End If
            Case 24 '進監獄
                If Me.Jail_status = 0 Then
                    MsgBox("你來到了監獄，需擲出2顆相同的骰子或原地擲骰三回合即可出獄", , "監獄")
                    Me.Jail_status = 3
                    Me.Position = 8
                    Me.direction = 2 : Me.Left = Form1.coordinates(8, 0) : Me.Top = Form1.coordinates(8, 1)
                Else
                    MsgBox("你擁有監獄豁免權, 因此不需要入獄", , "免使金牌")
                    Me.Jail_status += 1
                End If

            Case 25 '宜蘭
                If Yilan.Owner = 0 Then
                    Dim reply = MsgBox("你來到了宜蘭，此空地價格為" & Yilan.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Yilan.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Yilan.Buy
                            Yilan.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "宜蘭"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "宜蘭"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "宜蘭"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "宜蘭"
                            End Select
                            Yilan.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Yilan.Owner = Form1.Current_player Then
                    Select Case Yilan.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Yilan.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Yilan.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Yilan.Build
                                    Yilan.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Yilan.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Yilan.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Yilan.Build
                                    Yilan.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Yilan.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Yilan.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Yilan.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Yilan.Owner & "所持有，需付過路費" & Yilan.fee(Yilan.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Yilan.fee(Yilan.Level)
                        Select Case Yilan.Owner
                            Case 1
                                Form1.a.Money += Yilan.fee(Yilan.Level)
                            Case 2
                                Form1.b.Money += Yilan.fee(Yilan.Level)
                            Case 3
                                Form1.c.Money += Yilan.fee(Yilan.Level)
                            Case 4
                                Form1.d.Money += Yilan.fee(Yilan.Level)
                        End Select
                    End If
                End If
            Case 26 '新竹
                If Hsinchu.Owner = 0 Then
                    Dim reply = MsgBox("你來到了新竹，此空地價格為" & Hsinchu.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Hsinchu.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Hsinchu.Buy
                            Hsinchu.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "新竹"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "新竹"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "新竹"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "新竹"
                            End Select
                            Hsinchu.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Hsinchu.Owner = Form1.Current_player Then
                    Select Case Hsinchu.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Hsinchu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Hsinchu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Hsinchu.Build
                                    Hsinchu.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Hsinchu.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Hsinchu.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Hsinchu.Build
                                    Hsinchu.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Hsinchu.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Hsinchu.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Hsinchu.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Hsinchu.Owner & "所持有，需付過路費" & Hsinchu.fee(Hsinchu.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Hsinchu.fee(Hsinchu.Level)
                        Select Case Hsinchu.Owner
                            Case 1
                                Form1.a.Money += Hsinchu.fee(Hsinchu.Level)
                            Case 2
                                Form1.b.Money += Hsinchu.fee(Hsinchu.Level)
                            Case 3
                                Form1.c.Money += Hsinchu.fee(Hsinchu.Level)
                            Case 4
                                Form1.d.Money += Hsinchu.fee(Hsinchu.Level)
                        End Select
                    End If
                End If
            Case 27 '桃園
                If Taoyuan.Owner = 0 Then
                    Dim reply = MsgBox("你來到了桃園，此空地價格為" & Taoyuan.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Taoyuan.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Taoyuan.Buy
                            Taoyuan.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "桃園"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "桃園"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "桃園"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "桃園"
                            End Select
                            Taoyuan.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Taoyuan.Owner = Form1.Current_player Then
                    Select Case Taoyuan.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Taoyuan.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taoyuan.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taoyuan.Build
                                    Taoyuan.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Taoyuan.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taoyuan.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taoyuan.Build
                                    Taoyuan.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Taoyuan.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taoyuan.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Taoyuan.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Taoyuan.Owner & "所持有，需付過路費" & Taoyuan.fee(Taoyuan.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Taoyuan.fee(Taoyuan.Level)
                        Select Case Taoyuan.Owner
                            Case 1
                                Form1.a.Money += Taoyuan.fee(Taoyuan.Level)
                            Case 2
                                Form1.b.Money += Taoyuan.fee(Taoyuan.Level)
                            Case 3
                                Form1.c.Money += Taoyuan.fee(Taoyuan.Level)
                            Case 4
                                Form1.d.Money += Taoyuan.fee(Taoyuan.Level)
                        End Select
                    End If
                End If
            Case 28 '新北
                If New_Taipei.Owner = 0 Then
                    Dim reply = MsgBox("你來到了新北，此空地價格為" & New_Taipei.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < New_Taipei.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= New_Taipei.Buy
                            New_Taipei.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "新北"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "新北"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "新北"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "新北"
                            End Select
                            New_Taipei.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf New_Taipei.Owner = Form1.Current_player Then
                    Select Case New_Taipei.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & New_Taipei.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < New_Taipei.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= New_Taipei.Build
                                    New_Taipei.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & New_Taipei.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < New_Taipei.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= New_Taipei.Build
                                    New_Taipei.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & New_Taipei.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < New_Taipei.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    New_Taipei.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & New_Taipei.Owner & "所持有，需付過路費" & New_Taipei.fee(New_Taipei.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= New_Taipei.fee(New_Taipei.Level)
                        Select Case New_Taipei.Owner
                            Case 1
                                Form1.a.Money += New_Taipei.fee(New_Taipei.Level)
                            Case 2
                                Form1.b.Money += New_Taipei.fee(New_Taipei.Level)
                            Case 3
                                Form1.c.Money += New_Taipei.fee(New_Taipei.Level)
                            Case 4
                                Form1.d.Money += New_Taipei.fee(New_Taipei.Level)
                        End Select
                    End If
                End If
            Case 29 '事件
                Destiny(position)
            Case 30 '台北
                If Taipei.Owner = 0 Then
                    Dim reply = MsgBox("你來到了台北，此空地價格為" & Taipei.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Taipei.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Taipei.Buy
                            Taipei.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "台北"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "台北"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "台北"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "台北"
                            End Select
                            Taipei.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Taipei.Owner = Form1.Current_player Then
                    Select Case Taipei.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Taipei.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taipei.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taipei.Build
                                    Taipei.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Taipei.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taipei.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Taipei.Build
                                    Taipei.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Taipei.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Taipei.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Taipei.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Taipei.Owner & "所持有，需付過路費" & Taipei.fee(Taipei.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Taipei.fee(Taipei.Level)
                        Select Case Taipei.Owner
                            Case 1
                                Form1.a.Money += Taipei.fee(Taipei.Level)
                            Case 2
                                Form1.b.Money += Taipei.fee(Taipei.Level)
                            Case 3
                                Form1.c.Money += Taipei.fee(Taipei.Level)
                            Case 4
                                Form1.d.Money += Taipei.fee(Taipei.Level)
                        End Select
                    End If
                End If
            Case 31 '基隆
                If Keelung.Owner = 0 Then
                    Dim reply = MsgBox("你來到了基隆，此空地價格為" & Keelung.Buy & "是否要購買它", MsgBoxStyle.YesNo, "購買空地")
                    If reply = MsgBoxResult.Yes Then
                        If Money < Keelung.Buy Then
                            MsgBox("現金不足，無法購買", , "金額不足")
                        Else
                            Money -= Keelung.Buy
                            Keelung.Owner = Form1.Current_player '購買後擁有者就是該玩家持有
                            Select Case Form1.Current_player
                                Case 1
                                    Form1.play_land1.Text = Form1.play_land1.Text & " " & "基隆"
                                Case 2
                                    Form1.play_land2.Text = Form1.play_land2.Text & " " & "基隆"
                                Case 3
                                    Form1.play_land3.Text = Form1.play_land3.Text & " " & "基隆"
                                Case 4
                                    Form1.play_land4.Text = Form1.play_land4.Text & " " & "基隆"
                            End Select
                            Keelung.Level = 0 '等級為0代表空地
                        End If
                    End If
                    '自己持有
                ElseIf Keelung.Owner = Form1.Current_player Then
                    Select Case Keelung.Level
                        Case 0
                            Dim reply = MsgBox("該空地是你所持有，新建房屋費用為" & Keelung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Keelung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Keelung.Build
                                    Keelung.Level = 1 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 1
                            Dim reply = MsgBox("該土地以擁有一棟房屋，新建房屋費用為" & Keelung.Build & "是否要建設房屋", MsgBoxStyle.YesNo, "建設房屋")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Keelung.Build Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Keelung.Build
                                    Keelung.Level = 2 '購買後擁有者就是該玩家持有
                                End If
                            End If
                        Case 2
                            Dim reply = MsgBox("該土地以擁有兩棟房屋，新建房屋費用為" & Keelung.SpecialBuild & "是否要建設地標", MsgBoxStyle.YesNo, "建設地標")
                            If reply = MsgBoxResult.Yes Then
                                If Money < Keelung.SpecialBuild Then
                                    MsgBox("現金不足，無法建設", , "金額不足")
                                Else
                                    Money -= Changhua.SpecialBuild
                                    Keelung.Level = 3 '購買後擁有者就是該玩家持有
                                End If
                            End If
                    End Select
                    '別人持有
                Else
                    Dim reply = MsgBox("該土地為玩家" & Keelung.Owner & "所持有，需付過路費" & Keelung.fee(Keelung.Level), MsgBoxStyle.OkOnly, "過路費")
                    If reply = MsgBoxResult.Ok Then
                        Money -= Keelung.fee(Keelung.Level)
                        Select Case Keelung.Owner
                            Case 1
                                Form1.a.Money += Keelung.fee(Keelung.Level)
                            Case 2
                                Form1.b.Money += Keelung.fee(Keelung.Level)
                            Case 3
                                Form1.c.Money += Keelung.fee(Keelung.Level)
                            Case 4
                                Form1.d.Money += Keelung.fee(Keelung.Level)
                        End Select
                    End If
                End If
        End Select

        'Play sound after losting money
        My.Computer.Audio.Play(My.Resources.ResourceManager.GetObject("coindrop"),
                                AudioPlayMode.Background)

        Form1.Player_Update()

        If Form1.a.Check() = True Or Form1.b.Check() = True Or Form1.c.Check() = True Or Form1.d.Check() = True Then
            Form1.Bankrupt()
        End If

        '遊戲結束
        If Form1.Gameover_player = Form1.Total_player - 1 Or Form1.a.Money >= Form1.EndCoinNum Or Form1.b.Money >= Form1.EndCoinNum Or Form1.c.Money >= Form1.EndCoinNum Or Form1.d.Money >= Form1.EndCoinNum Or Form1.Round > Form1.EndDayNum Then
            Form1.GameOver()
        End If
    End Sub


#End Region

#Region "檢查是否破產"
    '檢查是否破產
    Public Function Check()
        If Money < 0 And Survival_statestatus = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region


#Region "命運事件"
    Private Sub Destiny(ByVal DestinyPosition)
        Randomize()
        Dim DestinyNum = CInt(Int(10 * Rnd())) + 1
        Select Case DestinyNum
            Case 1  '吃火鷄肉飯
                MsgBox("前往嘉義吃火鷄肉飯" & vbCrLf &
                       "經過起點領取現金！", , "事件")
                'Me.Move(31 - DestinyPosition + 1)
                Select Case Position
                    Case 5
                        Move(12)
                    Case 13
                        Move(4)
                    Case 21
                        Location = New Point(Form1.coordinates(29, 0), Form1.coordinates(29, 1))
                        direction = 4
                        Position = 29
                        Move(20)
                    Case 29
                        Move(20)
                End Select
            Case 2 '繳交程式課學費
                MsgBox("繳交程式課學費" & vbCrLf &
                       "需付費5000元！", , "事件")
                Me.Money -= 5000
            Case 3 '直接坐牢
                MsgBox("你違反中華民國刑法！" & vbCrLf &
                       "直接坐牢！", , "事件")
                If Position > 24 Then
                    Me.Move(32 - Position + 24)
                Else
                    Me.Move(24 - Position)
                End If
            Case 4 '改車
                MsgBox("車輛毀損！" & vbCrLf &
                       "需支付4000元修車費！", , "事件")
                Me.Money -= 4000
            Case 5 '冠軍
                MsgBox("100公尺賽跑冠軍！" & vbCrLf &
                       "獲得獎金4000元！", , "事件")
                Me.Money += 4000
            Case 6 '入獄豁免
                MsgBox("獲得私人飛機！" & vbCrLf &
                       "下次坐牢直接逃到國外（豁免一次）", , "事件")
                Me.Jail_status -= 1
            Case 7 '酗酒鬧事
                MsgBox("酒後鬧事！" & vbCrLf &
                       "重罰 8000元！", , "事件")
                Me.Money -= 8000
            Case 8 '壽星
                MsgBox("今日" & Me.Name & "生日！" & vbCrLf &
                       "所有人支付給壽星2000元！", , "事件")
                CelebrateMoney()
            Case 9 '領薪水
                MsgBox("打工發薪日！" & vbCrLf &
                       "獲得底薪6000元！", , "事件")
                Me.Money += 6000
            Case 10 '風景區
                MsgBox("前往八卦山朝拜！" & vbCrLf &
                       "移動到彰化經過起點可領取2000元", , "事件")
                Select Case Position
                    Case 5
                        Move(15)
                    Case 13
                        Move(7)
                    Case 21
                        Location = New Point(Form1.coordinates(29, 0), Form1.coordinates(29, 1))
                        direction = 4
                        Position = 29
                        Move(23)
                    Case 29
                        Move(23)
                End Select
        End Select
        Form1.Player_Update()
    End Sub

    Sub CelebrateMoney()
        Dim Celebrate As Integer = 0
        Select Case Form1.Current_player
            Case 1
                If Form1.b.Survival_statestatus = 1 Then
                    Form1.b.Money -= 2000
                    Celebrate += 1
                End If

                If Form1.c.Survival_statestatus = 1 Then
                    Form1.c.Money -= 2000
                    Celebrate += 1
                End If

                If Form1.d.Survival_statestatus = 1 Then
                    Form1.d.Money -= 2000
                    Celebrate += 1
                End If

                Me.Money += 2000 * Celebrate
            Case 2
                If Form1.a.Survival_statestatus = 1 Then
                    Form1.a.Money -= 2000
                    Celebrate += 1
                End If
                If Form1.c.Survival_statestatus = 1 Then
                    Form1.c.Money -= 2000
                    Celebrate += 1
                End If
                If Form1.d.Survival_statestatus = 1 Then
                    Form1.d.Money -= 2000
                    Celebrate += 1
                End If

                Me.Money += 2000 * Celebrate
            Case 3
                If Form1.a.Survival_statestatus = 1 Then
                    Form1.a.Money -= 2000
                    Celebrate += 1
                End If
                If Form1.b.Survival_statestatus = 1 Then
                    Form1.b.Money -= 2000
                    Celebrate += 1
                End If
                If Form1.d.Survival_statestatus = 1 Then
                    Form1.d.Money -= 2000
                    Celebrate += 1

                End If
                Me.Money += 2000 * Celebrate
            Case 4
                If Form1.a.Survival_statestatus = 1 Then
                    Form1.a.Money -= 2000
                    Celebrate += 1
                End If
                If Form1.b.Survival_statestatus = 1 Then
                    Form1.b.Money -= 2000
                    Celebrate += 1
                End If
                If Form1.c.Survival_statestatus = 1 Then
                    Form1.c.Money -= 2000
                    Celebrate += 1
                End If
                Me.Money += 2000 * Celebrate
        End Select
    End Sub
#End Region

End Class
#End Region

#Region "土地類別"

Public Class Kinmen '金門
    Public Const Buy As Integer = 2600 '土地購買金額
    Public Const Build As Integer = 1500 '建設房屋金額
    Public Const SpecialBuild As Integer = 4500 '建設地標金額
    Public Shared Level As Integer = -1 ' 土地等級(-1代表無人持有 0為空地; 1為1間房; 2為2間房 ;3為地標)
    Public Shared fee(3) As Integer '收費金額(0為空地; 1為1間房; 2為2間房 ;3為地標)
    Public Shared Owner As Integer '擁有者是誰 (0代表無人持有; 1代表玩家1 依此類推)
    Public Shared Sub Initialize() '初始化
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class

Public Class Mazu '馬祖
    Public Const Buy As Integer = 2400
    Public Const Build As Integer = 1000
    Public Const SpecialBuild As Integer = 3000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class

Public Class Liuchiu '小琉球
    Public Const Buy As Integer = 2000
    Public Const Build As Integer = 1000
    Public Const SpecialBuild As Integer = 3000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class

Public Class Koto_island '蘭嶼
    Public Const Buy As Integer = 2600
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize() '初始化
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class

Public Class Lyudao '綠島
    Public Const Buy As Integer = 2200
    Public Const Build As Integer = 1100
    Public Const SpecialBuild As Integer = 3300
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class
Public Class Penghu '澎湖
    Public Const Buy As Integer = 2000
    Public Const Build As Integer = 1000
    Public Const SpecialBuild As Integer = 3000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class
Public Class Taitung '台東
    Public Const Buy As Integer = 2600
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Hualien '花蓮
    Public Const Buy As Integer = 3000
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Pingtung '屏東
    Public Const Buy As Integer = 2200
    Public Const Build As Integer = 1000
    Public Const SpecialBuild As Integer = 3000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class
Public Class Kaohsiung '高雄
    Public Const Buy As Integer = 3000
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Tainan '台南
    Public Const Buy As Integer = 2400
    Public Const Build As Integer = 1000
    Public Const SpecialBuild As Integer = 3000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class
Public Class Chiayi '嘉義
    Public Const Buy As Integer = 3600
    Public Const Build As Integer = 2000
    Public Const SpecialBuild As Integer = 6000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 400 : fee(1) = 2000 : fee(2) = 8000 : fee(3) = 16000
        Owner = 0
    End Sub
End Class
Public Class Nantou '南投
    Public Const Buy As Integer = 3400
    Public Const Build As Integer = 2000
    Public Const SpecialBuild As Integer = 6000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 400 : fee(1) = 2000 : fee(2) = 8000 : fee(3) = 16000
        Owner = 0
    End Sub
End Class
Public Class Yunlin '雲林
    Public Const Buy As Integer = 3600
    Public Const Build As Integer = 2000
    Public Const SpecialBuild As Integer = 6000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 400 : fee(1) = 2000 : fee(2) = 8000 : fee(3) = 16000
        Owner = 0
    End Sub
End Class
Public Class Changhua ' 彰化
    Public Const Buy As Integer = 3400
    Public Const Build As Integer = 2000
    Public Const SpecialBuild As Integer = 6000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 400 : fee(1) = 2000 : fee(2) = 8000 : fee(3) = 16000
        Owner = 0
    End Sub
End Class
Public Class Taichung '台中
    Public Const Buy As Integer = 2400
    Public Const Build As Integer = 1000
    Public Const SpecialBuild As Integer = 3000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 200 : fee(1) = 1200 : fee(2) = 6000 : fee(3) = 11000
        Owner = 0
    End Sub
End Class
Public Class Miaoli '苗栗
    Public Const Buy As Integer = 2600
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Yilan '宜蘭
    Public Const Buy As Integer = 3000
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Hsinchu '新竹
    Public Const Buy As Integer = 3200
    Public Const Build As Integer = 2000
    Public Const SpecialBuild As Integer = 6000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 400 : fee(1) = 2000 : fee(2) = 8000 : fee(3) = 16000
        Owner = 0
    End Sub
End Class
Public Class Taoyuan '桃園
    Public Const Buy As Integer = 3200
    Public Const Build As Integer = 2000
    Public Const SpecialBuild As Integer = 6000
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 400 : fee(1) = 2000 : fee(2) = 8000 : fee(3) = 16000
        Owner = 0
    End Sub
End Class
Public Class New_Taipei '新北
    Public Const Buy As Integer = 2800
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Taipei '台北
    Public Const Buy As Integer = 2800
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
Public Class Keelung '基隆
    Public Const Buy As Integer = 3000
    Public Const Build As Integer = 1500
    Public Const SpecialBuild As Integer = 4500
    Public Shared Level As Integer = -1
    Public Shared fee(3) As Integer
    Public Shared Owner As Integer
    Public Shared Sub Initialize()
        Level = -1
        fee(0) = 300 : fee(1) = 1600 : fee(2) = 7000 : fee(3) = 13000
        Owner = 0
    End Sub
End Class
#End Region