Imports System.Drawing
Public Class beginning
#Region "全域變數"
    Dim pickCrCount As Integer = 0
    Public Shared mainCrPick(4) As Integer
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

    Public Shared EndCoinNum As Integer = 150000
    Public Shared EndDayNum As Integer = 60
    Public Shared BeginCionNum As Integer = 15000
#End Region
    '抓客戶電腦畫面
    Dim screenRectangle As Rectangle = Screen.PrimaryScreen.WorkingArea
    Dim x As Integer = (screenRectangle.Width - Me.Width) \ 2
    Dim y As Integer = (screenRectangle.Height - Me.Height) \ 2
    Dim x1 As Integer = (screenRectangle.Width - Me.Width) \ 3
    Dim y1 As Integer = (screenRectangle.Height - Me.Height)

#End Region

#Region "首頁&& 初始化"
#Region "初始化"
    Private Sub beginning_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        hidePickCr()
        hidechoose()
        'beginning
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1582, 1053)
        Me.SuspendLayout()
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "beginning"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.WindowState = FormWindowState.Maximized
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
        Me.gochoiseGame.Location = New System.Drawing.Point(x1 + 700, y1 + 175)
        Me.gochoiseGame.Text = "下一步"
        Me.gochoiseGame.Name = "gochoiseGame"
        Me.gochoiseGame.Size = New System.Drawing.Size(100, 40)
        Me.gochoiseGame.TabIndex = 14
        Me.gochoiseGame.UseVisualStyleBackColor = True
        '
        'backBigin
        AddHandler backBigin.Click, AddressOf backBigin_Click
        Me.Controls.Add(backBigin)
        Me.backBigin.Text = "回首頁"
        Me.backBigin.Location = New System.Drawing.Point(x1 - 150, y1 + 75)
        Me.backBigin.Name = "backBigin"
        Me.backBigin.Size = New System.Drawing.Size(100, 40)
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
        Me.PlayerPick1.Location = New System.Drawing.Point(x1 - 125, y1 - 145)
        Me.PlayerPick1.Name = "PlayerPick1"
        Me.PlayerPick1.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick1.TabIndex = 4
        Me.PlayerPick1.TabStop = False
        '

        'PlayerPick2
        Me.Controls.Add(PlayerPick2)
        Me.PlayerPick1.BackColor = Color.Transparent
        Me.PlayerPick2.Location = New System.Drawing.Point(x1 + 145, y1 - 145)
        Me.PlayerPick2.Name = "PlayerPick2"
        Me.PlayerPick2.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick2.TabIndex = 5
        Me.PlayerPick2.TabStop = False
        '
        'PlayerPick3
        Controls.Add(PlayerPick3)
        Me.PlayerPick3.BackColor = Color.Transparent
        Me.PlayerPick3.Location = New System.Drawing.Point(x1 + 430, y1 - 145)
        Me.PlayerPick3.Name = "PlayerPick3"
        Me.PlayerPick3.Size = New System.Drawing.Size(130, 130)
        Me.PlayerPick3.TabIndex = 6
        Me.PlayerPick3.TabStop = False
        '
        'PlayerPick4
        Me.Controls.Add(PlayerPick4)
        Me.PlayerPick4.BackColor = Color.Transparent
        Me.PlayerPick4.Location = New System.Drawing.Point(x1 + 710, y1 - 145)
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
        goMainGame.Location = New System.Drawing.Point(x + 500, y + 300)
        goMainGame.Size = New System.Drawing.Size(100, 40)
        goMainGame.TabIndex = 19
        goMainGame.UseVisualStyleBackColor = True

        AddHandler backCr.Click, AddressOf backCr_Click
        Me.Controls.Add(backCr)
        backCr.BackColor = Color.Transparent
        backCr.FlatStyle = FlatStyle.Popup
        backCr.Location = New System.Drawing.Point(x - 300, y + 300)
        backCr.Size = New System.Drawing.Size(100, 40)
        backCr.TabIndex = 19
        backCr.UseVisualStyleBackColor = True

        AddHandler leftBeginCion.Click, AddressOf leftBeginCion_Click
        Me.Controls.Add(leftBeginCion)
        leftBeginCion.BackColor = Color.Transparent
        leftBeginCion.FlatStyle = FlatStyle.Popup
        leftBeginCion.Location = New System.Drawing.Point(x - 100, y)
        leftBeginCion.Size = New System.Drawing.Size(100, 40)
        leftBeginCion.TabIndex = 19
        leftBeginCion.UseVisualStyleBackColor = True

        AddHandler leftEndDay.Click, AddressOf leftEndDay_Click
        Me.Controls.Add(leftEndDay)
        leftEndDay.BackColor = Color.Transparent
        leftEndDay.FlatStyle = FlatStyle.Popup
        leftEndDay.Location = New System.Drawing.Point(x - 100, y + 150)
        leftEndDay.Size = New System.Drawing.Size(100, 40)
        leftEndDay.TabIndex = 19
        leftEndDay.UseVisualStyleBackColor = True

        AddHandler leftEndCion.Click, AddressOf leftEndCion_Click
        Me.Controls.Add(leftEndCion)
        leftEndCion.BackColor = Color.Transparent
        leftEndCion.FlatStyle = FlatStyle.Popup
        leftEndCion.Location = New System.Drawing.Point(x - 100, y + 300)
        leftEndCion.Size = New System.Drawing.Size(100, 40)
        leftEndCion.TabIndex = 19
        leftEndCion.UseVisualStyleBackColor = True

        AddHandler rightBeginCion.Click, AddressOf rightBeginCion_Click
        Me.Controls.Add(rightBeginCion)
        rightBeginCion.BackColor = Color.Transparent
        rightBeginCion.FlatStyle = FlatStyle.Popup
        rightBeginCion.Location = New System.Drawing.Point(x + 300, y)
        rightBeginCion.Size = New System.Drawing.Size(100, 40)
        rightBeginCion.TabIndex = 19
        rightBeginCion.UseVisualStyleBackColor = True

        AddHandler rightEndDay.Click, AddressOf rightEndDay_Click
        Me.Controls.Add(rightEndDay)
        rightEndDay.BackColor = Color.Transparent
        rightEndDay.FlatStyle = FlatStyle.Popup
        rightEndDay.Location = New System.Drawing.Point(x + 300, y + 150)
        rightEndDay.Size = New System.Drawing.Size(100, 40)
        rightEndDay.TabIndex = 19
        rightEndDay.UseVisualStyleBackColor = True



        AddHandler rightEndCoin.Click, AddressOf rightEndCoin_Click
        Me.Controls.Add(rightEndCoin)
        rightEndCoin.BackColor = Color.Transparent
        rightEndCoin.FlatStyle = FlatStyle.Popup
        rightEndCoin.Location = New System.Drawing.Point(x + 300, y + 300)
        rightEndCoin.Size = New System.Drawing.Size(100, 40)
        rightEndCoin.TabIndex = 19
        rightEndCoin.UseVisualStyleBackColor = True

        AddHandler BeginCion.Click, AddressOf BeginCion_Click
        Me.Controls.Add(BeginCion)
        BeginCion.BackColor = Color.Transparent
        BeginCion.FlatStyle = FlatStyle.Popup
        Me.BeginCion.Location = New System.Drawing.Point(x + 50, y)
        Me.BeginCion.Name = "BeginCion"
        Me.BeginCion.Size = New System.Drawing.Size(200, 45)
        Me.BeginCion.TabIndex = 24
        Me.BeginCion.TabStop = False

        AddHandler EndDay.Click, AddressOf EndDay_Click
        Me.Controls.Add(EndDay)
        EndDay.BackColor = Color.Transparent
        EndDay.FlatStyle = FlatStyle.Popup
        EndDay.Location = New System.Drawing.Point(x + 50, y + 150)
        EndDay.Size = New System.Drawing.Size(200, 45)
        EndDay.TabIndex = 19
        EndDay.TabStop = False

        AddHandler EndCoin.Click, AddressOf EndCoin_Click
        Me.Controls.Add(EndCoin)
        EndCoin.BackColor = Color.Transparent
        EndCoin.FlatStyle = FlatStyle.Popup
        Me.EndCoin.Location = New System.Drawing.Point(x + 50, y + 300)
        Me.EndCoin.Name = "BeginCion"
        Me.EndCoin.Size = New System.Drawing.Size(200, 45)
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

#End Region
#Region "首頁按鈕"
    Private Sub OpenBegin_Click(sender As Object, e As EventArgs)
        '主頁面隱藏
        hideMain()
        showPickCr()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("pickCrScience")

    End Sub

    Private Sub LeaveGane_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub Adjustment_Click(sender As Object, e As EventArgs)

    End Sub
#End Region

#End Region

#Region "pickCr事件"

#Region "點選'選擇腳色'事件"
    Private Sub pickCr1_Click(sender As Object, e As EventArgs)
        pickCr1.Enabled = False
        mainCrPick(pickCrCount) = 1
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr2_Click(sender As Object, e As EventArgs)
        pickCr2.Enabled = False
        mainCrPick(pickCrCount) = 2
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr3_Click(sender As Object, e As EventArgs)
        pickCr3.Enabled = False
        mainCrPick(pickCrCount) = 3
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr4_Click(sender As Object, e As EventArgs)
        pickCr4.Enabled = False
        mainCrPick(pickCrCount) = 4
        pickCrCount += 1
        If pickCrCount = 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub pickCr5_Click(sender As Object, e As EventArgs)
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
            If pickCrCount = 1 Then
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
        hidechoose()
        Me.BackgroundImage = Nothing
    End Sub
    Private Sub backCr_Click(sender As Object, e As EventArgs)
        hidechoose()
        showPickCr()
        Me.BackgroundImage = My.Resources.ResourceManager.GetObject("pickCrScience")
        EndCoinNum = 100000
        EndDayNum = 60
        BeginCionNum = 15000
        If pickCrCount >= 2 Then
            gochoiseGame.Show()
        End If
    End Sub

    Private Sub leftBeginCion_Click(sender As Object, e As EventArgs)
        BeginCionNum = 10000
        leftBeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        BeginCion.BackColor = Color.Transparent
        rightBeginCion.BackColor = Color.Transparent
        leftBeginCion.Enabled = False
        BeginCion.Enabled = True
        rightBeginCion.Enabled = True
    End Sub

    Private Sub BeginCion_Click(sender As Object, e As EventArgs)
        BeginCionNum = 15000
        leftBeginCion.BackColor = Color.Transparent
        BeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        rightBeginCion.BackColor = Color.Transparent
        leftBeginCion.Enabled = True
        BeginCion.Enabled = False
        rightBeginCion.Enabled = True
    End Sub
    Private Sub rightBeginCion_Click(sender As Object, e As EventArgs)
        BeginCionNum = 20000
        leftBeginCion.BackColor = Color.Transparent
        BeginCion.BackColor = Color.Transparent
        rightBeginCion.BackColor = Color.FromArgb(128, Color.Gray)
        leftBeginCion.Enabled = True
        BeginCion.Enabled = True
        rightBeginCion.Enabled = False
    End Sub
    Private Sub rightEndDay_Click(sender As Object, e As EventArgs)
        EndDayNum = 100
        rightEndDay.BackColor = Color.FromArgb(128, Color.Gray)
        EndDay.BackColor = Color.Transparent
        leftEndDay.BackColor = Color.Transparent
        rightEndDay.Enabled = False
        EndDay.Enabled = True
        leftEndDay.Enabled = True
    End Sub
    Private Sub EndDay_Click(sender As Object, e As EventArgs)
        EndDayNum = 80
        rightEndDay.BackColor = Color.Transparent
        EndDay.BackColor = Color.FromArgb(128, Color.Gray)
        leftEndDay.BackColor = Color.Transparent
        rightEndDay.Enabled = True
        EndDay.Enabled = False
        leftEndDay.Enabled = True
    End Sub
    Private Sub leftEndDay_Click(sender As Object, e As EventArgs)
        EndDayNum = 60
        rightEndDay.BackColor = Color.Transparent
        EndDay.BackColor = Color.Transparent
        leftEndDay.BackColor = Color.FromArgb(128, Color.Gray)
        rightEndDay.Enabled = True
        EndDay.Enabled = True
        leftEndDay.Enabled = False
    End Sub

    Private Sub rightEndCoin_Click(sender As Object, e As EventArgs)
        EndCoinNum = 50000
        rightEndCoin.BackColor = Color.FromArgb(128, Color.Gray)
        EndCoin.BackColor = Color.Transparent
        leftEndCion.BackColor = Color.Transparent
        rightEndCoin.Enabled = False
        EndCoin.Enabled = True
        leftEndCion.Enabled = True
    End Sub

    Private Sub EndCoin_Click(sender As Object, e As EventArgs)
        EndCoinNum = 80000
        rightEndCoin.BackColor = Color.Transparent
        EndCoin.BackColor = Color.FromArgb(128, Color.Gray)
        leftEndCion.BackColor = Color.Transparent
        rightEndCoin.Enabled = True
        EndCoin.Enabled = False
        leftEndCion.Enabled = True
    End Sub
    Private Sub leftEndCion_Click(sender As Object, e As EventArgs)
        EndCoinNum = 100000
        rightEndCoin.BackColor = Color.Transparent
        EndCoin.BackColor = Color.Transparent
        leftEndCion.BackColor = Color.FromArgb(128, Color.Gray)
        rightEndCoin.Enabled = True
        EndCoin.Enabled = True
        leftEndCion.Enabled = False
    End Sub

#End Region

End Class
