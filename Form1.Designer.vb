﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        DiceTimer1 = New Timer(components)
        DiceTimer2 = New Timer(components)
        SuspendLayout()
        ' 
        ' DiceTimer1
        ' 
        ' 
        ' DiceTimer2
        ' 
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(622, 355)
        Margin = New Padding(2, 2, 2, 2)
        Name = "Form1"
        Text = "Form1"
        ResumeLayout(False)

    End Sub

    Friend WithEvents DiceTimer1 As Timer
    Friend WithEvents DiceTimer2 As Timer
    Friend WithEvents Button1 As Button

End Class
