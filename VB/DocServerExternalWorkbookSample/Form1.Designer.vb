Namespace DocServerExternalWorkbookSample

    Partial Class Form1

        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (Me.components IsNot Nothing) Then
                Me.components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

'#Region "Windows Form Designer generated code"
        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.button1 = New System.Windows.Forms.Button()
            Me.button2 = New System.Windows.Forms.Button()
            Me.SuspendLayout()
            ' 
            ' button1
            ' 
            Me.button1.Location = New System.Drawing.Point(65, 46)
            Me.button1.Name = "button1"
            Me.button1.Size = New System.Drawing.Size(155, 23)
            Me.button1.TabIndex = 0
            Me.button1.Text = "Add External Workbook"
            Me.button1.UseVisualStyleBackColor = True
            AddHandler Me.button1.Click, New System.EventHandler(AddressOf Me.button1_Click)
            ' 
            ' button2
            ' 
            Me.button2.Location = New System.Drawing.Point(65, 89)
            Me.button2.Name = "button2"
            Me.button2.Size = New System.Drawing.Size(155, 23)
            Me.button2.TabIndex = 1
            Me.button2.Text = "Insert External Reference"
            Me.button2.UseVisualStyleBackColor = True
            AddHandler Me.button2.Click, New System.EventHandler(AddressOf Me.button2_Click)
            ' 
            ' Form1
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(284, 186)
            Me.Controls.Add(Me.button2)
            Me.Controls.Add(Me.button1)
            Me.Name = "Form1"
            Me.Text = "Form1"
            Me.ResumeLayout(False)
        End Sub

'#End Region
        Private button1 As System.Windows.Forms.Button

        Private button2 As System.Windows.Forms.Button
    End Class
End Namespace
