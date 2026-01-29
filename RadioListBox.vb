
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles

Public Class RadioListBox
    Inherits ListBox

    Private ReadOnly Align As StringFormat

    Private IsTransparent As Boolean = False

    Private BackBrush As Brush

    Public Overrides Property BackColor As Color
        Get
            If IsTransparent Then Return Color.Transparent Else Return MyBase.BackColor
        End Get

        Set(ByVal value As Color)
            If value = Color.Transparent Then
                IsTransparent = True
                MyBase.BackColor = If((Me.Parent Is Nothing), SystemColors.Window, Me.Parent.BackColor)
            Else
                IsTransparent = False
                MyBase.BackColor = value
            End If

            If Me.BackBrush IsNot Nothing Then Me.BackBrush.Dispose()
            BackBrush = New SolidBrush(MyBase.BackColor)
            Invalidate()
        End Set
    End Property

    <Browsable(False)>
    Public Overrides Property DrawMode As DrawMode
        Get
            Return MyBase.DrawMode
        End Get

        Set(ByVal value As DrawMode)
            If value <> DrawMode.OwnerDrawFixed Then Throw New Exception("Invalid value for DrawMode property") Else MyBase.DrawMode = value
        End Set
    End Property

    <Browsable(False)>
    Public Overrides Property SelectionMode As SelectionMode
        Get
            Return MyBase.SelectionMode
        End Get

        Set(ByVal value As SelectionMode)
            If value <> SelectionMode.One Then Throw New Exception("Invalid value for SelectionMode property") Else MyBase.SelectionMode = value
        End Set
    End Property

    Public Sub New()
        Me.DrawMode = DrawMode.OwnerDrawFixed
        Me.SelectionMode = SelectionMode.One
        Me.ItemHeight = Me.FontHeight
        Me.Align = New StringFormat(StringFormat.GenericDefault)
        Me.Align.LineAlignment = StringAlignment.Center
        Me.BackColor = Me.BackColor
    End Sub

    Protected Overrides Sub OnDrawItem(ByVal e As DrawItemEventArgs)
        Dim maxItem As Integer = Me.Items.Count - 1
        If e.Index < 0 OrElse e.Index > maxItem Then
            e.Graphics.FillRectangle(BackBrush, Me.ClientRectangle)
            Return
        End If

        'Dim size As Integer = e.Font.Height
        Dim backRect As Rectangle = e.Bounds
        If e.Index = maxItem Then backRect.Height = Me.ClientRectangle.Top + Me.ClientRectangle.Height - e.Bounds.Top
        e.Graphics.FillRectangle(BackBrush, backRect)
        Dim textBrush As Brush
        Dim isChecked As Boolean = (e.State And DrawItemState.Selected) = DrawItemState.Selected
        Dim state As RadioButtonState = If(isChecked, RadioButtonState.CheckedNormal, RadioButtonState.UncheckedNormal)
        If (e.State And DrawItemState.Disabled) = DrawItemState.Disabled Then
            textBrush = SystemBrushes.GrayText
            state = If(isChecked, RadioButtonState.CheckedDisabled, RadioButtonState.UncheckedDisabled)
        ElseIf (e.State And DrawItemState.Grayed) = DrawItemState.Grayed Then
            textBrush = SystemBrushes.GrayText
            state = If(isChecked, RadioButtonState.CheckedDisabled, RadioButtonState.UncheckedDisabled)
        Else
            textBrush = SystemBrushes.FromSystemColor(Me.ForeColor)
        End If

        Dim glyphSize As Size = RadioButtonRenderer.GetGlyphSize(e.Graphics, state)
        Dim glyphLocation As Point = e.Bounds.Location
        glyphLocation.Y += CInt((e.Bounds.Height - glyphSize.Height) / 2)
        Dim RectBounds As Rectangle = New Rectangle(e.Bounds.X + glyphSize.Width, e.Bounds.Y, e.Bounds.Width - glyphSize.Width, e.Bounds.Height)
        RadioButtonRenderer.DrawRadioButton(e.Graphics, glyphLocation, state)
        If Not String.IsNullOrEmpty(DisplayMember) Then
            e.Graphics.DrawString((CType(Me.Items(e.Index), System.Data.DataRowView))(Me.DisplayMember).ToString(), e.Font, textBrush, RectBounds, Me.Align)
        Else
            e.Graphics.DrawString(Me.Items(e.Index).ToString(), e.Font, textBrush, RectBounds, Me.Align)
        End If
        e.DrawFocusRectangle()
    End Sub

    Protected Overrides Sub DefWndProc(ByRef m As Message)
        If m.Msg = 20 Then
            m.Result = CType(1, IntPtr)
            Return
        End If

        MyBase.DefWndProc(m)
    End Sub

    Protected Overrides Sub OnHandleCreated(ByVal e As EventArgs)
        If Me.FontHeight > Me.ItemHeight Then Me.ItemHeight = Me.FontHeight
        MyBase.OnHandleCreated(e)
    End Sub

    Protected Overrides Sub OnFontChanged(ByVal e As EventArgs)
        MyBase.OnFontChanged(e)
        If Me.FontHeight > Me.ItemHeight Then Me.ItemHeight = Me.FontHeight
        Update()
    End Sub

    Protected Overrides Sub OnParentChanged(ByVal e As EventArgs)
        Me.BackColor = Me.BackColor
    End Sub

    Protected Overrides Sub OnParentBackColorChanged(ByVal e As EventArgs)
        Me.BackColor = Me.BackColor
    End Sub
End Class
