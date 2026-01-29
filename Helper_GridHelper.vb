Imports System.Drawing
Imports DevComponents.DotNetBar
Imports DevComponents.DotNetBar.SuperGrid
Imports DevComponents.Editors

Public Class Helper_GridHelper
    Public Enum GVInputType
        IntegerInput
        DoubleInput
        StringInput
        BooleanInput
        NoInput
    End Enum

    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal Text As String, ByVal Bold As Boolean, ByVal Optional TextColor As Color = Nothing, ByVal Optional BackColor1 As Color = Nothing, ByVal Optional BackColor2 As Color = Nothing, ByVal Optional GradientAngle As Integer = 0, ByVal Optional visible As Boolean = True, ByVal Optional image As Image = Nothing, ByVal Optional Tag As Object = Nothing, ByVal Optional AllowEdit As Boolean = False, ByVal Optional InputType As GVInputType = GVInputType.StringInput, ByVal Optional IsReadOnly As Boolean = True, ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        GridCell.Tag = Tag
        GridCell.Visible = visible
        If Not IsNothing(image) Then
            GridCell.CellStyles.Default.Image = image
        End If
        If Not IsNothing(TextColor) Then
            GridCell.CellStyles.Default.TextColor = TextColor
        End If
        If Not IsNothing(BackColor1) Then
            GridCell.CellStyles.Default.Background.Color1 = BackColor1
        End If
        If Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.Color2 = BackColor2
        End If
        If Not IsNothing(BackColor1) AndAlso Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.GradientAngle = GradientAngle
            GridCell.CellStyles.Default.Background.BackFillType = Style.BackFillType.VerticalCenter
        End If
        If Bold Then
            Dim font As New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            GridCell.CellStyles.Default.Font = font
        End If
        GridCell.AllowEdit = AllowEdit
        If Not AllowEdit Then
            'GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf040"))
            'GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf11c"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Pink
        End If
        GridCell.ReadOnly = IsReadOnly
        If IsReadOnly Then
            'GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf040"))
            'GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf11c"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.LightGray
        Else
        End If
        If InputType = GVInputType.IntegerInput Then
            GridCell.EditorType = GetType(GridIntegerInputEditControl)
        ElseIf InputType = GVInputType.DoubleInput Then
            GridCell.EditorType = GetType(GridDoubleInputEditControl)
        ElseIf InputType = GVInputType.BooleanInput Then
            GridCell.EditorType = GetType(GridCheckBoxEditControl)
        ElseIf InputType = GVInputType.StringInput Then
            GridCell.EditorType = GetType(GridTextBoxXEditControl)
        End If
        GridCell.CellStyles.Default.Alignment = Alignment

        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal ValDbl As Double, ByVal Bold As Boolean, ByVal Optional TextColor As Color = Nothing, ByVal Optional BackColor1 As Color = Nothing, ByVal Optional BackColor2 As Color = Nothing, ByVal Optional GradientAngle As Integer = 0, ByVal Optional visible As Boolean = True, ByVal Optional image As Image = Nothing, ByVal Optional Tag As Object = Nothing, ByVal Optional AllowEdit As Boolean = False, ByVal Optional InputType As GVInputType = GVInputType.StringInput, ByVal Optional IsReadOnly As Boolean = True, ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = Format(ValDbl, "0.00")
        GridCell.Tag = Tag
        GridCell.Visible = visible
        If Not IsNothing(image) Then
            GridCell.CellStyles.Default.Image = image
        End If
        If Not IsNothing(TextColor) Then
            GridCell.CellStyles.Default.TextColor = TextColor
        End If
        If Not IsNothing(BackColor1) Then
            GridCell.CellStyles.Default.Background.Color1 = BackColor1
        End If
        If Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.Color2 = BackColor2
        End If
        If Not IsNothing(BackColor1) AndAlso Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.GradientAngle = GradientAngle
            GridCell.CellStyles.Default.Background.BackFillType = Style.BackFillType.VerticalCenter
        End If
        If Bold Then
            Dim font As New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            GridCell.CellStyles.Default.Font = font
        End If
        GridCell.AllowEdit = AllowEdit
        If Not AllowEdit Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf040"))
            'GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf11c"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Pink
        End If
        GridCell.ReadOnly = IsReadOnly
        If IsReadOnly Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf040"))
            'GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf11c"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.LightGray
        Else
        End If
        If InputType = GVInputType.IntegerInput Then
            GridCell.EditorType = GetType(GridIntegerInputEditControl)
        ElseIf InputType = GVInputType.DoubleInput Then
            GridCell.EditorType = GetType(GridDoubleInputEditControl)
        ElseIf InputType = GVInputType.BooleanInput Then
            GridCell.EditorType = GetType(GridCheckBoxEditControl)
        ElseIf InputType = GVInputType.StringInput Then
            GridCell.EditorType = GetType(GridTextBoxXEditControl)
        End If
        GridCell.CellStyles.Default.Alignment = Alignment

        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal Text As String, ByVal ShowCombo As Boolean, ByVal ComboContent As List(Of ComboItem), ByVal Optional BackColor1 As Color = Nothing, ByVal Optional BackColor2 As Color = Nothing, ByVal Optional GradientAngle As Integer = 0, ByVal Optional IsReadOnly As Boolean = True, ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf0dd"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.CellStyles.Default.Alignment = Alignment
        If Not IsNothing(BackColor1) Then
            GridCell.CellStyles.Default.Background.Color1 = BackColor1
        End If
        If Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.Color2 = BackColor2
        End If
        If Not IsNothing(BackColor1) AndAlso Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.GradientAngle = GradientAngle
            GridCell.CellStyles.Default.Background.BackFillType = Style.BackFillType.VerticalCenter
        End If

        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal Text As String, ByVal ShowCombo As Boolean, ByVal ComboContent As List(Of String), ByVal Optional BackColor1 As Color = Nothing, ByVal Optional BackColor2 As Color = Nothing, ByVal Optional GradientAngle As Integer = 0, ByVal Optional IsReadOnly As Boolean = True, ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt("&Hf0dd"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.CellStyles.Default.Alignment = Alignment
        If Not IsNothing(BackColor1) Then
            GridCell.CellStyles.Default.Background.Color1 = BackColor1
        End If
        If Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.Color2 = BackColor2
        End If
        If Not IsNothing(BackColor1) AndAlso Not IsNothing(BackColor2) Then
            GridCell.CellStyles.Default.Background.GradientAngle = GradientAngle
            GridCell.CellStyles.Default.Background.BackFillType = Style.BackFillType.VerticalCenter
        End If

        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal Text As String, ByVal symbol As String, ByVal symbolcolor As Color, ByVal Optional tag As String = "", ByVal Optional Bold As Boolean = False, ByVal Optional TextColor As Color = Nothing, ByVal Optional BackColor As Color = Nothing, ByVal Optional IsReadOnly As Boolean = True, ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        GridCell.Tag = tag
        If symbol <> "" Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(CInt(symbol))
        End If
        GridCell.CellStyles.Default.SymbolDef.SymbolColor = symbolcolor
        If Not IsNothing(TextColor) Then
            GridCell.CellStyles.Default.TextColor = TextColor
        End If
        If Not IsNothing(BackColor) Then
            GridCell.CellStyles.Default.Background.Color1 = BackColor
        End If
        If Bold Then
            Dim font As New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            GridCell.CellStyles.Default.Font = font
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.CellStyles.Default.Alignment = Alignment

        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal image As Image, ByVal Optional Tag As String = "", ByVal Optional visible As Boolean = True, ByVal Optional IsReadOnly As Boolean = True) As GridRow
        Dim GridCell As New GridCell
        GridCell.CellStyles.Default.Image = image
        GridCell.Tag = Tag
        GridCell.Value = ""
        GridCell.Visible = visible
        GridCell.ReadOnly = IsReadOnly
        GridCell.CellStyles.Default.ImageAlignment = Style.Alignment.MiddleCenter
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow, ByVal checked As Boolean, ByVal Optional visible As Boolean = True, ByVal Optional IsReadOnly As Boolean = True) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = checked
        GridCell.Visible = visible
        GridCell.ReadOnly = IsReadOnly
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function

#Region "FragrantComboBox"

    Friend Class FragrantComboBox
        Inherits GridComboBoxExEditControl
        Public Sub New(ByVal orderArray As IEnumerable(Of String))
            DataSource = orderArray
        End Sub
    End Class

#End Region
End Class
