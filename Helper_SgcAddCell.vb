Imports System.Drawing
Imports System.Threading
Imports DevComponents.DotNetBar.SuperGrid
Imports DevComponents.Editors
Imports TimeShare_Helper
Imports TimeShare_Error


Public Class Helper_SgcAddCell
    Public Enum GVInputType
        IntegerInput
        DoubleInput
        StringInput
        BooleanInput
        NoInput
    End Enum
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal checked As Boolean,
                             ByVal Optional visible As Boolean = True,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = checked
        GridCell.Visible = visible
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal image As Image,
                             ByVal Optional Tag As String = "",
                             ByVal Optional visible As Boolean = True,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.CellStyles.Default.Image = image
        GridCell.Tag = Tag
        'GridCell.Value = ""
        GridCell.Visible = visible
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridCell.CellStyles.Default.Alignment = Style.Alignment.MiddleLeft
        'GridCell.Value="1"
        'GridCell.ColumnIndex
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal SwitchValue As Boolean,
                             ByVal TextOn As String,
                             ByVal TextOff As String,
                             ByVal Optional visible As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = SwitchValue
        GridCell.Visible = visible
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal Optional Tag As String = "",
                             ByVal Optional Bold As Boolean = False,
                             ByVal Optional TextColor As Color = Nothing,
                             ByVal Optional BackColor1 As Color = Nothing,
                             ByVal Optional BackColor2 As Color = Nothing,
                             ByVal Optional GradientAngle As Integer = 0,
                             ByVal Optional visible As Boolean = True,
                             ByVal Optional image As Image = Nothing,
                             ByVal Optional AllowEdit As Boolean = False,
                             ByVal Optional IsNumericInput As Boolean = False,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True,
                             ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft,
                             ByVal Optional InputType As GVInputType = GVInputType.StringInput
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        GridCell.Visible = visible
        GridCell.Tag = Tag
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
        End If
        If Bold Then
            Dim font As New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            GridCell.CellStyles.Default.Font = font
        End If
        GridCell.AllowEdit = AllowEdit
        If AllowEdit Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&HF040"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
            'GridCell.InfoImage = My.Resources.info
            'GridCell.InfoText = "Eingabe"
        End If
        If IsNumericInput Then
            GridCell.EditorType = GetType(GridIntegerInputEditControl)
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection

        If Not IsNothing(Alignment) Then
            GridCell.CellStyles.Default.Alignment = Alignment
        End If
        If IsNothing(InputType) Then
        ElseIf InputType = GVInputType.IntegerInput Then
            GridCell.EditorType = GetType(GridIntegerInputEditControl)
        ElseIf InputType = GVInputType.DoubleInput Then
            GridCell.EditorType = GetType(GridDoubleInputEditControl)
        ElseIf InputType = GVInputType.BooleanInput Then
            GridCell.EditorType = GetType(GridCheckBoxEditControl)
        ElseIf InputType = GVInputType.StringInput Then
            GridCell.EditorType = GetType(GridTextBoxXEditControl)
        End If
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal symbol As Char,
                             ByVal symbolcolor As Color,
                             ByVal Optional tag As String = "",
                             ByVal Optional Bold As Boolean = False,
                             ByVal Optional TextColor As Color = Nothing,
                             ByVal Optional BackColor As Color = Nothing,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
          ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        GridCell.Tag = tag
        If Not IsNothing(symbol) Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = symbol
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
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal ShowCombo As Boolean,
                             ByVal ComboContent As List(Of ComboItem),
                             ByVal Optional BackColor1 As Color = Nothing,
                             ByVal Optional BackColor2 As Color = Nothing,
                             ByVal Optional GradientAngle As Integer = 0,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&Hf0dd"))
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
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal ShowCombo As Boolean,
                             ByVal ComboContent As List(Of String),
                             ByVal Optional BackColor1 As Color = Nothing,
                             ByVal Optional BackColor2 As Color = Nothing,
                             ByVal Optional GradientAngle As Integer = 0,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional Alignment As Style.Alignment = Style.Alignment.MiddleLeft
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&Hf0dd"))
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
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal ShowCombo As Boolean,
                             ByVal ComboContent As List(Of ComboItem),
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&HF0DD"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal ShowCombo As Boolean,
                             ByVal ComboContent As List(Of String)
                             ) As GridRow
        Try
            Dim GridCell As New GridCell
            If ShowCombo Then
                GridCell.EditorType = GetType(FragrantComboBox)
                GridCell.EditorParams = New Object() {ComboContent}
                GridCell.AllowEdit = True
                GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&Hf0dd"))
                GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
            End If
            GridRow.Cells.Add(GridCell)
        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId,  False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return GridRow
    End Function
    Public Shared Function addCell(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal ShowCombo As Boolean,
                             ByVal ComboContent As List(Of String),
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&HF0DD"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function

    Public Shared Function addCell_ChkBox(ByVal GridRow As GridRow,
                             ByVal checked As Boolean,
                             ByVal Optional visible As Boolean = True,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True,
                             ByVal Optional Tag As Object = Nothing
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = checked
        If Not IsNothing(Tag) Then
            GridCell.Tag = Tag
        End If
        GridCell.Visible = visible
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell_Image(ByVal GridRow As GridRow,
                             ByVal image As Image,
                             ByVal Optional Tag As String = "",
                             ByVal Optional visible As Boolean = True,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.CellStyles.Default.Image = image
        GridCell.Tag = Tag
        GridCell.Value = ""
        GridCell.Visible = visible
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell_SwBtn(ByVal GridRow As GridRow,
                             ByVal SwitchValue As Boolean,
                             ByVal TextOn As String,
                             ByVal TextOff As String,
                             ByVal Optional visible As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = SwitchValue
        GridCell.Visible = visible
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell_Text(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal Optional Tag As String = "",
                             ByVal Optional Bold As Boolean = False,
                             ByVal Optional TextColor As Color = Nothing,
                             ByVal Optional BackColor1 As Color = Nothing,
                             ByVal Optional BackColor2 As Color = Nothing,
                             ByVal Optional GradientAngle As Integer = 0,
                             ByVal Optional visible As Boolean = True,
                             ByVal Optional image As Image = Nothing,
                             ByVal Optional AllowEdit As Boolean = False,
                             ByVal Optional IsNumericInput As Boolean = False,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
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
        End If
        If Bold Then
            Dim font As New Font("Microsoft Sans Serif", 8.25, FontStyle.Bold)
            GridCell.CellStyles.Default.Font = font
        End If
        GridCell.AllowEdit = AllowEdit
        If AllowEdit Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&HF040"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
            'GridCell.InfoImage = My.Resources.info
            'GridCell.InfoText = "Eingabe"
        End If
        If IsNumericInput Then
            GridCell.EditorType = GetType(GridIntegerInputEditControl)
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell_Text(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal symbol As Integer,
                             ByVal symbolcolor As Color,
                             ByVal Optional tag As String = "",
                             ByVal Optional Bold As Boolean = False,
                             ByVal Optional TextColor As Color = Nothing,
                             ByVal Optional BackColor As Color = Nothing,
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
          ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        GridCell.Tag = tag
        If symbol > 0 Then
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger(symbol))
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
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell_Cbb(ByVal GridRow As GridRow,
                             ByVal Text As String,
                             ByVal ShowCombo As Boolean,
                             ByVal ComboContent As List(Of ComboItem),
                             ByVal Optional IsReadOnly As Boolean = True,
                             ByVal Optional AllowSelection As Boolean = True
                             ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        If ShowCombo Then
            GridCell.EditorType = GetType(FragrantComboBox)
            GridCell.EditorParams = New Object() {ComboContent}
            GridCell.AllowEdit = True
            GridCell.CellStyles.Default.SymbolDef.Symbol = ChrW(Helper_Convert.ConvertToInteger("&HF0DD"))
            GridCell.CellStyles.Default.SymbolDef.SymbolColor = Color.Blue
        End If
        GridCell.ReadOnly = IsReadOnly
        GridCell.AllowSelection = AllowSelection
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function
    Public Shared Function addCell_Btn(ByVal GridRow As GridRow,
                                       ByVal Text As String,
                                       ByVal Tag As String,
                                       ByVal image As Image
                                       ) As GridRow
        Dim GridCell As New GridCell
        GridCell.Value = CStr(Text)
        GridCell.Tag = Tag
        GridCell.EditorType = GetType(GridButtonXEditControl)
        GridRow.Cells.Add(GridCell)
        Return GridRow
    End Function

#Region "FragrantComboBox"
    Public Class FragrantComboBox
        Inherits GridComboBoxExEditControl
        Public Sub New(ByVal orderArray As IEnumerable(Of String))
            DataSource = orderArray
        End Sub
    End Class
#End Region

End Class
