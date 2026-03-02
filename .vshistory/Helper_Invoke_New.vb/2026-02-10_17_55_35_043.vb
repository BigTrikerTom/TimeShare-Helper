Imports System.Drawing
Imports System.Threading
Imports DevComponents.AdvTree
Imports DevComponents.DotNetBar
Imports DevComponents.DotNetBar.Controls
Imports DevComponents.Editors
Imports DevComponents.Editors.DateTimeAdv
Imports System.Windows.Forms
Imports DevComponents.DotNetBar.Validator
Imports TimeShare_Error


Public Class Helper_Invoke_New
    Private shared rb As Boolean = False
    Private shared ri As Integer = 0
#Region "AdvTree"
    ' ---------- Invoke_AdvTree_SetImage --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetImage(ByVal Target As AdvTree, ByVal node As Node, ByVal Image As Image, ByVal Optional Tag As String = "-.-")
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Image = Image
                              If Tag <> "-.-" Then
                                  node.Image.Tag = Tag
                              End If
                          End Sub)
        Else
            node.Image = Image
            If Tag <> "-.-" Then
                node.Image.Tag = Tag
            End If
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetImageTag --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetImageTag(ByVal Target As AdvTree, ByVal node As Node, ByVal Tag As String)
        If (Target.InvokeRequired) AndAlso node IsNot Nothing AndAlso node.Image IsNot Nothing Then
            Target.Invoke(Sub()
                              node.Image.Tag = Tag
                          End Sub)
        ElseIf node IsNot Nothing AndAlso node.Image IsNot Nothing Then
            node.Image.Tag = Tag
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetNodeText --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetNodeText(ByVal Target As AdvTree, ByVal node As Node, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Text = Text
                          End Sub)
        Else
            node.Text = Text
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetNodeDataKey --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetNodeDataKey(ByVal Target As AdvTree, ByVal node As Node, ByVal DataKeyValue As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.DataKey = CInt(DataKeyValue)
                          End Sub)
        Else
            node.DataKey = CInt(DataKeyValue)
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetNodeDataKeyString --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetNodeDataKeyString(ByVal Target As AdvTree, ByVal node As Node, ByVal DataKeyString As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.DataKeyString = DataKeyString
                          End Sub)
        Else
            node.DataKeyString = DataKeyString
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetNodeExpanded --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetNodeExpanded(ByVal Target As AdvTree, ByVal node As Node, ByVal Expanded As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Expand()
                          End Sub)
        Else
            node.Expand()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetNodeEnsureVisible --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetNodeEnsureVisible(ByVal Target As AdvTree, ByVal node As Node)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.EnsureVisible()
                          End Sub)
        Else
            node.EnsureVisible()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SetNodeExpand --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SetNodeExpand(ByVal Target As AdvTree, ByVal node As Node)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Expand()
                          End Sub)
        Else
            node.Expand()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_SuspendLayout --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_SuspendLayout(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SuspendLayout()
                          End Sub)
        Else
            Target.SuspendLayout()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_ResumeLayout --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_ResumeLayout(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ResumeLayout(True)
                          End Sub)
        Else
            Target.ResumeLayout(True)
        End If
    End Sub
    ' ---------- Invoke_AdvTree_NodesClear --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_NodesClear(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Nodes.Clear()
                          End Sub)
        Else
            Target.Nodes.Clear()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_BeginUpdate --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_BeginUpdate(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BeginUpdate()
                          End Sub)
        Else
            Target.BeginUpdate()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_EndUpdate --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_EndUpdate(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.EndUpdate()
                          End Sub)
        Else
            Target.EndUpdate()
        End If
    End Sub
    ' ---------- Invoke_AdvTree_ParentNode_Nodes_Add --------------------------------------------------
    Public Shared Function Invoke_AdvTree_ParentNode_Nodes_Add(ByVal Target As AdvTree, ByVal ParentNode As Node, ByVal NewNode As Node) As Integer
        If Target.InvokeRequired Then
            Return CInt(Target.Invoke(Function() As Integer
                                          ri = ParentNode.Nodes.Add(NewNode)
                                          Return ri
                                      End Function
                        ))
        Else
            ri = ParentNode.Nodes.Add(NewNode)
        End If
        Return ri
    End Function
    ' ---------- Invoke_AdvTree_Nodes_Add --------------------------------------------------
    Public Shared Function Invoke_AdvTree_Nodes_Add(ByVal Target As AdvTree, ByVal NewNode As Node) As Integer
        If Target.InvokeRequired Then
            Return CInt(Target.Invoke(Function() As Integer
                                          ri = Target.Nodes.Add(NewNode)
                                          Return ri
                                      End Function
                        ))
        Else
            ri = Target.Nodes.Add(NewNode)
        End If
        Return ri
    End Function
    ' ---------- Invoke_AdvTree_Nodes_AddRange --------------------------------------------------
    Public Shared Sub Invoke_AdvTree_Nodes_AddRange(ByVal Target As AdvTree, ByVal NewNodeList As List(Of Node))
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Nodes.AddRange(NewNodeList.ToArray)
                          End Sub
                        )
        Else
            Target.Nodes.AddRange(NewNodeList.ToArray)
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSuspendLayout --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSuspendLayout(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SuspendLayout()
                          End Sub)
        Else
            Target.SuspendLayout()
        End If
    End Sub
    ' ---------- Invoke_AdvTreeResumeLayout --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeResumeLayout(ByVal Target As AdvTree)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ResumeLayout(True)
                          End Sub)
        Else
            Target.ResumeLayout(True)
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetImageTag --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetImageTag(ByVal Target As AdvTree, ByVal node As Node, ByVal Tag As String)
        If Target.InvokeRequired AndAlso node IsNot Nothing AndAlso node.Image IsNot Nothing Then
            Target.Invoke(Sub()
                              node.Image.Tag = Tag
                          End Sub)
        ElseIf node IsNot Nothing AndAlso node.Image IsNot Nothing Then
            node.Image.Tag = Tag
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetNodeDataKey --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetNodeDataKey(ByVal Target As AdvTree, ByVal node As Node, ByVal DataKeyValue As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.DataKey = CInt(DataKeyValue)
                          End Sub)
        Else
            node.DataKey = CInt(DataKeyValue)
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetNodeDataKeyString --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetNodeDataKeyString(ByVal Target As AdvTree, ByVal node As Node, ByVal DataKeyString As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.DataKeyString = DataKeyString
                          End Sub)
        Else
            node.DataKeyString = DataKeyString
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetNodeText --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetNodeText(ByVal Target As AdvTree, ByVal node As Node, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Text = Text
                          End Sub)
        Else
            node.Text = Text
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetNodeExpanded --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetNodeExpanded(ByVal Target As AdvTree, ByVal node As Node, ByVal Expanded As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Expand()
                          End Sub)
        Else
            node.Expand()
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetNodeEnsureVisible --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetNodeEnsureVisible(ByVal Target As AdvTree, ByVal node As Node)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.EnsureVisible()
                          End Sub)
        Else
            node.EnsureVisible()
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetNodeExpand --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetNodeExpand(ByVal Target As AdvTree, ByVal node As Node)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Expand()
                          End Sub)
        Else
            node.Expand()
        End If
    End Sub
    ' ---------- Invoke_AdvTreeSetImage --------------------------------------------------
    Public Shared Sub Invoke_AdvTreeSetImage(ByVal Target As AdvTree, ByVal node As Node, ByVal Image As Image, ByVal Optional Tag As String = "-.-")
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              node.Image = Image
                              If Tag <> "-.-" Then
                                  node.Image.Tag = Tag
                              End If
                          End Sub)
        Else
            node.Image = Image
            If Tag <> "-.-" Then
                node.Image.Tag = Tag
            End If
        End If
    End Sub

#End Region

#Region "Bar LabelItem"
    ' ---------- Invoke_Bar_LabelItem_BackColor --------------------------------------------------
    Public Shared Sub Invoke_Bar_LabelItem_BackColor(ByVal Target As Bar, ByVal lbli As LabelItem, BackColor As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              lbli.BackColor = BackColor
                          End Sub)
        Else
            lbli.BackColor = BackColor
        End If
    End Sub
    ' ---------- Invoke_Bar_LabelItem_Image --------------------------------------------------
    Public Shared Sub Invoke_Bar_LabelItem_Image(ByVal Target As Bar, ByVal tssl As LabelItem, ByVal Parameter2 As Image)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              tssl.Image = Parameter2
                          End Sub)
        Else
            tssl.Image = Parameter2
        End If
    End Sub
    ' ---------- Invoke_Bar_LabelItem_Text --------------------------------------------------
    Public Shared Sub Invoke_Bar_LabelItem_Text(ByVal Target As Bar, ByVal tssl As LabelItem, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              tssl.Text = Text
                          End Sub)
        Else
            tssl.Text = Text
        End If
    End Sub
    ' ---------- Invoke_Bar_LabelItem_Visible --------------------------------------------------
    Public Shared Sub Invoke_Bar_LabelItem_Visible(ByVal Target As Bar, ByVal tssl As LabelItem, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              tssl.Visible = Visible
                          End Sub)
        Else
            tssl.Visible = Visible
        End If
    End Sub
#End Region
#Region "ButtonItem"
    ' ---------- Invoke_ButtonItem_Enable --------------------------------------------------
    Public Shared Sub Invoke_ButtonItem_Enabled(ByVal Target As ButtonItem, ByVal Enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Enabled = Enabled
                          End Sub)
        Else
            Target.Enabled = Enabled
        End If
    End Sub
    ' ---------- Invoke_ButtonItem_ForeColor --------------------------------------------------
    Public Shared Sub Invoke_ButtonItem_ForeColor(ByVal Target As ButtonItem, ForeColor As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ForeColor = ForeColor
                          End Sub)
        Else
            Target.ForeColor = ForeColor
        End If
    End Sub
    ' ---------- Invoke_ButtonItem_Visible --------------------------------------------------
    Public Shared Sub Invoke_ButtonItem_Visible(ByVal Target As ButtonItem, ByVal Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_ButtonItem_Symbol --------------------------------------------------
    Public Shared Sub Invoke_ButtonItem_Symbol(ByVal Target As ButtonItem, ByVal Symbol As String, Optional ByVal symbolcolor As Color = Nothing, Optional ByVal symbolSize As Integer = 0, Optional ByVal symbolset As eSymbolSet = eSymbolSet.Awesome)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Symbol = Symbol
                              If Not symbolcolor = Nothing Then
                                  Target.SymbolColor = symbolcolor
                              End If
                              If symbolSize > 0 Then
                                  Target.SymbolSize = symbolSize
                              End If
                              Target.SymbolSet = symbolset
                          End Sub)
        Else
            Target.Symbol = Symbol
            If Not symbolcolor = Nothing Then
                Target.SymbolColor = symbolcolor
            End If
            If symbolSize > 0 Then
                Target.SymbolSize = symbolSize
            End If
            Target.SymbolSet = symbolset
        End If
    End Sub
    ' ---------- Invoke_ButtonItem_Tag --------------------------------------------------
    Public Shared Sub Invoke_ButtonItem_Tag(ByVal Target As ButtonItem, ByVal Tag As Object)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Tag = Tag
                          End Sub)
        Else
            Target.Tag = Tag
        End If
    End Sub
#End Region
#Region "ButtonX"
    ' ---------- Invoke_ButtonX_BackColor --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_BackColor(ByVal Target As ButtonX, ByVal BackColor As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BackColor = BackColor
                          End Sub)
        Else
            Target.BackColor = BackColor
        End If
    End Sub
    ' ---------- Invoke_ButtonX_ColorTable --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_ColorTable(ByVal Target As ButtonX, ByVal ColorTable As eButtonColor)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ColorTable = ColorTable
                          End Sub)
        Else
            Target.ColorTable = ColorTable
        End If
    End Sub
    ' ---------- Invoke_ButtonX_PerformClick --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_PerformClick(ByVal Target As ButtonX)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.PerformClick()
                          End Sub)
        Else
            Target.PerformClick()
        End If
    End Sub
    ' ---------- Invoke_ButtonX_Image --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_Image(ByVal Target As ButtonX, ByVal image As Image)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Image = image
                          End Sub)
        Else
            Target.Image = image
        End If
    End Sub
    ' ---------- Invoke_ButtonX_Enabled --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_Enabled(ByVal Target As ButtonX, ByVal Enabled As Boolean)
        If TypeOf Target Is ButtonX Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.Enabled = Enabled
                              End Sub)
            Else
                Target.Enabled = Enabled
            End If
        End If
    End Sub
    ' ---------- Invoke_ButtonX_Visible --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_Visible(ByVal Target As ButtonX, ByVal Visible As Boolean)
        If TypeOf Target Is ButtonX Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.Visible = Visible
                              End Sub)
            Else
                Target.Visible = Visible
            End If
        End If
    End Sub
    ' ---------- Invoke_ButtonX_Tag --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_Tag(ByVal Target As ButtonX, ByVal Tag As String)
        If TypeOf Target Is ButtonX Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.Tag = Tag
                              End Sub)
            Else
                Target.Tag = Tag
            End If
        End If
    End Sub
    ' ---------- Invoke_ButtonX_Enabled_Special --------------------------------------------------
    Public Shared Sub Invoke_ButtonX_Enabled_Special(ByVal Target As ButtonX, ByVal Enabled As Boolean, ByVal BackColor1 As Color, ByVal BackColor2 As Color, Optional ByVal BorderColor As Color = Nothing)
        Dim StrikeOut As Boolean = Not Enabled
        If BorderColor = Nothing Then
            BorderColor = BackColor1
        End If
        If TypeOf Target Is ButtonX Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  If StrikeOut Then
                                      Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style Or FontStyle.Strikeout)
                                  Else
                                      Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style And Not FontStyle.Strikeout)
                                  End If
                                  CType(Target.Parent, PanelEx).Style.BackColor1.Color = BackColor1
                                  CType(Target.Parent, PanelEx).Style.BackColor1.ColorSchemePart = eColorSchemePart.Custom
                                  CType(Target.Parent, PanelEx).Style.BackColor2.Color = BackColor2
                                  CType(Target.Parent, PanelEx).Style.BackColor2.ColorSchemePart = eColorSchemePart.Custom
                                  CType(Target.Parent, PanelEx).Style.BorderColor.Color = BorderColor
                                  CType(Target.Parent, PanelEx).Style.BorderColor.ColorSchemePart = eColorSchemePart.Custom
                                  Target.Enabled = Enabled
                              End Sub)
            Else
                If StrikeOut Then
                    Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style Or FontStyle.Strikeout)
                Else
                    Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style And Not FontStyle.Strikeout)
                End If
                CType(Target.Parent, PanelEx).Style.BackColor1.Color = BackColor1
                CType(Target.Parent, PanelEx).Style.BackColor1.ColorSchemePart = eColorSchemePart.Custom
                CType(Target.Parent, PanelEx).Style.BackColor2.Color = BackColor2
                CType(Target.Parent, PanelEx).Style.BackColor2.ColorSchemePart = eColorSchemePart.Custom
                CType(Target.Parent, PanelEx).Style.BorderColor.Color = BorderColor
                CType(Target.Parent, PanelEx).Style.BorderColor.ColorSchemePart = eColorSchemePart.Custom
                Target.Enabled = Enabled
            End If
        End If
    End Sub
#End Region
#Region "CheckBoxX"
    ' ---------- Invoke_CheckBoxX_Checked --------------------------------------------------
    Public Shared Sub Invoke_CheckBoxX_Checked(ByVal Target As CheckBoxX, ByVal value As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Checked = value
                          End Sub)
        Else
            Target.Checked = value
        End If
    End Sub
#End Region
#Region "CheckedListBox"
    ' ---------- Invoke_ListBoxAdv_ItemsClear --------------------------------------------------
    Public Shared Sub Invoke_CheckedListBox_ItemsClear(ByVal Target As CheckedListBox)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                          End Sub)
        Else
            Target.Items.Clear()
        End If
    End Sub

#End Region

#Region "CircularProgress"
    ' ---------- Invoke_CircularProgress_IsRunning --------------------------------------------------
    Public Shared Sub Invoke_CircularProgress_IsRunning(ByVal Target As CircularProgress, ByVal IsRunning As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.IsRunning = IsRunning
                          End Sub)
        Else
            Target.IsRunning = IsRunning
        End If
    End Sub
    ' ---------- Invoke_CircularProgress_Maximum --------------------------------------------------
    Public Shared Sub Invoke_CircularProgress_Maximum(ByVal Target As CircularProgress, Max As Double)
        If Not Helper.IsGuiThread() Then 'If target.InvokeRequired Then
            Dim icv_action As Action = Sub()
                                           Target.Invoke(Sub()
                                                             Target.Maximum = CInt(Max)
                                                         End Sub
                                                                )
                                       End Sub
            icv_action.BeginInvoke(AddressOf icv_action.EndInvoke, Nothing)
        Else
            Target.Maximum = CInt(Max)
        End If
    End Sub
    ' ---------- Invoke_CircularProgress_Value --------------------------------------------------
    Public Shared Sub Invoke_CircularProgress_Value(ByVal Target As CircularProgress, value As Double)
        If Not Helper.IsGuiThread() Then 'If target.InvokeRequired Then
            Dim icv_action As Action = Sub()
                                           Target.Invoke(Sub()
                                                             Target.Value = CInt(value)
                                                         End Sub
                                                                )
                                       End Sub
            icv_action.BeginInvoke(AddressOf icv_action.EndInvoke, Nothing)
        Else
            Target.Value = CInt(value)
        End If
    End Sub
    ' ---------- Invoke_CircularProgress_ProgressText --------------------------------------------------
    Public Shared Sub Invoke_CircularProgress_ProgressText(ByVal Target As CircularProgress, Text As String)
        If Not Helper.IsGuiThread() Then 'If target.InvokeRequired Then
            Dim icv_action As Action = Sub()
                                           Target.Invoke(Sub()
                                                             Target.ProgressText = Text
                                                         End Sub
                                                                )
                                       End Sub
            icv_action.BeginInvoke(AddressOf icv_action.EndInvoke, Nothing)
        Else
            Target.ProgressText = Text
        End If
    End Sub
#End Region
#Region "ColorPickerButton"
    ' ---------- Invoke_ColorPickerButton_SelectedColor --------------------------------------------------
    Public Shared Sub Invoke_ColorPickerButton_SelectedColor(ByVal Target As ColorPickerButton, ByVal Color As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedColor = Color
                          End Sub)
        Else
            Target.SelectedColor = Color
        End If
    End Sub
#End Region
#Region "ComboBoxEx"
    ' ---------- Invoke_ComboBoxEx_BeginUpdate --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_BeginUpdate(ByVal Target As ComboBoxEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Try
                                  Target.BeginUpdate()
                              Catch ex As Exception
                                  call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False, "")
                                  If Helper.IsIDE() Then Stop
                              End Try
                          End Sub)
        Else
            Target.BeginUpdate()
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_DropDownStyle --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_DropDownStyle(ByVal Target As ComboBoxEx, ByVal DropDownStyle As eDotNetBarStyle)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Style = DropDownStyle
                          End Sub)
        Else
            Target.Style = DropDownStyle
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_EndUpdate --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_EndUpdate(ByVal Target As ComboBoxEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.EndUpdate()
                          End Sub)
        Else
            Target.EndUpdate()
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_ItemsClear --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_ItemsClear(ByVal Target As ComboBoxEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                              Target.Text = ""
                          End Sub)
        Else
            Target.Items.Clear()
            Target.Text = ""
        End If
    End Sub
    '' ---------- Invoke_ComboBoxEx_Items_AddRange --------------------------------------------------
    'Public Shared Sub Invoke_ComboBoxEx_Items_Addrange(ByVal Target As ComboBoxEx, ByVal ComboItemList As List(Of ComboItem))
    '    If Target.InvokeRequired Then
    '        Target.Invoke(Sub()
    '                               Target.Items.Addrange(ComboItemList)
    '                           End Sub)
    '    Else
    '        Target.Items.Addrange(ComboItemList)
    '    End If
    'End Sub
    ' ---------- Invoke_ComboBoxEx_Items_Add --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_Items_Add(ByVal Target As ComboBoxEx, ByVal comboitem As ComboItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(comboitem)
                          End Sub)
        Else
            Target.Items.Add(comboitem)
        End If
    End Sub
    Public Shared Sub Invoke_ComboBoxEx_Items_Add(ByVal Target As ComboBoxEx, ByVal Text As String, ByVal Optional Tag As String = "")
        Dim Item2Insert As New ComboItem
        Item2Insert.Text = Text
        Item2Insert.Tag = Tag
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(Item2Insert)
                          End Sub)
        Else
            Target.Items.Add(Item2Insert)
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_Items_AddRange --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_Items_AddRange(ByVal Target As ComboBoxEx, ByVal Liste As List(Of String), ByVal Optional Text As String = "")
        Dim cbi As New ComboItem
        cbi.Text = Text
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.AddRange(Liste.ToArray)
                          End Sub)
        Else
            Target.Items.AddRange(Liste.ToArray)
        End If
    End Sub
    Public Shared Sub Invoke_ComboBoxEx_Items_AddRange(ByVal Target As ComboBoxEx, ByVal comboitemList As List(Of ComboItem), Optional ByVal SelectFirstItem As Boolean = False)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.AddRange(comboitemList.ToArray)
                              If SelectFirstItem AndAlso Target.Items.Count > 0 Then
                                  Target.SelectedIndex = 1
                              End If
                          End Sub)
        Else
            Target.Items.AddRange(comboitemList.ToArray)
            If SelectFirstItem AndAlso Target.Items.Count > 0 Then
                Target.SelectedIndex = 0
            End If
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_Tag --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_Tag(ByVal Target As ComboBoxEx, ByVal Tag As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Tag = Tag
                          End Sub)
        Else
            Target.Tag = Tag
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_Text --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_Text(ByVal Target As ComboBoxEx, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub)
        Else
            Target.Text = Text
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_Items_Check --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_Items_Check(ByVal Target As ComboBoxEx, ByVal comboitem As ComboItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(comboitem)
                          End Sub)
        Else
            Target.Items.Add(comboitem)
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_SelectedIndex --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_SelectedIndex(ByVal Target As ComboBoxEx, ByVal SelectedIndex As Integer)
        If Target.Items.Count > 0 Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.SelectedIndex = SelectedIndex
                              End Sub)
            Else
                Target.SelectedIndex = SelectedIndex
            End If
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_SelectByString --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_SelectByString(ByVal Target As ComboBoxEx, ByVal FindString As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedIndex = Target.FindString(FindString)
                          End Sub)
        Else
            Target.SelectedIndex = Target.FindString(FindString)
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_SelectedItem --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_SelectedItem(ByVal Target As ComboBoxEx, ByVal comboitem As ComboItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedItem = comboitem
                          End Sub)
        Else
            Target.SelectedItem = comboitem
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_SelectedText --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_SelectedText(ByVal Target As ComboBoxEx, ByVal SelectedText As String)
        If Target.Items.Count > 0 Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Dim cbi As Integer = Target.FindStringExact(SelectedText)
                                  Target.SelectedIndex = cbi
                                  'Target.SelectedText = SelectedText
                              End Sub)
            Else
                Dim cbi As Integer = Target.FindStringExact(SelectedText)
                Target.SelectedIndex = cbi
                'Target.SelectedText = SelectedText
            End If
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_SuspendLayout --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_SuspendLayout(ByVal Target As ComboBoxEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SuspendLayout()
                          End Sub)
        Else
            Target.ResumeLayout(True)
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_ResumeLayout --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_ResumeLayout(ByVal Target As ComboBoxEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ResumeLayout(True)
                          End Sub)
        Else
            Target.ResumeLayout(True)
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_by_Text --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_by_Text(ByVal Target As ComboBoxEx, ByVal SuchText As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedIndex = Target.FindStringExact(SuchText)

                          End Sub)
        Else
            Target.SelectedIndex = Target.FindStringExact(SuchText)
        End If
    End Sub
    ' ---------- Invoke_ComboBoxEx_ResetText --------------------------------------------------
    Public Shared Sub Invoke_ComboBoxEx_ResetText(ByVal Target As ComboBoxEx)
        If Target.Items.Count > 0 Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.ResetText()
                              End Sub)
            Else
                Target.ResetText()
            End If
        End If
    End Sub
    ' ---------- Invoke_Get_ComboBoxEx_SelectedItem --------------------------------------------------
    Public Shared Function Invoke_Get_ComboBoxEx_SelectedItem(ByVal Target As ComboBoxEx) As ComboItem
        If Target.InvokeRequired Then
            Return CType(Target.Invoke(Function()
                                           Return CType(Target.SelectedItem, ComboItem)
                                       End Function), ComboItem)
        Else
            Return CType(Target.SelectedItem, ComboItem)
        End If
    End Function
    ' ---------- Invoke_Get_ComboBoxEx_SelectedItemValue --------------------------------------------------
    Public Shared Function Invoke_Get_ComboBoxEx_SelectedItemValue(ByVal Target As ComboBoxEx) As String
        If Target.InvokeRequired Then
            Return CType(Target.Invoke(Function()
                                           Return CType(Target.SelectedItem, ComboItem).Value
                                       End Function), String)
        Else
            Return CType(CType(Target.SelectedItem, ComboItem).Value, String)
        End If
    End Function

#End Region
#Region "ContextMenu"
    ' ---------- Invoke_ContextMenu_Strip_Label --------------------------------------------------
    Public Shared Sub Invoke_ContextMenu_Strip_Label(ByVal Target As Label, ByVal Text As ContextMenuStrip)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ContextMenuStrip = Text
                          End Sub)
        Else
            Target.ContextMenuStrip = Text
        End If
    End Sub
    ' ---------- Invoke_ContextMenu_Strip_LabelX --------------------------------------------------
    Public Shared Sub Invoke_ContextMenu_Strip_LabelX(ByVal Target As LabelX, ByVal Text As ContextMenuStrip)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ContextMenuStrip = Text
                          End Sub)
        Else
            Target.ContextMenuStrip = Text
        End If
    End Sub
#End Region
#Region "ContextMenuStrip"
    ' ---------- Invoke_ContextMenuStrip_Label --------------------------------------------------
    Public Shared Sub Invoke_ContextMenuStrip_Label(ByVal Target As Label, ByVal Text As ContextMenuStrip)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ContextMenuStrip = Text
                          End Sub)
        Else
            Target.ContextMenuStrip = Text
        End If
    End Sub
#End Region
#Region "Control"
    ' ---------- Invoke_Control_BackColor --------------------------------------------------
    Public Shared Sub Invoke_Control_BackColor(ByVal Target As Control, ByVal BackColor As Color)
        If Target.InvokeRequired Then
            'Target.BeginInvoke
            Target.Invoke(Sub()
                              Target.BackColor = BackColor
                          End Sub)
        Else
            Target.BackColor = BackColor
        End If
    End Sub
    ' ---------- Invoke_Control_BringToFront --------------------------------------------------
    Public Shared Sub Invoke_Control_BringToFront(ByVal Target As Control)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BringToFront()
                          End Sub)
        Else
            Target.BringToFront()
        End If
    End Sub
    ' ---------- Invoke_Control_ContextMenuStrip --------------------------------------------------
    Public Shared Sub Invoke_Control_ContextMenuStrip(ByVal Target As Control, ByVal ContextMenuStrip As ContextMenuStrip)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ContextMenuStrip = ContextMenuStrip
                          End Sub)
        Else
            Target.ContextMenuStrip = ContextMenuStrip
        End If
    End Sub
    ' ---------- Invoke_Control_Enabled --------------------------------------------------
    Public Shared Sub Invoke_Control_Enabled(ByVal Target As Control, ByVal Enabled As Boolean)
        If TypeOf Target Is Control Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.Enabled = Enabled
                              End Sub)
            Else
                Target.Enabled = Enabled
            End If
        End If
    End Sub
    ' ReSharper disable once UnusedMember.Local
    Private Shared Sub Test(ByVal Target As Control, ByVal Enabled As Boolean)
        Target.Enabled = Enabled

    End Sub
    ' ---------- Invoke_Control_ForeColor --------------------------------------------------
    Public Shared Sub Invoke_Control_ForeColor(ByVal Target As Control, ByVal ForeColor As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ForeColor = ForeColor
                          End Sub)
        Else
            Target.ForeColor = ForeColor
        End If
    End Sub
    ' ---------- Invoke_Control_Left --------------------------------------------------
    Public Shared Sub Invoke_Control_Left(ByVal Target As Control, ByVal left As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Left = left
                          End Sub)
        Else
            Target.Left = left
        End If
    End Sub
    ' ---------- Invoke_Control_SetFont --------------------------------------------------
    Public Shared Sub Invoke_Control_SetFont(ByVal Target As Control, ByVal DefFont As Font)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Font = DefFont
                          End Sub)
        Else
            Target.Font = DefFont
        End If
    End Sub
    ' ---------- Invoke_Control_SetFontStrikeOut --------------------------------------------------
    Public Shared Sub Invoke_Control_SetFontStrikeOut(ByVal Target As Control, ByVal StrikeOut As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              If StrikeOut Then
                                  Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style Or FontStyle.Strikeout)
                              Else
                                  Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style And Not FontStyle.Strikeout)
                              End If
                          End Sub)
        Else
            If StrikeOut Then
                Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style Or FontStyle.Strikeout)
            Else
                Target.Font = New Font(Target.Font.FontFamily, Target.Font.Size, Target.Font.Style And Not FontStyle.Strikeout)
            End If
        End If
    End Sub
    ' ---------- Invoke_Control_Tag --------------------------------------------------
    Public Shared Sub Invoke_Control_Tag(ByVal Target As Control, ByVal Tag As Object)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Tag = Tag
                          End Sub)
        Else
            Target.Tag = Tag
        End If
    End Sub
    ' ---------- Invoke_Control_Text --------------------------------------------------
    Public Shared Sub Invoke_Control_Text(ByVal Target As Control, ByVal Text As String, Optional ByVal Tag As Object = Nothing)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                              If Tag IsNot Nothing Then
                                  Target.Tag = Tag
                              End If
                          End Sub)
        Else
            Target.Text = Text
            If Tag IsNot Nothing Then
                Target.Tag = Tag
            End If
        End If
    End Sub
    ' ---------- Invoke_Control_Top --------------------------------------------------
    Public Shared Sub Invoke_Control_Top(ByVal Target As Control, ByVal top As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Top = top
                          End Sub)
        Else
            Target.Top = top
        End If
    End Sub
    ' ---------- Invoke_Control_Visible --------------------------------------------------
    Public Shared Sub Invoke_Control_Visible(ByVal Target As Control, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    Public Shared Function Invoke_Control_Get_Text(ByVal Target As Control) As String
        Dim retval As String = ""
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              retval = Target.Text
                          End Sub)
        Else
            retval = Target.Text
        End If
        Return retval
    End Function

#End Region
#Region "DataGridView"
    ' ---------- Invoke_DataGridViewRows_Clear --------------------------------------------------
    Public Shared Sub Invoke_DataGridViewRows_Clear(Target As DataGridView)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Rows.Clear()
                          End Sub)
        Else
            Target.Rows.Clear()
        End If
    End Sub
    ' ---------- Invoke_DataGridViewRow_Cells_ForeColor --------------------------------------------------
    Public Shared Sub Invoke_DataGridViewRow_Cells_ForeColor(ByVal Target As DataGridView, ByVal dgvr As DataGridViewRow, ByVal Cell As Integer, ByVal Value As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              dgvr.Cells(Cell).Style.ForeColor = Value
                          End Sub)
        Else
            dgvr.Cells(Cell).Style.ForeColor = Value
        End If
    End Sub
    ' ---------- Invoke_DataGridViewRow_Cells_ToolTip --------------------------------------------------
    Public Shared Sub Invoke_DataGridViewRow_Cells_ToolTip(ByVal Target As DataGridView, ByVal dgvr As DataGridViewRow, ByVal Cell As Integer, ByVal Value As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              dgvr.Cells(Cell).ToolTipText = Value
                          End Sub)
        Else
            dgvr.Cells(Cell).ToolTipText = Value
        End If
    End Sub
    ' ---------- Invoke_DataGridViewRow_Cells_Value --------------------------------------------------
    Public Shared Sub Invoke_DataGridViewRow_Cells_Value(ByVal Target As DataGridView, ByVal dgvr As DataGridViewRow, ByVal Cell As Integer, ByVal Value As Object)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              dgvr.Cells(Cell).Value = Value
                          End Sub)
        Else
            dgvr.Cells(Cell).Value = Value
        End If
    End Sub
    ' ---------- Invoke_DataGridView_ClearSelection --------------------------------------------------
    Public Shared Sub Invoke_DataGridView_ClearSelection(ByVal Target As DataGridView)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ClearSelection()
                          End Sub)
        Else
            Target.ClearSelection()
        End If
    End Sub
    ' ---------- Invoke_DataGridView_Columns_Visible --------------------------------------------------
    Public Shared Sub Invoke_DataGridView_Columns_Visible(ByVal Target As DataGridView, ByVal dgvc As DataGridViewColumn, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              dgvc.Visible = Visible
                          End Sub)
        Else
            dgvc.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_DataGridView_CurrentCell --------------------------------------------------
    Public Shared Sub Invoke_DataGridView_CurrentCell(ByVal Target As DataGridView, ByVal value As DataGridViewCell)
        If Target.IsDisposed Then
            Exit Sub
        End If
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.CurrentCell = value
                          End Sub)
        Else
            Target.CurrentCell = value
        End If
    End Sub
    ' ---------- Invoke_DataGridView_Default_CellStyle_BackColor --------------------------------------------------
    Public Shared Sub Invoke_DataGridView_Default_CellStyle_BackColor(ByVal Target As DataGridView, ByVal dgvr As DataGridViewRow, ByVal Color As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              dgvr.DefaultCellStyle.BackColor = Color
                          End Sub)
        Else
            dgvr.DefaultCellStyle.BackColor = Color
        End If
    End Sub
    ' ---------- Invoke_DataGridView_FirstDisplayedScrollingRowIndex --------------------------------------------------
    Public Shared Sub Invoke_DataGridView_FirstDisplayedScrollingRowIndex(ByVal Target As DataGridView, ByVal value As Integer)
        If Target.IsDisposed Then
            Exit Sub
        End If
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.FirstDisplayedScrollingRowIndex = value
                          End Sub)
        Else
            Target.FirstDisplayedScrollingRowIndex = value
        End If
    End Sub
    ' ---------- Function Invoke_DataGridView_Row_Add --------------------------------------------------
    Public Shared Function Invoke_DataGridView_Row_Add(ByVal Target As DataGridView, ByVal ParamArray Values() As Object) As Integer
        Dim ri1 As Integer = 0
        If Target.InvokeRequired Then
            Return CInt(Target.Invoke(Function()
                                          ri1 = Target.Rows.Add(Values)
                                          Return ri1
                                      End Function
                        ))
        Else
            ri1 = Target.Rows.Add(Values)
        End If
        Return ri1
    End Function
    ' ---------- Invoke_DataGridView_Row_Cell_Tag --------------------------------------------------
    Public Shared Sub Invoke_DataGridView_Row_Cell_Tag(ByVal Target As DataGridView, ByVal Row As Integer, ByVal Cell As Integer, ByVal Value As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Rows(Row).Cells(Cell).Tag = Value
                          End Sub)
        Else
            Target.Rows(Row).Cells(Cell).Tag = Value
        End If
    End Sub
#End Region
#Region "DateTimeInput"
    ' ---------- Invoke_DateTimeInput_value --------------------------------------------------
    Public Shared Sub Invoke_DateTimeInput_value(ByVal Target As DateTimeInput, ByVal value As Date)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = value
                          End Sub)
        Else
            Target.Value = value
        End If
    End Sub
#End Region
#Region "DoubleInput"
    ' ---------- Invoke_DoubleInput_Visible --------------------------------------------------
    Public Shared Sub Invoke_DoubleInput_Visible(ByVal Target As DoubleInput, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_DoubleInput_Value --------------------------------------------------
    Public Shared Sub Invoke_DoubleInput_Value(ByVal Target As DoubleInput, ByVal Value As Double)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = Value
                          End Sub)
        Else
            Target.Value = Value
        End If
    End Sub
    ' ---------- Invoke_DoubleInput_ResetText --------------------------------------------------
    Public Shared Sub Invoke_DoubleInput_ResetText(ByVal Target As DoubleInput)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ResetText()
                          End Sub)
        Else
            Target.ResetText()
        End If
    End Sub

#End Region
#Region "Form"
    ' ---------- Invoke_Form_ControlBox --------------------------------------------------
    Public Shared Sub Invoke_Form_ControlBox(ByVal Target As Form, ByVal enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ControlBox = enabled
                          End Sub)
        Else
            Target.ControlBox = enabled
        End If
    End Sub
    ' ---------- Invoke_Form_Close --------------------------------------------------
    Public Shared Sub Invoke_Form_Close(ByVal Target As Form)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Close()
                          End Sub)
        Else
            Target.Close()
        End If
    End Sub
    ' ---------- Invoke_Form_Enabled --------------------------------------------------
    Public Shared Sub Invoke_Form_Enabled(ByVal Target As Form, ByVal Enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Enabled = Enabled
                          End Sub)
        Else
            Target.Enabled = Enabled
        End If
    End Sub
    ' ---------- Invoke_Form_Text --------------------------------------------------
    Public Shared Sub Invoke_Form_Text(ByVal Target As Form, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub)
        Else
            Target.Text = Text
        End If
    End Sub
#End Region
#Region "GroupPanel"
    ' ---------- Invoke_GroupPanel_BackColor --------------------------------------------------
    Public Shared Sub Invoke_GroupPanel_BackColor(ByVal Target As GroupPanel, ByVal BackColor1 As Color, ByVal BackColor2 As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Style.BackColor = BackColor1
                              Target.Style.BackColor2 = BackColor2
                          End Sub)
        Else
            Target.Style.BackColor = BackColor1
            Target.Style.BackColor2 = BackColor2
        End If
    End Sub
#End Region
#Region "IntegerInput"
    ' ---------- Invoke_IntegerInput_Visible --------------------------------------------------
    Public Shared Sub Invoke_IntegerInput_Visible(ByVal Target As IntegerInput, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_IntegerInput_Value --------------------------------------------------
    Public Shared Sub Invoke_IntegerInput_Value(ByVal Target As IntegerInput, ByVal Value As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = Value
                          End Sub)
        Else
            Target.Value = Value
        End If
    End Sub
    ' ---------- Invoke_IntegerInput_ResetText --------------------------------------------------
    Public Shared Sub Invoke_IntegerInput_ResetText(ByVal Target As IntegerInput)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ResetText()
                          End Sub)
        Else
            Target.ResetText()
        End If
    End Sub

#End Region
#Region "LabelItem"
    ' ---------- Invoke_LabelItem_Image --------------------------------------------------
    Public Shared Sub Invoke_LabelItem_Image(ByVal Target As LabelItem, ByVal Image As Image, ByVal Optional Tag As String = "")
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Image = Image
                              If Tag <> "" Then
                                  Target.Tag = Tag
                              End If
                          End Sub
                              )
        Else
            Target.Image = Image
            If Tag <> "" Then
                Target.Tag = Tag
            End If
        End If
    End Sub
    ' ---------- Invoke_LabelItem_Text --------------------------------------------------
    Public Shared Sub Invoke_LabelItem_Text(ByVal Target As LabelItem, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub)
        Else
            Target.Text = Text
        End If
    End Sub
    ' ---------- Invoke_LabelItem_Visible --------------------------------------------------
    Public Shared Sub Invoke_LabelItem_Visible(ByVal Target As LabelItem, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_LabelItem_Tooltip --------------------------------------------------
    Public Shared Sub Invoke_LabelItem_Tooltip(ByVal Target As LabelItem, ByVal Tooltip As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Tooltip = Tooltip
                          End Sub
                                 )
        Else
            Target.Tooltip = Tooltip
        End If
    End Sub
#End Region
#Region "LabelX"
    ' ---------- Invoke_LabelX_Visible --------------------------------------------------
    Public Shared Sub Invoke_LabelX_Visible(ByVal Target As LabelX, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_LabelX_BackColor --------------------------------------------------
    Public Shared Sub Invoke_LabelX_BackColor(ByVal Target As LabelX, ByVal BackColor As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BackColor = BackColor
                          End Sub)
        Else
            Target.BackColor = BackColor
        End If
    End Sub
    ' ---------- Invoke_LabelX_BorderType --------------------------------------------------
    Public Shared Sub Invoke_LabelX_BorderType(ByVal Target As LabelX, ByVal BorderType As eBorderType)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BorderType = BorderType
                          End Sub)
        Else
            Target.BorderType = BorderType
        End If
    End Sub
    ' ---------- Invoke_LabelX_Symbol --------------------------------------------------
    Public Shared Sub Invoke_LabelX_Symbol(ByVal Target As LabelX, ByVal Symbol As String, Optional ByVal symbolcolor As Color = Nothing, Optional ByVal symbolSize As Integer = 0, Optional ByVal symbolset As eSymbolSet = eSymbolSet.Awesome)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Symbol = Symbol
                              If Not symbolcolor = Nothing Then
                                  Target.SymbolColor = symbolcolor
                              End If
                              If symbolSize > 0 Then
                                  Target.SymbolSize = symbolSize
                              End If
                              Target.SymbolSet = symbolset
                          End Sub)
        Else
            Target.Symbol = Symbol
            If Not symbolcolor = Nothing Then
                Target.SymbolColor = symbolcolor
            End If
            If symbolSize > 0 Then
                Target.SymbolSize = symbolSize
            End If
            Target.SymbolSet = symbolset
        End If
    End Sub
    ' ---------- Invoke_LabelX_Image --------------------------------------------------
    Public Shared Sub Invoke_LabelX_Image(ByVal Target As LabelX, ByVal Image As Image, Optional ByVal Tag As Object = Nothing)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Image = Image
                              If Not Tag Is Nothing Then
                                  Target.Tag = Tag
                              End If
                          End Sub)
        Else
            Target.Image = Image
            If Tag IsNot Nothing Then
                Target.Tag = Tag
            End If
        End If
    End Sub
#End Region
#Region "ListBox"
    ' ---------- Invoke_ListBox_Items_Add --------------------------------------------------
    Public Shared Sub Invoke_ListBox_Items_Add(ByVal Target As ListBox, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(Text)
                          End Sub)
        Else
            Target.Items.Add(Text)
        End If
    End Sub
    ' ---------- Invoke_ListBox_Items_Clear --------------------------------------------------
    Public Shared Sub Invoke_ListBox_Items_Clear(ByVal Target As ListBox)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                          End Sub)
        Else
            Target.Items.Clear()
        End If
    End Sub
#End Region
#Region "ListBoxAdv"
    ' ---------- Invoke_ListBoxAdv_Background --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_Background(ByVal Target As ListBoxAdv, ByVal Backcolor1 As Color, Optional ByVal Backcolor2 As Color = Nothing, Optional ByVal Angle As Integer = 90)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BackgroundStyle.BackColor = Backcolor1
                              If Not Backcolor2 = Nothing Then
                                  Target.BackgroundStyle.BackColor2 = Backcolor2
                              End If
                              If Not Backcolor2 = Nothing Then
                                  Target.BackgroundStyle.BackColorGradientAngle = Angle
                              End If
                          End Sub)
        Else
            Target.BackgroundStyle.BackColor = Backcolor1
            If Not Backcolor2 = Nothing Then
                Target.BackgroundStyle.BackColor2 = Backcolor2
            End If
            If Not Backcolor2 = Nothing Then
                Target.BackgroundStyle.BackColorGradientAngle = Angle
            End If
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_BeginUpdate --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_BeginUpdate(Target As ListBoxAdv)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BeginUpdate()
                          End Sub)
        Else
            Target.BeginUpdate()
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_EndUpdate --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_EndUpdate(Target As ListBoxAdv)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.EndUpdate()
                          End Sub)
        Else
            Target.EndUpdate()
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_ItemsClear --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_ItemsClear(ByVal Target As ListBoxAdv)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                          End Sub)
        Else
            Target.Items.Clear()
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_Items_Add --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_Items_Add(ByVal Target As ListBoxAdv, ByVal ListBoxItem As ListBoxItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(ListBoxItem)
                          End Sub)
        Else
            Target.Items.Add(ListBoxItem)
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_Items_Remove --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_Items_Remove(ByVal Target As ListBoxAdv, ByVal ListBoxItem As ListBoxItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(ListBoxItem)
                          End Sub)
        Else
            Target.Items.Add(ListBoxItem)
        End If
    End Sub
    Public Shared Sub Invoke_ListBoxAdv_Items_Remove(ByVal Target As ListBoxAdv, ByVal ListBoxItem As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Dim lbi As New ListBoxItem
                              Dim i As Integer = Target.FindString(ListBoxItem)
                              lbi = CType(Target.Items(i), ListBoxItem)
                              Target.Items.Remove(lbi)
                          End Sub)
        Else
            Target.Items.Add(ListBoxItem)
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_Items_Add --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_Items_Add(ByVal Target As ListBoxAdv, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(Text)
                          End Sub)
        Else
            Target.Items.Add(Text)
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_ResumeLayout --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_ResumeLayout(ByVal Target As ListBoxAdv)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ResumeLayout(True)
                          End Sub)
        Else
            Target.ResumeLayout(True)
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_SuspendLayout --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_SuspendLayout(ByVal Target As ListBoxAdv)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SuspendLayout()
                          End Sub)
        Else
            Target.ResumeLayout(True)
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_Visible --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_Visible(ByVal Target As ListBoxAdv, ByVal Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_ListBox_ListViewItem_SubItems_Text --------------------------------------------------
    Public Shared Sub Invoke_ListBox_ListViewItem_SubItems_Text(ByVal Target As ListViewEx, ByVal lvi As ListViewItem, ByVal SubItemNr As Integer, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              lvi.SubItems(SubItemNr).Text = Text
                          End Sub)
        Else
            lvi.SubItems(SubItemNr).Text = Text
        End If
    End Sub
    ' ---------- Invoke_ListBoxAdv_Items_Clear --------------------------------------------------
    Public Shared Sub Invoke_ListBoxAdv_Items_Clear(ByVal Target As ListBoxAdv)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                          End Sub)
        Else
            Target.Items.Clear()
        End If
    End Sub

#End Region
#Region "ListViewEx"
    ' ---------- Invoke_ListViewEx_BeginUpdate --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_BeginUpdate(Target As ListViewEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BeginUpdate()
                          End Sub)
        Else
            Target.BeginUpdate()
        End If
    End Sub
    ' ---------- Invoke_ListViewEx_Columns_Autowidth --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_Columns_Autowidth(ByVal Target As ListViewEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
                          End Sub)
        Else
            Target.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        End If
    End Sub
    ' ---------- Invoke_ListViewEx_Columns_Add --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_Columns_Add(ByVal Target As ListViewEx, ByVal ColHead As System.Windows.Forms.ColumnHeader)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Columns.Add(ColHead)
                          End Sub)
        Else
            Target.Columns.Add(ColHead)
        End If
    End Sub
    ' ---------- Invoke_ListViewEx_EndUpdate --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_EndUpdate(Target As ListViewEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.EndUpdate()
                          End Sub)
        Else
            Target.EndUpdate()
        End If
    End Sub
    ' ---------- Invoke_Function ListView_FindItemWithText --------------------------------------------------
    Public Shared Function Invoke_ListViewEx_FindItemWithText(ByVal Target As ListViewEx, ByVal SearchText As String) As ListViewItem
        Dim ReturnWert As New ListViewItem
        If Target.InvokeRequired Then
            If Target.Items.Count > 0 Then
                Return CType(Target.Invoke(Function()
                                               Return Target.FindItemWithText(SearchText, False, 0, False)
                                           End Function
                                       ), ListViewItem)
            End If
        Else
            ReturnWert = Target.FindItemWithText(SearchText, False, 0, False)
        End If
        Return ReturnWert
    End Function
    ' ---------- Invoke_Function ListView_Items_Add --------------------------------------------------
    Public Shared Function Invoke_ListViewEx_Items_Add(ByVal Target As ListViewEx, ByVal Item As String) As ListViewItem
        Dim lvi As New ListViewItem
        If Target.InvokeRequired Then
            Target.Invoke(Function()
                              lvi = Target.Items.Add(Item)
                              Return lvi
                          End Function
                                  )
        Else
            lvi = Target.Items.Add(Item)
        End If
        Return lvi
    End Function
    Public Shared Function Invoke_ListViewEx_Items_Add(ByVal Target As ListViewEx, ByVal ListViewItem As ListViewItem) As ListViewItem
        Dim lvi As New ListViewItem
        If Target.InvokeRequired Then
            Return CType(Target.Invoke(Function()
                                           lvi = Target.Items.Add(ListViewItem)
                                           Return lvi
                                       End Function
                                  ), ListViewItem)
        Else
            lvi = Target.Items.Add(ListViewItem)
        End If
        Return lvi
    End Function
    ' ---------- Invoke_Function ListView_Items_Find --------------------------------------------------
    Public Shared Function Invoke_ListViewEx_Items_Find(ByVal Target As ListViewEx, ByVal ItemString As String, ByVal SearchSubitems As Boolean) As ListViewItem()
        Dim lvi() As ListViewItem = Nothing
        If Target.InvokeRequired Then
            Return CType(Target.Invoke(Function()
                                           lvi = Target.Items.Find(ItemString, SearchSubitems)
                                           Return lvi
                                       End Function
                                  ), ListViewItem())
        Else
            lvi = Target.Items.Find(ItemString, SearchSubitems)
        End If
        Return lvi
    End Function
    ' ---------- Invoke_Function ListView_Items_List --------------------------------------------------
    Public Shared Function Invoke_ListViewEx_Items_List(ByVal Target As ListViewEx) As ListViewItem()
        Dim lvi As IEnumerable(Of ListViewItem) = Nothing
        If Target.InvokeRequired Then
            Return CType(Target.Invoke(Function()
                                           lvi = Target.Items.Cast(Of ListViewItem).ToArray
                                           Return lvi
                                       End Function
                                  ), ListViewItem())
        Else
            lvi = Target.Items.Cast(Of ListViewItem).ToArray
        End If
        Return lvi.ToArray
    End Function
    ' ---------- Invoke_ListView_Items_Remove --------------------------------------------------
    Public Shared Sub Invoke_ListView_Items_Remove(ByVal Target As ListViewEx, ByVal ListViewItem As ListViewItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Remove(ListViewItem)
                          End Sub)

        Else
            Target.Items.Remove(ListViewItem)
        End If
    End Sub
    ' ---------- Invoke_Function ListView_Items_AddRange --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_Items_AddRange(ByVal Target As ListViewEx, ByVal ListViewItemList As List(Of ListViewItem))
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.AddRange(ListViewItemList.ToArray)
                          End Sub
                              )
        Else
            Target.Items.AddRange(ListViewItemList.ToArray)
        End If
    End Sub
    ' ---------- Invoke_ListView_Items_Add --------------------------------------------------
    Public Shared Sub Invoke_ListView_Items_Add(ByVal Target As ListView, ByVal Item As ListViewItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Add(Item)
                          End Sub
                              )
        Else
            Target.Items.Add(Item)
        End If
    End Sub
    ' ---------- Invoke_ListViewEx_Items_Clear --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_Items_Clear(Target As ListViewEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                          End Sub)
        Else
            Target.Items.Clear()
        End If
    End Sub
    ' ---------- Invoke_Function ListView_Items_Count --------------------------------------------------
    Public Shared Function Invoke_ListViewEx_Items_Count(ByVal Target As ListViewEx) As Integer
        If Target.InvokeRequired Then
            Return CInt(Target.Invoke(Function()
                                          Return Target.Items.Count()
                                      End Function
                          ))
        Else
            ri = Target.Items.Count()
        End If
        Return ri
    End Function
    ' ---------- Invoke_ListViewEx_SubItems_Add --------------------------------------------------
    Public Shared Sub Invoke_ListViewEx_SubItems_Add(TargetCtl As ListViewEx, ByVal lvi As ListViewItem, ByVal Values As String)
        If TargetCtl.InvokeRequired Then
            TargetCtl.BeginInvoke(Sub()
                                      lvi.SubItems.Add(Values)
                                  End Sub)
        Else
            lvi.SubItems.Add(Values)
        End If
    End Sub
    ' ---------- Invoke_ListView_Columns_Width --------------------------------------------------
    Public Shared Sub Invoke_ListView_Columns_Width(ByVal Target As ListView, ByVal ColNr As Integer, ByVal Width As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Columns(ColNr).Width = Width
                          End Sub)
        Else
            Target.Columns(ColNr).Width = Width
        End If
    End Sub
    ' ---------- Function Invoke_ListView_FindItemWithText --------------------------------------------------
    Public Shared Function Invoke_ListView_FindItemWithText(ByVal Target As ListView, ByVal SearchText As String) As ListViewItem
        Dim ReturnWert As ListViewItem = Nothing
        If Target.InvokeRequired Then
            If Target.Items.Count > 0 Then
                Return CType(Target.Invoke(Function() As ListViewItem
                                               Dim lvi As New ListViewItem
                                               lvi = Target.FindItemWithText(SearchText)
                                               Return lvi
                                           End Function
                                       ), ListViewItem)
            End If
        Else
            Dim lvi As New ListViewItem
            lvi = Target.FindItemWithText(SearchText)
            Return lvi
        End If
        Return ReturnWert
    End Function
    ' ---------- Invoke_ListView_Items_Clear --------------------------------------------------
    Public Shared Sub Invoke_ListView_Items_Clear(Target As ListView)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Items.Clear()
                          End Sub)
        Else
            Target.Items.Clear()
        End If
    End Sub
    ' ---------- Function Invoke_ListView_Items_Count --------------------------------------------------
    Public Shared Function Invoke_ListView_Items_Count(ByVal Target As ListView) As Integer
        If Target.InvokeRequired Then
            Return CInt(Target.Invoke(Function()
                                          Return Target.Items.Count()
                                      End Function
                          ))
        Else
            ri = Target.Items.Count()
        End If
        Return ri
    End Function
    ' ---------- Invoke_ListView_SubItems_Add --------------------------------------------------
    Public Shared Sub Invoke_ListView_SubItems_Add(TargetCtl As ListView, ByVal lvi As ListViewItem, ByVal Values As String)
        If TargetCtl.InvokeRequired Then
            TargetCtl.BeginInvoke(Sub()
                                      lvi.SubItems.Add(Values)
                                  End Sub)
        Else
            lvi.SubItems.Add(Values)
        End If
    End Sub
    ' ---------- Invoke_ListView_BeginUpdate --------------------------------------------------
    Public Shared Sub Invoke_ListView_BeginUpdate(Target As ListViewEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BeginUpdate()
                          End Sub)
        Else
            Target.BeginUpdate()
        End If

    End Sub
    ' ---------- Invoke_ListView_EndUpdate --------------------------------------------------
    Public Shared Sub Invoke_ListView_EndUpdate(Target As ListViewEx)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.EndUpdate()
                          End Sub)
        Else
            Target.EndUpdate()
        End If

    End Sub
    ' ---------- Invoke_ListView_Item_Checked --------------------------------------------------
    Public Shared Sub Invoke_ListView_Item_Checked(Target As ListView, ByVal ListItem As ListViewItem, ByVal checked As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              ListItem.Checked = checked
                          End Sub)
        Else
            ListItem.Checked = checked
        End If

    End Sub
#End Region
#Region "Me"
    ' ---------- Invoke_Me_AcceptButton --------------------------------------------------
    Public Shared Sub Invoke_Me_AcceptButton(ByVal Target As Form, ByVal Button As ButtonX)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.AcceptButton = Button
                          End Sub)
        Else
            Target.AcceptButton = Button
        End If
    End Sub
    ' ---------- Invoke_Me_BringToFront --------------------------------------------------
    Public Shared Sub Invoke_Me_BringToFront(ByVal Target As Form)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BringToFront()
                          End Sub)
        Else
            Target.BringToFront()
        End If
    End Sub
    ' ---------- Invoke_Me_Disable --------------------------------------------------
    Public Shared Sub Invoke_Me_Enabled(ByVal Target As Form, ByVal Enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Enabled = Enabled
                          End Sub)
        Else
            Target.Enabled = Enabled
        End If
    End Sub
    ' ---------- Invoke_Me_TopMost --------------------------------------------------
    Public Shared Sub Invoke_Me_TopMost(ByVal Target As Form, ByVal OnTop As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.TopMost = OnTop
                          End Sub)
        Else
            Target.TopMost = OnTop
        End If
    End Sub
    ' ---------- Invoke_Me_Cursor --------------------------------------------------
    Public Shared Sub Invoke_Me_Cursor(ByVal Target As Form, ByVal Cursor As Cursor)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Cursor = Cursor
                          End Sub)
        Else
            Target.Cursor = Cursor
        End If
    End Sub
#End Region
#Region "MessageBoxEx"
    ' ---------- Function Invoke_MessageBoxEx_Show --------------------------------------------------
    Public Shared Function Invoke_MessageBoxEx_Show(ByVal Target As Form, ByVal Text As String, ByVal Caption As String, ByVal Buttons As MessageBoxButtons, ByVal Icon As MessageBoxIcon, DefaultButton As MessageBoxDefaultButton) As DialogResult
        Dim Result As DialogResult
        If Target.InvokeRequired Then
            Return CType(Target.Invoke(Function() As DialogResult
                                           Result = MessageBoxEx.Show(Target, Text, Caption, Buttons, Icon, DefaultButton)
                                           Return Result
                                       End Function
                          ), DialogResult)
        Else
            Result = MessageBoxEx.Show(Target, Text, Caption, Buttons, Icon, DefaultButton)
            Return Result
        End If
    End Function
    Public Shared Function Invoke_MessageBoxEx_Show(ByVal Target As Form, ByVal Text As String, ByVal Caption As String, ByVal Buttons As MessageBoxButtons, ByVal Icon As MessageBoxIcon) As DialogResult
        Dim Result As DialogResult
        If Target.InvokeRequired Then
            Target.Invoke(Function()
                              Result = MessageBoxEx.Show(Target, Text, Caption, Buttons, Icon)
                              Return Result
                          End Function
                          )
        Else
            Result = MessageBoxEx.Show(Target, Text, Caption, Buttons, Icon)
            Return Result
        End If
        Return Result
    End Function

#End Region
#Region "NumericUpDown"
    ' ---------- Invoke_NumericUpDown_Visible --------------------------------------------------
    Public Shared Sub Invoke_LabelItem_Visible(ByVal Target As NumericUpDown, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_NumericUpDown_Value --------------------------------------------------
    Public Shared Sub Invoke_NumericUpDown_Value(ByVal Target As NumericUpDown, ByVal Value As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = Value
                          End Sub)
        Else
            Target.Value = Value
        End If
    End Sub
#End Region
#Region "PanelEx"
    ' ---------- Invoke_PanelEx_BackColor --------------------------------------------------
    Public Shared Sub Invoke_PanelEx_BackColor(ByVal Target As PanelEx,
                                               ByVal BackColor1 As Color,
                                               ByVal BackColor2 As Color,
                                               Optional ByVal BorderColor As Color = Nothing
                                               )
        If BorderColor = Nothing Then
            BorderColor = BackColor1
        End If
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Style.BackColor1.Color = BackColor1
                              Target.Style.BackColor1.ColorSchemePart = eColorSchemePart.Custom
                              Target.Style.BackColor2.Color = BackColor2
                              Target.Style.BackColor2.ColorSchemePart = eColorSchemePart.Custom
                              Target.Style.BorderColor.Color = BorderColor
                              Target.Style.BorderColor.ColorSchemePart = eColorSchemePart.Custom
                          End Sub)
        Else
            Target.Style.BackColor1.Color = BackColor1
            Target.Style.BackColor1.ColorSchemePart = eColorSchemePart.Custom
            Target.Style.BackColor2.Color = BackColor2
            Target.Style.BackColor2.ColorSchemePart = eColorSchemePart.Custom
            Target.Style.BorderColor.Color = BorderColor
            Target.Style.BorderColor.ColorSchemePart = eColorSchemePart.Custom
        End If
    End Sub

#End Region
#Region "PictureBox"
    ' ---------- Invoke_PictureBox_Image --------------------------------------------------
    Public Shared Sub Invoke_PictureBox_Image(ByVal Target As PictureBox, ByVal Image As Image, Optional ByVal Tag As Object = Nothing)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Image = Image
                              If Tag IsNot Nothing Then
                                  Target.Tag = Tag
                              End If
                          End Sub)
        Else
            Target.Image = Image
            If Tag IsNot Nothing Then
                Target.Tag = Tag
            End If
        End If
    End Sub
    ' ---------- Invoke_PictureBox_Tag --------------------------------------------------
    Public Shared Sub Invoke_PictureBox_Tag(ByVal Target As PictureBox, ByVal Tag As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Tag = Tag
                          End Sub)
        Else
            Target.Tag = Tag
        End If
    End Sub
    ' ---------- Invoke_PictureBox_BackColor --------------------------------------------------
    Public Shared Sub Invoke_PictureBox_BackColor(ByVal Target As PictureBox, ByVal BackColor As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BackColor = BackColor
                          End Sub)
        Else
            Target.BackColor = BackColor
        End If
    End Sub
#End Region

#Region "ProgressBarItem"
    ' ---------- Invoke_ProgressBarItem_Text --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_Text(ByVal Target As ProgressBarItem, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub
                                 )
        Else
            Target.Text = Text
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_Value --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_Value(ByVal Target As ProgressBarItem, ByVal Value As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = Value
                          End Sub
                                 )
        Else
            Target.Value = Value
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_Maximum --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_Maximum(ByVal Target As ProgressBarItem, ByVal Max As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Maximum = Max
                          End Sub
                               )
        Else
            Target.Maximum = Max
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_Minimum --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_Minimum(ByVal Target As ProgressBarItem, ByVal Min As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Minimum = Min
                          End Sub
                               )
        Else
            Target.Minimum = Min
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_Enabled --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_Enabled(ByVal Target As ProgressBarItem, ByVal Enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Enabled = Enabled
                          End Sub
                                 )
        Else
            Target.Enabled = Enabled
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_Visible --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_Visible(ByVal Target As ProgressBarItem, Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub
                                 )
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_ShowText --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_ShowText(ByVal Target As ProgressBarItem, Show As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.TextVisible = Show
                          End Sub
                                 )
        Else
            Target.TextVisible = Show
        End If
    End Sub
    ' ---------- Invoke_ProgressBarItem_ProgressType --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarItem_ProgressType(ByVal Target As ProgressBarItem, ProgressType As eProgressItemType)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.ProgressType = ProgressType
                          End Sub
                                 )
        Else
            Target.ProgressType = ProgressType
        End If
    End Sub
#End Region
#Region "ProgressBarX"
    ' ---------- Invoke_ProgressBarX_Maximum --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarX_Maximum(ByVal Target As ProgressBarX, ByVal Max As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Maximum = Max
                          End Sub)
        Else
            Target.Maximum = Max
        End If
    End Sub
    ' ---------- Invoke_ProgressBarX_Minimum --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarX_Minimum(ByVal Target As ProgressBarX, ByVal Min As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Maximum = Min
                          End Sub)
        Else
            Target.Maximum = Min
        End If
    End Sub
    ' ---------- Invoke_ProgressBarX_Value --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarX_Value(ByVal Target As ProgressBarX, ByVal Value As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = Value
                          End Sub)
        Else
            Target.Value = Value
        End If
    End Sub
    ' ---------- Invoke_ProgressBarX_Visible --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarX_Visible(ByVal Target As ProgressBarX, ByVal Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_ProgressBarX_TextVisible --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarX_TextVisible(ByVal Target As ProgressBarX, ByVal Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.TextVisible = Visible
                          End Sub)
        Else
            Target.TextVisible = Visible
        End If
    End Sub
    ' ---------- Invoke_ProgressBarX_Text --------------------------------------------------
    Public Shared Sub Invoke_ProgressBarX_Text(ByVal Target As ProgressBarX, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub)
        Else
            Target.Text = Text
        End If
    End Sub

#End Region
    '#Region "RadioListBox"
    '    ' ---------- Invoke_RadioListBox_BeginUpdate --------------------------------------------------
    '    Public Shared Sub Invoke_RadioListBox_BeginUpdate(ByVal Target As RadioListBox)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.BeginUpdate()
    '                          End Sub)
    '        Else
    '            Target.BeginUpdate()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_RadioListBox_EndUpdate --------------------------------------------------
    '    Public Shared Sub Invoke_RadioListBox_EndUpdate(ByVal Target As RadioListBox)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.EndUpdate()
    '                          End Sub)
    '        Else
    '            Target.EndUpdate()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_RadioListBox_ItemsClear --------------------------------------------------
    '    Public Shared Sub Invoke_RadioListBox_ItemsClear(Target As RadioListBox)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.Items.Clear()
    '                          End Sub)
    '        Else
    '            Target.Items.Clear()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_RadioListBox_ItemsAdd --------------------------------------------------
    '    Public Shared Sub Invoke_RadioListBox_ItemsAdd(Target As RadioListBox, ByVal lbi As ListBoxItem)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.Items.Add(lbi)
    '                          End Sub)
    '        Else
    '            Target.Items.Add(lbi)
    '        End If
    '    End Sub

    '#End Region
#Region "Screen"
    ' ---------- Invoke_grafik_CopyFromScreen --------------------------------------------------
    Public Shared Function Invoke_grafik_CopyFromScreen(ByVal Target As Form) As Bitmap
        Dim ScreenShot As New Bitmap(Target.Width, Target.Height)
        If Target.InvokeRequired Then
            Target.Invoke(Function()
                              Using grafik As Graphics = Graphics.FromImage(ScreenShot)
                                  grafik.CopyFromScreen(Target.PointToScreen(New Point(0, 0)), New Point(0, 0), New Size(Target.Width, Target.Height))
                              End Using
                              Return ScreenShot
                          End Function)
            Return ScreenShot
        Else
            Using grafik As Graphics = Graphics.FromImage(ScreenShot)
                grafik.CopyFromScreen(Target.PointToScreen(New Point(0, 0)), New Point(0, 0), New Size(Target.Width, Target.Height))
            End Using
            Return ScreenShot
        End If

    End Function

#End Region
    '#Region "SuperGridControl"
    '    ' ---------- Invoke_SuperGridControl_BeginUpdate --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_BeginUpdate(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.BeginUpdate()
    '                          End Sub)
    '        Else
    '            Target.BeginUpdate()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Cell_BackColor --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Cell_BackColor(ByVal Target As SuperGridControl, ByVal Row As GridRow, ByVal CellIndex As Integer, ByVal BackColor1 As Color, ByVal BackColor2 As Color, ByVal GradientAngle As Integer)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Row.Cells(CellIndex).CellStyles.Default.Background.Color1 = BackColor1
    '                              Row.Cells(CellIndex).CellStyles.Default.Background.Color2 = BackColor2
    '                              Row.Cells(CellIndex).CellStyles.Default.Background.GradientAngle = GradientAngle
    '                          End Sub)
    '        Else
    '            Row.Cells(CellIndex).CellStyles.Default.Background.Color1 = BackColor1
    '            Row.Cells(CellIndex).CellStyles.Default.Background.Color2 = BackColor2
    '            Row.Cells(CellIndex).CellStyles.Default.Background.GradientAngle = GradientAngle
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Cell_ReadOnly --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Cell_ReadOnly(ByVal Target As SuperGridControl, ByVal Row As GridRow, ByVal CellIndex As Integer, ByVal IsReadOnly As Boolean)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Row.Cells(CellIndex).ReadOnly = IsReadOnly
    '                          End Sub)
    '        Else
    '            Row.Cells(CellIndex).ReadOnly = IsReadOnly
    '        End If
    '    End Sub
    '    Public Shared Sub Invoke_SuperGridControl_Cell_Value(ByVal Target As SuperGridControl, ByVal Row As GridRow, ByVal CellIndex As Integer, ByVal value As Boolean)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Row.Cells(CellIndex).Value = value
    '                          End Sub)
    '        Else
    '            Row.Cells(CellIndex).Value = value
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Cell_Value --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Cell_Value(ByVal Target As SuperGridControl, ByVal Row As GridRow, ByVal CellIndex As Integer, ByVal value As Integer)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Row.Cells(CellIndex).Value = value
    '                          End Sub)
    '        Else
    '            Row.Cells(CellIndex).Value = value
    '        End If
    '    End Sub
    '    Public Shared Sub Invoke_SuperGridControl_Cell_Value(ByVal Target As SuperGridControl, ByVal Row As GridRow, ByVal CellIndex As Integer, ByVal value As String)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Row.Cells(CellIndex).Value = value
    '                          End Sub)
    '        Else
    '            Row.Cells(CellIndex).Value = value
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Columns_visible --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Columns_visible(ByVal Target As SuperGridControl, ByVal col As Integer, ByVal visible As Boolean)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Columns(col).Visible = visible
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Columns(col).Visible = visible
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_EndUpdate --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_EndUpdate(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.EndUpdate()
    '                          End Sub)
    '        Else
    '            Target.EndUpdate()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_PrimaryGrid_Rows_Add --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_PrimaryGrid_Rows_Add(ByVal Target As SuperGridControl, NewRow As GridRow)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.Add(NewRow)
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.Add(NewRow)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_PrimaryGrid_Rows_AddRange --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_PrimaryGrid_Rows_AddRange(ByVal Target As SuperGridControl, NewRowList As List(Of GridRow))
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.AddRange(NewRowList)
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.AddRange(NewRowList)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_PrimaryGrid_Rows_Clear --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_PrimaryGrid_Rows_Clear(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.Clear()
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.Clear()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Rows_Add --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Rows_Add(ByVal Target As SuperGridControl, NewRow As GridRow)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.Add(NewRow)
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.Add(NewRow)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Rows_Clear --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Rows_Clear(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.Clear()
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.Clear()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Items_Clear --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Items_Clear(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.Clear()
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.Clear()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Row_Cell_Value --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Row_Cell_Value(ByVal Target As SuperGridControl, ByVal row As GridRow, ByVal cell As Integer, ByVal Value As String)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              row.Cells(cell).Value = Value
    '                          End Sub)
    '        Else
    '            row.Cells(cell).Value = Value
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_SuspendLayout --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_SuspendLayout(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.SuspendLayout()
    '                          End Sub)
    '        Else
    '            Target.SuspendLayout()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_ResumeLayout --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_ResumeLayout(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.ResumeLayout(True)
    '                          End Sub)
    '        Else
    '            Target.ResumeLayout(True)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_SetGroup --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_SetGroup(ByVal Target As SuperGridControl, ByVal ColName As String)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.SetGroup(Target.PrimaryGrid.Columns(ColName))
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.SetGroup(Target.PrimaryGrid.Columns(ColName))
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_ClearGroup --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_ClearGroup(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.ClearGroup()
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.ClearGroup()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_RemoveGroup --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_RemoveGroup(ByVal Target As SuperGridControl, ByVal ColName As String)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.RemoveGroup(Target.PrimaryGrid.Columns(ColName))
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.RemoveGroup(Target.PrimaryGrid.Columns(ColName))
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_AddSortCol --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_AddSortCol(ByVal Target As SuperGridControl, ByVal ColName As String, ByVal SortDirection As SortDirection)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.AddSort(Target.PrimaryGrid.Columns(ColName), SortDirection)
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.AddSort(Target.PrimaryGrid.Columns(ColName), SortDirection)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Panel_Add --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Panel_Add(ByVal Target As SuperGridControl, NewPanel As GridPanel)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Rows.Add(NewPanel)
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Rows.Add(NewPanel)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Row_Panel_Add --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Row_Panel_Add(ByVal Target As SuperGridControl, row As GridRow, NewPanel As GridPanel)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              row.Rows.Add(NewPanel)
    '                          End Sub)
    '        Else
    '            row.Rows.Add(NewPanel)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Row_Panel_EnsureVisible --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Row_Panel_EnsureVisible(ByVal Target As SuperGridControl, ByVal Panel As GridPanel, ByVal visible As Boolean)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Panel.EnsureVisible(visible)
    '                          End Sub)
    '        Else
    '            Panel.EnsureVisible(visible)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Row_CollapseAll --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Row_CollapseAll(ByVal Target As SuperGridControl, ByVal row As GridRow)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              row.CollapseAll()
    '                          End Sub)
    '        Else
    '            row.CollapseAll()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Row_EnsureVisible --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Row_EnsureVisible(ByVal Target As SuperGridControl, ByVal row As GridRow)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              row.EnsureVisible()
    '                          End Sub)
    '        Else
    '            row.EnsureVisible()
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_Select_Row --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_Select_Row(ByVal Target As SuperGridControl, ByVal row As GridRow)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.PrimaryGrid.Select(row)
    '                          End Sub)
    '        Else
    '            Target.PrimaryGrid.Select(row)
    '        End If
    '    End Sub
    '    ' ---------- Invoke_SuperGridControl_ArrangeGrid --------------------------------------------------
    '    Public Shared Sub Invoke_SuperGridControl_ArrangeGrid(ByVal Target As SuperGridControl)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.ArrangeGrid()
    '                          End Sub)
    '        Else
    '            Target.ArrangeGrid()
    '        End If
    '    End Sub

    '#End Region
#Region "SuperTabControl"
    ' ---------- Invoke_SuperTabControl_SelectedTab --------------------------------------------------
    Public Shared Sub Invoke_SuperTabControl_SelectedTab(ByVal Target As SuperTabControl, ByVal TabItem As SuperTabItem)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedTab = TabItem
                          End Sub)
        Else
            Target.SelectedTab = TabItem
        End If
    End Sub
    ' ---------- Invoke_SuperTabControl_SelectedTabIndex --------------------------------------------------
    Public Shared Sub Invoke_SuperTabControl_SelectedTabIndex(ByVal Target As SuperTabControl, ByVal TabIndex As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedTabIndex = TabIndex
                          End Sub)
        Else
            Target.SelectedTabIndex = TabIndex
        End If
    End Sub
    ' ---------- Invoke_SuperTabItem_Text --------------------------------------------------
    Public Shared Sub Invoke_SuperTabItem_Text(ByVal Target As SuperTabItem, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub)
        Else
            Target.Text = Text
        End If
    End Sub
    ' ---------- Invoke_SuperTabItem_Visible --------------------------------------------------
    Public Shared Sub Invoke_SuperTabItem_Visible(ByVal Target As SuperTabItem, ByVal Visible As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Visible = Visible
                          End Sub)
        Else
            Target.Visible = Visible
        End If
    End Sub
    ' ---------- Invoke_SuperTabStrip_SelectedTabIndex --------------------------------------------------
    Public Shared Sub Invoke_SuperTabStrip_SelectedTabIndex(ByVal Target As SuperTabStrip, ByVal TabIndex As Integer)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectedTabIndex = TabIndex
                          End Sub)
        Else
            Target.SelectedTabIndex = TabIndex
        End If
    End Sub
    ' ---------- Invoke_SuperTabItem_Symbol --------------------------------------------------
    Public Shared Sub Invoke_SuperTabItem_Symbol(ByVal Target As SuperTabItem, ByVal Symbol As Int32, Optional ByVal SymbolColor As Color = Nothing, Optional ByVal SymbolSet As eSymbolSet = eSymbolSet.Awesome)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Symbol = Convert.ToChar(Symbol)
                              If Not SymbolColor = Nothing Then
                                  Target.SymbolColor = SymbolColor
                              End If
                              Target.SymbolSet = SymbolSet
                          End Sub)
        Else
            Target.Symbol = Convert.ToChar(Symbol)
            If Not SymbolColor = Nothing Then
                Target.SymbolColor = SymbolColor
            End If
            Target.SymbolSet = SymbolSet
        End If
    End Sub
#End Region
#Region "StyleManager"
    ' ---------- Invoke_StyleManager_ManagerColorTint --------------------------------------------------
    Public Shared Sub Invoke_StyleManager_ManagerColorTint(ByVal Target As Form, ByVal StyleManager As StyleManager, ByVal ManagerColorTint As Color)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              StyleManager.ManagerColorTint = ManagerColorTint
                          End Sub)
        Else
            StyleManager.ManagerColorTint = ManagerColorTint
        End If
    End Sub
#End Region
#Region "SwitchButton"
    ' ---------- Invoke_SwitchButton_Value --------------------------------------------------
    Public Shared Sub Invoke_SwitchButton_Value(ByVal Target As SwitchButton, ByVal value As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = value
                          End Sub)
        Else
            Target.Value = value
        End If
    End Sub
#End Region
#Region "SwitchButtonItem"
    ' ---------- Invoke_SwitchButtonItem_Value --------------------------------------------------
    Public Shared Sub Invoke_SwitchButtonItem_Value(ByVal Target As SwitchButtonItem, ByVal value As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Value = value
                          End Sub)
        Else
            Target.Value = value
        End If
    End Sub
#End Region
#Region "TextBox"
    ' ---------- Invoke_TextBoxX_Focus --------------------------------------------------
    Public Shared Sub Invoke_TextBoxX_Focus(ByVal Target As TextBoxX)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Focus()
                          End Sub)
        Else
            Target.Focus()
        End If
    End Sub
    ' ---------- Invoke_TextBoxX_ReadOnly --------------------------------------------------
    Public Shared Sub Invoke_TextBoxX_ReadOnly(ByVal Target As TextBoxX, ByVal ValReadOnly As Boolean)
        If TypeOf Target Is TextBoxX Then
            If Target.InvokeRequired Then
                Target.Invoke(Sub()
                                  Target.ReadOnly = ValReadOnly
                              End Sub)
            Else
                Target.ReadOnly = ValReadOnly
            End If
        End If
    End Sub
    ' ---------- Invoke_TextBoxX_SelectAll --------------------------------------------------
    Public Shared Sub Invoke_TextBoxX_SelectAll(ByVal Target As TextBoxX)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectAll()
                          End Sub)
        Else
            Target.SelectAll()
        End If
    End Sub
    ' ---------- Invoke_TextBoxX_Text --------------------------------------------------
    Public Shared Sub Invoke_TextBoxX_Text(ByVal Target As TextBoxX, ByVal Text As String)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Text = Text
                          End Sub)
        Else
            Target.Text = Text
        End If
    End Sub
    ' ---------- Invoke_TextBox_AutoCompleteSource --------------------------------------------------
    Public Shared Sub Invoke_TextBox_AutoCompleteSource(ByVal Target As TextBox, ByVal ValAutoCompleteSource As AutoCompleteSource)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.AutoCompleteSource = AutoCompleteSource.CustomSource
                          End Sub)
        Else
            Target.AutoCompleteSource = AutoCompleteSource.CustomSource

        End If
    End Sub
    ' ---------- Invoke_TextBox_AutoCompleteCustomSource --------------------------------------------------
    Public Shared Sub Invoke_TextBox_AutoCompleteCustomSource(ByVal Target As TextBox, ByVal ValAutoCompleteCustomSource As AutoCompleteStringCollection)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.AutoCompleteCustomSource = ValAutoCompleteCustomSource
                          End Sub)
        Else
            Target.AutoCompleteCustomSource = ValAutoCompleteCustomSource
        End If
    End Sub
    ' ---------- Invoke_TextBox_AutoCompleteMode --------------------------------------------------
    Public Shared Sub Invoke_TextBox_AutoCompleteMode(ByVal Target As TextBox, ByVal ValAutoCompleteMode As AutoCompleteMode)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.AutoCompleteMode = ValAutoCompleteMode
                          End Sub)
        Else
            Target.AutoCompleteMode = ValAutoCompleteMode
        End If
    End Sub
    ' ---------- Invoke_TextBox_SelectAll --------------------------------------------------
    Public Shared Sub Invoke_TextBox_SelectAll(ByVal Target As TextBox)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.SelectAll()
                          End Sub)
        Else
            Target.SelectAll()
        End If
    End Sub
#End Region
    '#Region "WebEditor"
    '    ' ---------- Invoke_WebEditor_TextHtml --------------------------------------------------
    '    Public Shared Sub Invoke_WebEditor_TextHtml(ByVal Target As Editor, ByVal TextHtml As String, Optional ByVal Tag As Object = Nothing)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.BodyHtml = TextHtml
    '                              If  Not Tag Nothing Then
    '                                  Target.Tag = Tag
    '                              End If
    '                          End Sub)
    '        Else
    '            Target.BodyHtml = TextHtml
    '            If  Not Tag Nothing Then
    '                Target.Tag = Tag
    '            End If
    '        End If
    '    End Sub
    '    ' ---------- Invoke_WebEditor_TextText --------------------------------------------------
    '    Public Shared Sub Invoke_WebEditor_TextText(ByVal Target As Editor, ByVal TextText As String, Optional ByVal Tag As Object = Nothing)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.BodyText = TextText
    '                              If  Not Tag Nothing Then
    '                                  Target.Tag = Tag
    '                              End If
    '                          End Sub)
    '        Else
    '            Target.BodyText = TextText
    '            If  Not Tag Nothing Then
    '                Target.Tag = Tag
    '            End If
    '        End If
    '    End Sub
    '    ' ---------- Invoke_WebEditor_Tag --------------------------------------------------
    '    Public Shared Sub Invoke_WebEditor_Tag(ByVal Target As Editor, ByVal Tag As Object)
    '        If Target.InvokeRequired Then
    '            Target.Invoke(Sub()
    '                              Target.Tag = Tag
    '                          End Sub)
    '        Else
    '            Target.Tag = Tag
    '        End If
    '    End Sub

    '#End Region
#Region "Window"
    ' ---------- Invoke_Window_BringToFront --------------------------------------------------
    Public Shared Sub Invoke_Window_BringToFront(ByVal Target As Form)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.BringToFront()
                          End Sub)
        Else
            Target.BringToFront()
        End If

    End Sub
    ' ---------- Invoke_Window_Enabled --------------------------------------------------
    Public Shared Sub Invoke_Window_Enabled(ByVal Target As Form, ByVal enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Target.Enabled = enabled
                          End Sub)
        Else
            Target.Enabled = enabled
        End If

    End Sub
#End Region
#Region "Wizard"
    ' ---------- Invoke_Wizard_CancelButtonControl --------------------------------------------------
    Public Shared Sub Invoke_Wizard_CancelButtonControl(ByVal Target As Wizard, ByVal Control As Control, ByVal enabled As Boolean)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Control.Enabled = enabled
                          End Sub)
        Else
            Control.Enabled = enabled
        End If
    End Sub
#End Region
'#Region "FragrantComboBox"
'    Friend Class FragrantComboBox
'        Inherits GridComboBoxExEditControl
'        Public Sub New(ByVal orderArray As IEnumerable)
'            DataSource = orderArray
'        End Sub
'    End Class
'#End Region
#Region "Highlighter"
    ' ---------- Invoke_Highlighter_SetHighlightColor --------------------------------------------------
    Public Shared Sub Invoke_Highlighter_SetHighlightColor(ByVal Target As Form, ByVal Highlighter As Highlighter, ByVal Ctrl As Control, ByVal color As eHighlightColor)
        If Target.InvokeRequired Then
            Target.Invoke(Sub()
                              Highlighter.SetHighlightColor(Ctrl, color)
                          End Sub)
        Else
            Highlighter.SetHighlightColor(Ctrl, color)
        End If
    End Sub

#End Region


End Class
