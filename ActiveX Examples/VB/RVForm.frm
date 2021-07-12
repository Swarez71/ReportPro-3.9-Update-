VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{7610B470-BE49-11D0-877E-00609726A5CE}#3.0#0"; "RPRT300.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RVForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Basic Report View"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RpRuntimeCtl.RpRuntime oReport 
      Left            =   4320
      OleObjectBlob   =   "RVForm.frx":0000
      Top             =   480
   End
   Begin VB.CommandButton EditBtn 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   330
      Left            =   6165
      TabIndex        =   12
      Top             =   570
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog OpenDialog 
      Left            =   3840
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "rpt"
      DialogTitle     =   "Open Report"
      Filter          =   "(*.rpt)|*.rpt"
      PrinterDefault  =   0   'False
   End
   Begin VB.CommandButton CloseBtn 
      Caption         =   "Close"
      Height          =   375
      Left            =   6180
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton ExportBtn 
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4755
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton PrintBtn 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3660
      TabIndex        =   9
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton PreviewBtn 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2565
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton SetupDlgBtn 
      Caption         =   "Setup Dialog"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1155
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton PrintDlgBtn 
      Caption         =   "Print Dialog"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CheckBox EventsCheckBox 
      Caption         =   "Show Events"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton FileBtn 
      Height          =   315
      Left            =   6900
      Picture         =   "RVForm.frx":005A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   165
      Width           =   330
   End
   Begin VB.TextBox FileEdit 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   180
      Width           =   5640
   End
   Begin ComctlLib.ListView ListView 
      Height          =   3210
      Left            =   2910
      TabIndex        =   1
      Top             =   960
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   5662
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin ComctlLib.TreeView TreeView 
      Height          =   3210
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   5662
      _Version        =   327682
      LabelEdit       =   1
      Style           =   6
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.Label FileLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Report File:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   855
   End
End
Attribute VB_Name = "RVForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const NO_EDIT As Integer = 0
Private Const EO_EXP As Integer = 1
Private Const EO_EXP_LOGIC As Integer = 2
Private Const EO_SLE_EDIT As Integer = 3
Private Const EO_MLE_EDIT As Integer = 4
Private Const EO_MLE_RO As Integer = 5
Private Const RIT_REPORT As String = "R"
Private Const RIT_PRINTER As String = "P"
Private Const RIT_SECTION As String = "S"
Private Const RIT_TABLE As String = "T"
Private Const RIT_VARIABLE As String = "V"
Private Sub AddChildrenTables2TreeView(nSection As Integer, oParent As Node, cTableName As String, nNest As Integer)

    ' This routine gathers information about the children of a table.
    ' Since relationship between tables is hierarchical, this routine
    ' will be called recursively as it adds the tables to our tree view.
    
    ' Note: the child table names are returned in a comma delimited string.

    ' Parameters:
    '  nSection -   The section number we are working in.
    '  oParent -    The Node that will be the parent for each item added to
    '               the tree view.
    '  cTableName - The name of the table that we are retrieving children
    '               table information for.
    
    Dim cChildTables As String
    Dim nCount As Integer
    Dim nPos As Integer
    Dim cName As String
    Dim oTVI As Node
       
    cChildTables = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_CHILD_TABLES)
    Do While cChildTables <> ""
    
        ' parse out each of the table names
        nPos = InStr(1, cChildTables, ", ")
        If nPos < 1 Then
            cName = cChildTables
            cChildTables = ""
        Else
            cName = Left(cChildTables, nPos - 1)
            cChildTables = Right(cChildTables, Len(cChildTables) - nPos)
        End If

        ' add each of the child tables to the tree view
        Set oTVI = TreeView.Nodes.Add(oParent, tvwChild, MakeTreeViewKey(RIT_TABLE, nSection, nCount) + Format(nNest, "@@@"), cName)

        ' see if this table has child tables
        AddChildrenTables2TreeView nSection, oTVI, cName, nNest + 1
    Loop
   
End Sub

Private Sub AddItem2ListView(cKey As String, cTitle As String, cValue As String)
    
    ' This routine adds an item to the list view control
    
    Dim oItem As ListItem
    
    Set oItem = ListView.ListItems.Add(, cKey, cTitle)
    oItem.SubItems(1) = cValue

End Sub

Private Function MakeListViewKey(nEditOption As Integer, nItemId As Integer) As String

    ' This function creates a unique key for the list view items
    ' The key contains the item attribute number and the edit option
    ' for the attribute.
    MakeListViewKey = Format(nEditOption, "@") + Format(nItemId, "@@@@@@")
    
End Function

Private Sub LoadPrinterInfo()

    ' This routine retrieves printer related information from the report and
    ' loads it into the list view.  It also associates the editing options
    ' available for each attribute.

    Dim cValue As String
    Dim nValue As Integer
    
    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PRINTER_NAME)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_PRINTER_NAME), "Printer", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PRINT_JOB_TITLE)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_PRINT_JOB_TITLE), "Print Job Title", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PRINT2FILE_NAME)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_PRINT2FILE_NAME), "Print To File Name", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PRINT_CAPTION)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_PRINT_CAPTION), "Printing Dialog Caption", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PRINT_MESSAGE1)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_PRINT_MESSAGE1), "Printing Dialog Message 1", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PRINT_MESSAGE2)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_PRINT_MESSAGE2), "Printing Dialog Message 2", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PREVIEW_MODAL)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_PREVIEW_MODAL), "Modal Preview", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PREVIEW_CAPTION)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_PREVIEW_CAPTION), "Preview Caption", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_PREVIEW_NOZOOM)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_PREVIEW_NOZOOM), "Preview No Zoom", cValue

    nValue = oReport.GetReportIntAttribute(RPT_ATTR_PREVIEW_ZOOM_MODE)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_PREVIEW_ZOOM_MODE), "Preview Zoom Mode", Str(nValue)

    nValue = oReport.GetReportIntAttribute(RPT_ATTR_PREVIEW_PAGECOUNT)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_PREVIEW_PAGECOUNT), "Preview Panes", Str(nValue)

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_EXPORT_FILE_NAME)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_EXPORT_FILE_NAME), "Export File Name", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_EXPORT_CAPTION)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_EXPORT_CAPTION), "Export Dialog Caption", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_EXPORT_MESSAGE)
    AddItem2ListView MakeListViewKey(EO_SLE_EDIT, RPT_ATTR_EXPORT_MESSAGE), "Export Dialog Message", cValue

End Sub

Private Sub LoadReportInfo()

    ' This routine retrieves report related information from the report
    ' and loads it into the grid.  It also associates the editing option
    ' available for each attribute.

    Dim cValue As String
    
    cValue = oReport.GetReportStringAttribute(RPT_ATTR_REPORT_TITLE)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_REPORT_TITLE), "Title", cValue
    
    cValue = oReport.GetReportStringAttribute(RPT_ATTR_REPORT_DESCRIPTION)
    AddItem2ListView MakeListViewKey(EO_MLE_RO, RPT_ATTR_REPORT_DESCRIPTION), "Description", cValue
    
    cValue = oReport.GetReportStringAttribute(RPT_ATTR_CONNECTED)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_CONNECTED), "Connected to datasource(s)", cValue

    cValue = oReport.GetReportStringAttribute(RPT_ATTR_SUPPORT_1_OF_N)
    AddItem2ListView MakeListViewKey(NO_EDIT, RPT_ATTR_SUPPORT_1_OF_N), "Supports 1 of N", cValue

End Sub

Private Sub LoadListView(ByRef oTVI As Node)

    ' This routine is called when a tree view item is selected.  It
    ' determines what category of information we are viewing and loads
    ' the appropriate information in the list view.

    Dim cKey As String
    
    ' delete all the items in the grid
    ListView.ListItems.Clear
    
    ' get the extra information we saved on the tree view item
    cKey = Left(oTVI.Key, 1)

    ' determine what type of information we want to load
    If cKey = RIT_REPORT Then
        LoadReportInfo

    ElseIf cKey = RIT_PRINTER Then
        LoadPrinterInfo

    ElseIf cKey = RIT_SECTION Then
        ' make sure this is a valid section and not the "Sections" node
        If Val(Mid(oTVI.Key, 2, 3)) > 0 Then
            LoadSectionInfo Val(Mid(oTVI.Key, 2, 3))
        End If
        
    ElseIf cKey = RIT_TABLE Then
        ' Make sure this is a valid node and not the "Tables" node
        If Val(Mid(oTVI.Key, 5, 3)) > 0 Then
            LoadTableInfo Val(Mid(oTVI.Key, 2, 3)), oTVI.Text
        End If
    
    ElseIf cKey = RIT_VARIABLE Then
        ' Make sure this is a valid node and not the "Variables" node
        If Val(Mid(oTVI.Key, 5, 3)) > 0 Then
            LoadVariableInfo Val(Mid(oTVI.Key, 2, 3)), oTVI.Text
        End If
    End If

End Sub

Private Sub LoadTableInfo(nSection As Integer, cTableName As String)

    ' This routine retrieves table related information from the report and
    ' loads it into the list view.  It also associates the editing options
    ' available for each attribute.

    Dim cValue As String
    
    ' Each type of table has different properties depending on its technology.
    ' We use the ReportPro class name to determine which attributes to retrieve.
    
    cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_CLASSNAME)

    If cValue = "rpRDDTable" Then
        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_DRIVER)
        AddItem2ListView MakeListViewKey(NO_EDIT, TABLE_ATTR_DRIVER), "RDD", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_TABLE)
        AddItem2ListView MakeListViewKey(NO_EDIT, TABLE_ATTR_TABLE), "DBF File", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_INDEX_FILE)
        AddItem2ListView MakeListViewKey(NO_EDIT, TABLE_ATTR_INDEX_FILE), "Index File", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_INDEX_TAG)
        AddItem2ListView MakeListViewKey(NO_EDIT, TABLE_ATTR_INDEX_TAG), "Index Tag", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_SEEK_EXPRESSION)
        AddItem2ListView MakeListViewKey(EO_EXP, TABLE_ATTR_SEEK_EXPRESSION), "Seek Expression", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_WHILE_EXPRESSION)
        AddItem2ListView MakeListViewKey(EO_EXP, TABLE_ATTR_WHILE_EXPRESSION), "While Expression", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, TABLE_ATTR_FILTER_EXPRESSION)
        AddItem2ListView MakeListViewKey(EO_EXP, TABLE_ATTR_FILTER_EXPRESSION), "Table Filter", cValue

    ElseIf cValue = "rpSQLQuery" Then
        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_ODBC_SOURCE)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_ODBC_SOURCE), "ODBC Source", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_NULL_AS_DEFAULT)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_SQL_NULL_AS_DEFAULT), "Null As Default", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_USER_COLS)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_SQL_USER_COLS), "User Defined Columns", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_DISTINCT)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_SQL_DISTINCT), "Distinct", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_FROM)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_SQL_FROM), "From", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_TABLE_WHERE)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_SQL_TABLE_WHERE), "Where (relational portion)", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_FILTER_WHERE)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, SQLQUERY_ATTR_SQL_FILTER_WHERE), "Where (filter portion)", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_GROUP_BY)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, SQLQUERY_ATTR_SQL_GROUP_BY), "Group by", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_HAVING)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, SQLQUERY_ATTR_SQL_HAVING), "Having", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_UNION)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, SQLQUERY_ATTR_SQL_UNION), "Union", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_ORDERBY)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, SQLQUERY_ATTR_SQL_ORDERBY), "Order by", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLQUERY_ATTR_SQL_COL_DELIM)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLQUERY_ATTR_SQL_COL_DELIM), "Delimiter", cValue
    
    ElseIf cValue = "rpSQLTable" Then
        cValue = oReport.GetTableStringAttribute(nSection, cTableName, SQLTABLE_ATTR_TABLE)
        AddItem2ListView MakeListViewKey(NO_EDIT, SQLTABLE_ATTR_TABLE), "Table", cValue

    ElseIf cValue = "rpJasQuery" Then

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, JASQUERY_ATTR_DATABASE)
        AddItem2ListView MakeListViewKey(NO_EDIT, JASQUERY_ATTR_DATABASE), "Database", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, JASQUERY_ATTR_ENV_FILE)
        AddItem2ListView MakeListViewKey(NO_EDIT, JASQUERY_ATTR_ENV_FILE), "Environment File", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, JASQUERY_ATTR_WHERE)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, JASQUERY_ATTR_WHERE), "ODQL Where", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, JASQUERY_ATTR_PREQUERY_ODQL)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, JASQUERY_ATTR_PREQUERY_ODQL), "ODQL Pre-query statements", cValue

        cValue = oReport.GetTableStringAttribute(nSection, cTableName, JASQUERY_ATTR_POSTQUERY_ODQL)
        AddItem2ListView MakeListViewKey(EO_MLE_EDIT, JASQUERY_ATTR_POSTQUERY_ODQL), "ODQL Post-query statements", cValue

    End If
    
End Sub

Private Function MakeTreeViewKey(cID As String, nSection As Integer, nOther As Integer) As String

    ' This function creates a unique key for the tree view items
    ' The key contains the report area, the section number and a "nOther"
    ' value to make the key unique.
    MakeTreeViewKey = cID + Format(nSection, "@@@") + Format(nOther, "@@@")

End Function

Private Sub LoadTreeView()

    ' This routine loads the tree view with all the different report
    ' categories and entity names.
    
    Dim oTVI, oTVI2, oTVI3, oTVI4 As Node
    Dim nSections, nSection       As Integer
    Dim nVariables, nVariable     As Integer
    Dim cValue                    As String
    
    ' Add a category for the report attributes.
    TreeView.Nodes.Add , , MakeTreeViewKey(RIT_REPORT, 0, 0), "Report"

    ' Add a category for the printer attributes.
    TreeView.Nodes.Add , , MakeTreeViewKey(RIT_PRINTER, 0, 0), "Printer"

    ' Add a category for the section attributes and save the section
    ' treeview item so we can add children to it.
    Set oTVI = TreeView.Nodes.Add(, , MakeTreeViewKey(RIT_SECTION, 0, 0), "Sections")

    ' Get the number of sections in the report.
    nSections = oReport.GetReportIntAttribute(RPT_ATTR_SECTION_COUNT)

    ' Add a tree view node for each section.
    For nSection = 1 To nSections

        ' Add a label for each section.
        Set oTVI2 = TreeView.Nodes.Add(oTVI.Index, tvwChild, MakeTreeViewKey(RIT_SECTION, nSection, 0), "Section " + Str(nSection))

        ' Add a sub heading for the tables.
        Set oTVI3 = TreeView.Nodes.Add(oTVI2.Index, tvwChild, MakeTreeViewKey(RIT_TABLE, nSection, 0), "Tables")

        ' Get the primary table for this section
        cValue = oReport.GetSectionStringAttribute(nSection, SECTION_ATTR_PRIMARY_TABLE)

        ' If there is a table
        If Not cValue = "" Then
            
            ' Add the primary table.
            Set oTVI4 = TreeView.Nodes.Add(oTVI3.Index, tvwChild, MakeTreeViewKey(RIT_TABLE, nSection, 1), cValue)

            ' Add the children for this table.
            AddChildrenTables2TreeView nSection, oTVI4, cValue, 0
        End If

        ' Add a sub heading for the variables
        Set oTVI3 = TreeView.Nodes.Add(oTVI2.Index, tvwChild, MakeTreeViewKey(RIT_VARIABLE, nSection, 0), "Variables")

        ' Get the number of variables in the section
        nVariables = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_VARIABLE_COUNT)
        For nVariable = 1 To nVariables

            ' Get the name of each variable
            cValue = oReport.GetVariableStringAttribute(nSection, Format(nVariable), VARIABLE_ATTR_NAME)

            ' Add the variable to the tree
            TreeView.Nodes.Add oTVI3.Index, tvwChild, MakeTreeViewKey(RIT_VARIABLE, nSection, nVariable), cValue
        Next nVariable
    Next nSection

    ' Select the first item in the tree view
    Set oTVI = TreeView.Nodes.Item(1)
    oTVI.Selected = True
    
    LoadListView TreeView.SelectedItem

End Sub

Private Sub EnableControls()
    
    ' This routine enables or disables the controls depending on whether
    ' or not a report is loaded.
    
    Dim lEnable As Boolean
        
    lEnable = oReport.IsValid
    
    TreeView.Enabled = lEnable
    ListView.Enabled = lEnable
    PrintDlgBtn.Enabled = lEnable
    SetupDlgBtn.Enabled = lEnable
    PreviewBtn.Enabled = lEnable
    PrintBtn.Enabled = lEnable
    ExportBtn.Enabled = lEnable
    
    ' Only disable the edit button here if a report is not loaded.
    ' Otherwise, we'll let the list view selection control it.
    If Not lEnable Then
        EditBtn.Enabled = False
    End If
    
End Sub

Private Sub ClearEdits()
    
    FileEdit.Text = ""
    TreeView.Nodes.Clear
    ListView.ListItems.Clear

End Sub

Private Sub CloseBtn_Click()
    
    ' Just another way out of the dialog
    Unload RVForm
    
End Sub

Private Sub EditBtn_Click()

    Dim oTVItem As Node
    Dim oLVItem As ListItem
    Dim cArea As String
    Dim cName As String
    Dim cValue As String
    Dim nSection As Integer
    Dim nEditOption As Integer
    Dim nAttribute As Integer
    
    ' get the selected tree view item
    Set oTVItem = TreeView.SelectedItem
    If oTVItem = Null Then
        Exit Sub
    End If
    
    ' get the selected list view item
    Set oLVItem = ListView.SelectedItem
    If oLVItem = Null Then
        Exit Sub
    End If
        
    ' get the area of the report the attribute belongs too
    cArea = Left(oTVItem.Key, 1)
    
    ' get the section number because we may need it below
    nSection = Val(Mid(oTVItem.Key, 2, 3))
        
    ' get the type of editing to be performed for the selected item
    nEditOption = Val(Left(oLVItem.Key, 1))
    
    ' get the number of the attribute we are editing
    nAttribute = Val(Right(oLVItem.Key, 6))
    
    ' get the name of the attribute (only used for variables and tables)
    cName = oLVItem.Text
    
    If nEditOption = NO_EDIT Then
        Exit Sub
    
    ElseIf nEditOption = EO_EXP Or nEditOption = EO_EXP_LOGIC Then
        ' the expression builder must be used in the context of a section
        If nSection < 1 Then
            Exit Sub
        End If
        
        ' show the ReportPro expression builder.
        If nEditOption = EO_EXP_LOGIC Then
            cValue = oReport.ExpressionBuilder(hWnd, nSection, oLVItem.Text, oLVItem.SubItems(1), True, EB_ENFORCE_LOGIC, True, True, True)
        Else
            cValue = oReport.ExpressionBuilder(hWnd, nSection, oLVItem.Text, oLVItem.SubItems(1), True, EB_ENFORCE_NONE, True, True, True)
        End If
        
        ' update the list view
        oLVItem.SubItems(1) = cValue
        
        ' set the report attribute
        SetStringAttribute cArea, nSection, cName, nAttribute, cValue
    
    ElseIf nEditOption = EO_SLE_EDIT Or nEditOption = EO_MLE_EDIT Or nEditOption = EO_MLE_RO Then
        
        cValue = oLVItem.SubItems(1)
        
        If EditDialog.InitParams(cName, cValue, (nEditOption = EO_MLE_EDIT Or nEditOption = EO_MLE_RO), (nEditOption = EO_MLE_RO)) Then
            ' update the list view
            oLVItem.SubItems(1) = cValue
            
            ' set the report attribute
            SetStringAttribute cArea, nSection, cName, nAttribute, cValue
        End If
    End If
        
End Sub

Private Sub ExportBtn_Click()
    
    ' This call starts the export process
    oReport.ExportReport

End Sub

Private Sub FileBtn_Click()

    If oReport.IsValid And Not oReport.Close Then
        ' There are rare cases where you can end up here and the report is
        ' cannot properly shut down.  In that case, display an error message.
        MsgBox "The current report is busy and cannot be closed at this time.  Please try again later.", , "Error!"
        Exit Sub
    End If

    ' Clear the tree view and list view controls
    ClearEdits
    EnableControls
    
    ' Show the file open idalog
    OpenDialog.Action = 1
    If OpenDialog.FileName = "" Then
        ' If the user didn't select a report, bail out
        Exit Sub
    End If

    ' Load the report
    oReport.LoadReport OpenDialog.FileName

    ' Make sure everything loaded correctly (you should always check this after a load)
    If oReport.IsValid Then
        ' Even though we already have the file name, we'll do it the long
        ' way to show how to get it from the report object
        FileEdit.Text = oReport.GetReportStringAttribute(RPT_ATTR_REPORT_FILE)
        LoadTreeView
    Else
        oReport.Close
        MsgBox "An error occurred while opening the report!", , "Error!"
    End If
    
    EnableControls
    
End Sub

Private Sub Form_Load()
    
    ' Add a couple of columns to the list view
    ListView.ColumnHeaders.Add , , "Attribute"
    ListView.ColumnHeaders.Add , , "Value"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Check to make sure the report is properly closed
    If oReport.IsValid And Not oReport.Close Then
        
        ' There are rare cases where you can end up here and the report
        ' cannot properly shut down. In that case, display an error message.
        Cancel = True
        MsgBox "The report is busy and cannot be closed at this time.  Please try again later.", , "Error!"
    
    End If

End Sub

Private Sub ListView_ItemClick(ByVal Item As ComctlLib.ListItem)

    ' Enable the edit button based on the selected list view item
    EditBtn.Enabled = (oReport.IsValid And Val(Left(Item.Key, 1)) > NO_EDIT)

End Sub

Private Sub oReport_OnReportEnd()
    
    ' This event is fired when the report terminates
    
    Dim cText As String

    If EventsCheckBox.Value Then
        cText = "The OnReportEnd() event was fired.  "

        If oReport.SuccessfulCompletion Then
            cText = cText + "The report successfully completed."
        Else
            cText = cText + "The report was terminated by the user."
        End If

        MsgBox cText, , "Report Event!"
       
    End If

End Sub

Private Sub oReport_OnReportStart()
    
    ' This event is fired when the report starts
    
    If EventsCheckBox.Value Then
        MsgBox "The OnReportStart() event was fired.", , "Report Event!"
    End If

End Sub

Private Sub PreviewBtn_Click()
    
    ' This call starts the print preview process
    oReport.PreviewReport

End Sub

Private Sub PrintBtn_Click()
    
    ' This call starts the printing process
    oReport.PrintReport

End Sub

Private Sub PrintDlgBtn_Click()

    Dim oTreeViewItem As Node

    ' Show the print dialog.
    ' The method returns true of the user presses the OK button
    If oReport.ShowPrintDlg(UPDATE_PAPERINFO_PROMPTUSER) Then
        
        ' If we are displaying printer attributes, refresh the list view
        If Not (oTreeViewItem = TreeView.SelectedItem) = Null Then
            If Left(oTreeViewItem.Key, 1) = RIT_PRINTER Then
                LoadListView oTreeViewItem
            End If
        End If
    End If

End Sub

Private Sub SetupDlgBtn_Click()
    
    Dim oTreeViewItem As Node

    ' Show the print dialog.
    ' The method returns true of the user presses the OK button
    If oReport.ShowPrinterSetupDlg(UPDATE_PAPERINFO_PROMPTUSER) Then
        
        ' If we are displaying printer attributes, refresh the list view
        If Not (oTreeViewItem = TreeView.SelectedItem) = Null Then
            If Left(oTreeViewItem.Key, 1) = RIT_PRINTER Then
                LoadListView oTreeViewItem
            End If
        End If
    End If

End Sub

Private Sub TreeView_Click()

    LoadListView TreeView.SelectedItem
    
End Sub

Private Sub LoadSectionInfo(nSection As Integer)

    ' This routine retrieves section related information from the report
    ' and loads it into the list view.  It also associates the editing options
    ' available for each attribute.
    
    Dim nValue As Integer
    Dim cValue As String
    
    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_PAPER_SIZE)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_PAPER_SIZE), "Paper Size (DMPAPER_XXX Constant)", Str(nValue)

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_PAPER_WIDTH)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_PAPER_WIDTH), "Paper Width (TWIPS)", Str(nValue)

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_PAPER_LENGTH)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_PAPER_LENGTH), "Paper Length (TWIPS)", Str(nValue)

    cValue = oReport.GetSectionStringAttribute(nSection, SECTION_ATTR_LANDSCAPE)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_LANDSCAPE), "Landscape", cValue

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_PAPER_BIN)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_PAPER_BIN), "Paper Bin (DMBIN_XXX Constant)", Str(nValue)

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_LEFT_MARGIN)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_LEFT_MARGIN), "Left Margin (TWIPS)", Str(nValue)

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_TOP_MARGIN)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_TOP_MARGIN), "Top Margin (TWIPS)", Str(nValue)

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_RIGHT_MARGIN)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_RIGHT_MARGIN), "Right Margin (TWIPS)", Str(nValue)

    nValue = oReport.GetSectionIntAttribute(nSection, SECTION_ATTR_BOTTOM_MARGIN)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_BOTTOM_MARGIN), "Bottom Margin (TWIPS)", Str(nValue)

    cValue = oReport.GetSectionStringAttribute(nSection, SECTION_ATTR_FILTER_EXP)
    AddItem2ListView MakeListViewKey(EO_EXP_LOGIC, SECTION_ATTR_FILTER_EXP), "Filter", cValue

    cValue = oReport.GetSectionStringAttribute(nSection, SECTION_ATTR_SORT_ORDER_TEXT)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_SORT_ORDER_TEXT), "Sort Order", cValue

    cValue = oReport.GetSectionStringAttribute(nSection, SECTION_ATTR_SORT_ORDER_UNIQUE)
    AddItem2ListView MakeListViewKey(NO_EDIT, SECTION_ATTR_SORT_ORDER_UNIQUE), "Unique Sort Order", cValue

End Sub

Private Sub LoadVariableInfo(nSection As Integer, cVarName As String)

    ' This routine retrieves variable related information from the report and
    ' loads it into the list view.  It also associates the editing options
    ' available for each attribute.
    
    Dim cValue As String
    
    cValue = oReport.GetVariableStringAttribute(nSection, cVarName, VARIABLE_ATTR_RESET_LEVEL)
    AddItem2ListView MakeListViewKey(NO_EDIT, VARIABLE_ATTR_RESET_LEVEL), "Reset At", cValue
    
    cValue = oReport.GetVariableStringAttribute(nSection, cVarName, VARIABLE_ATTR_INIT_EXPRESSION)
    AddItem2ListView MakeListViewKey(EO_EXP, VARIABLE_ATTR_INIT_EXPRESSION), "Initialization Expression", cValue

    cValue = oReport.GetVariableStringAttribute(nSection, cVarName, VARIABLE_ATTR_UPDATE_LEVEL)
    AddItem2ListView MakeListViewKey(NO_EDIT, VARIABLE_ATTR_UPDATE_LEVEL), "Update At", cValue

    cValue = oReport.GetVariableStringAttribute(nSection, cVarName, VARIABLE_ATTR_UPDATE_EXPRESSION)
    AddItem2ListView MakeListViewKey(EO_EXP, VARIABLE_ATTR_UPDATE_EXPRESSION), "Update Expression", cValue
    
End Sub

Private Sub SetIntAttribute(cItemID As String, nSection As Integer, cName As String, nAttribute As Integer, nValue As Integer)

    ' This routine shows how to set an integer attribute in the report.
    ' We do not use this routine in this sample app, but it is included
    ' here for completeness.

    ' Determine which method to call.
    If cItemID = RIT_REPORT Or cItemID = RIT_PRINTER Then
        oReport.SetReportIntAttribute nAttribute, nValue

    ElseIf cItemID = RIT_SECTION Then
        oReport.SetSectionIntAttribute nSection, nAttribute, nValue

    ElseIf cItemID = RIT_TABLE Then
        oReport.SetTableIntAttribute nSection, cName, nAttribute, nValue

    ElseIf cItemID = RIT_VARIABLE Then
        ' No integer attributes for variables at this time.
        
    End If
    
End Sub
    
Private Sub SetStringAttribute(cArea As String, nSection As Integer, cName As String, nAttribute As Integer, cValue As String)

    ' This routine shows how to set a string attribute in the report.
    
    ' Determine which method to call.
    If cArea = RIT_REPORT Or cArea = RIT_PRINTER Then
        oReport.SetReportStringAttribute nAttribute, cValue

    ElseIf cArea = RIT_SECTION Then
        oReport.SetSectionStringAttribute nSection, nAttribute, cValue

    ElseIf cArea = RIT_TABLE Then
        oReport.SetTableStringAttribute nSection, cName, nAttribute, cValue

    ElseIf cArea = RIT_VARIABLE Then
        oReport.SetVariableStringAttribute nSection, cName, nAttribute, cValue
    End If

End Sub
