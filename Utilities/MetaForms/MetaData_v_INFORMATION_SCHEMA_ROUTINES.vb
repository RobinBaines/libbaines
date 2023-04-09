'------------------------------------------------
'Name: Module gen_MetaData_v_INFORMATION_SCHEMA_ROUTINES.vb.
'Function: 
'Copyright Robin Baines 2008. All rights reserved.
'Created 12-12-2015 0:00:00.
'Notes: 
'Modifications:ROUTINESTableAdapter
'------------------------------------------------
Imports Utilities
Imports System.Windows.Forms
Imports System.Drawing
Public Class MetaData_v_INFORMATION_SCHEMA_ROUTINES
Inherits dgColumns
Friend WithEvents dgSPECIFIC_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSPECIFIC_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSPECIFIC_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgROUTINE_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgROUTINE_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgROUTINE_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgROUTINE_TYPE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgMODULE_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgMODULE_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgMODULE_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgUDT_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgUDT_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgUDT_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgDATA_TYPE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCHARACTER_MAXIMUM_LENGTH As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCHARACTER_OCTET_LENGTH As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCOLLATION_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCOLLATION_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCOLLATION_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCHARACTER_SET_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCHARACTER_SET_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCHARACTER_SET_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgNUMERIC_PRECISION As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgNUMERIC_PRECISION_RADIX As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgNUMERIC_SCALE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgDATETIME_PRECISION As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgINTERVAL_TYPE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgINTERVAL_PRECISION As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgTYPE_UDT_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgTYPE_UDT_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgTYPE_UDT_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSCOPE_CATALOG As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSCOPE_SCHEMA As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSCOPE_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgMAXIMUM_CARDINALITY As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgDTD_IDENTIFIER As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgROUTINE_BODY As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgROUTINE_DEFINITION As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgEXTERNAL_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgEXTERNAL_LANGUAGE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgPARAMETER_STYLE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgIS_DETERMINISTIC As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSQL_DATA_ACCESS As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgIS_NULL_CALL As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSQL_PATH As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgSCHEMA_LEVEL_ROUTINE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgMAX_DYNAMIC_RESULT_SETS As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgIS_USER_DEFINED_CAST As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgIS_IMPLICITLY_INVOCABLE As System.Windows.Forms.DataGridViewTextBoxColumn
Friend WithEvents dgCREATED As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgLAST_ALTERED As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents dgCOMPLETE_NAME As System.Windows.Forms.DataGridViewTextBoxColumn
    Public Sub New(ByVal strForm As String, ByVal _bs As BindingSource, ByVal _dg As dgvEnter, _
    ByVal _ta As MetaDataSetTableAdapters.ROUTINESTableAdapter, _
    ByVal _ds As DataSet, _
    ByVal _components As System.ComponentModel.Container, _
    ByVal _MainDefs As MainDefinitions, _
    ByVal blnRO As Boolean, _
    ByVal _Controls As Control.ControlCollection, ByVal _frmStandard As frmStandard, _
    ByVal blnFilters As Boolean)
        MyBase.New(strForm, "ROUTINES", _bs, _dg, _ta, _ds, _MainDefs, blnRO, _
        "", "", _Controls, _frmStandard, blnFilters)
        _ta.Connection.ConnectionString = GetConnectionString()
    End Sub
    Public Overrides Sub Createcolumns()
        dgSPECIFIC_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSPECIFIC_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSPECIFIC_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgROUTINE_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgROUTINE_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgROUTINE_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgROUTINE_TYPE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgMODULE_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgMODULE_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgMODULE_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgUDT_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgUDT_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgUDT_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgDATA_TYPE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHARACTER_MAXIMUM_LENGTH = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHARACTER_OCTET_LENGTH = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCOLLATION_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCOLLATION_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCOLLATION_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHARACTER_SET_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHARACTER_SET_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCHARACTER_SET_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgNUMERIC_PRECISION = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgNUMERIC_PRECISION_RADIX = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgNUMERIC_SCALE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgDATETIME_PRECISION = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgINTERVAL_TYPE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgINTERVAL_PRECISION = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgTYPE_UDT_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgTYPE_UDT_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgTYPE_UDT_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSCOPE_CATALOG = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSCOPE_SCHEMA = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSCOPE_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgMAXIMUM_CARDINALITY = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgDTD_IDENTIFIER = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgROUTINE_BODY = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgROUTINE_DEFINITION = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgEXTERNAL_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgEXTERNAL_LANGUAGE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgPARAMETER_STYLE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgIS_DETERMINISTIC = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSQL_DATA_ACCESS = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgIS_NULL_CALL = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSQL_PATH = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgSCHEMA_LEVEL_ROUTINE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgMAX_DYNAMIC_RESULT_SETS = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgIS_USER_DEFINED_CAST = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgIS_IMPLICITLY_INVOCABLE = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCREATED = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgLAST_ALTERED = New System.Windows.Forms.DataGridViewTextBoxColumn
        dgCOMPLETE_NAME = New System.Windows.Forms.DataGridViewTextBoxColumn
    End Sub
    Public Overrides Sub Adjustcolumns(ByVal blnAdjustWidth As Boolean)
        MyBase.Adjustcolumns(blnAdjustWidth)
        Try
            DefineColumn(dgSPECIFIC_CATALOG, "SPECIFIC_CATALOG", blnRO, ds.ROUTINES.SPECIFIC_CATALOGColumn.MaxLength)
            DefineColumn(dgSPECIFIC_SCHEMA, "SPECIFIC_SCHEMA", blnRO, ds.ROUTINES.SPECIFIC_SCHEMAColumn.MaxLength)
            DefineColumn(dgSPECIFIC_NAME, "SPECIFIC_NAME", blnRO, ds.ROUTINES.SPECIFIC_NAMEColumn.MaxLength)
            DefineColumn(dgROUTINE_CATALOG, "ROUTINE_CATALOG", blnRO, ds.ROUTINES.ROUTINE_CATALOGColumn.MaxLength)
            DefineColumn(dgROUTINE_SCHEMA, "ROUTINE_SCHEMA", blnRO, ds.ROUTINES.ROUTINE_SCHEMAColumn.MaxLength)
            DefineColumn(dgROUTINE_NAME, "ROUTINE_NAME", blnRO, ds.ROUTINES.ROUTINE_NAMEColumn.MaxLength)
            DefineColumn(dgROUTINE_TYPE, "ROUTINE_TYPE", blnRO, ds.ROUTINES.ROUTINE_TYPEColumn.MaxLength)
            DefineColumn(dgMODULE_CATALOG, "MODULE_CATALOG", blnRO, ds.ROUTINES.MODULE_CATALOGColumn.MaxLength)
            DefineColumn(dgMODULE_SCHEMA, "MODULE_SCHEMA", blnRO, ds.ROUTINES.MODULE_SCHEMAColumn.MaxLength)
            DefineColumn(dgMODULE_NAME, "MODULE_NAME", blnRO, ds.ROUTINES.MODULE_NAMEColumn.MaxLength)
            DefineColumn(dgUDT_CATALOG, "UDT_CATALOG", blnRO, ds.ROUTINES.UDT_CATALOGColumn.MaxLength)
            DefineColumn(dgUDT_SCHEMA, "UDT_SCHEMA", blnRO, ds.ROUTINES.UDT_SCHEMAColumn.MaxLength)
            DefineColumn(dgUDT_NAME, "UDT_NAME", blnRO, ds.ROUTINES.UDT_NAMEColumn.MaxLength)
            DefineColumn(dgDATA_TYPE, "DATA_TYPE", blnRO, ds.ROUTINES.DATA_TYPEColumn.MaxLength)
            DefineColumn(dgCHARACTER_MAXIMUM_LENGTH, "CHARACTER_MAXIMUM_LENGTH", blnRO, ds.ROUTINES.CHARACTER_MAXIMUM_LENGTHColumn.MaxLength)
            DefineColumn(dgCHARACTER_OCTET_LENGTH, "CHARACTER_OCTET_LENGTH", blnRO, ds.ROUTINES.CHARACTER_OCTET_LENGTHColumn.MaxLength)
            DefineColumn(dgCOLLATION_CATALOG, "COLLATION_CATALOG", blnRO, ds.ROUTINES.COLLATION_CATALOGColumn.MaxLength)
            DefineColumn(dgCOLLATION_SCHEMA, "COLLATION_SCHEMA", blnRO, ds.ROUTINES.COLLATION_SCHEMAColumn.MaxLength)
            DefineColumn(dgCOLLATION_NAME, "COLLATION_NAME", blnRO, ds.ROUTINES.COLLATION_NAMEColumn.MaxLength)
            DefineColumn(dgCHARACTER_SET_CATALOG, "CHARACTER_SET_CATALOG", blnRO, ds.ROUTINES.CHARACTER_SET_CATALOGColumn.MaxLength)
            DefineColumn(dgCHARACTER_SET_SCHEMA, "CHARACTER_SET_SCHEMA", blnRO, ds.ROUTINES.CHARACTER_SET_SCHEMAColumn.MaxLength)
            DefineColumn(dgCHARACTER_SET_NAME, "CHARACTER_SET_NAME", blnRO, ds.ROUTINES.CHARACTER_SET_NAMEColumn.MaxLength)
            DefineColumn(dgNUMERIC_PRECISION, "NUMERIC_PRECISION", blnRO, ds.ROUTINES.NUMERIC_PRECISIONColumn.MaxLength)
            DefineColumn(dgNUMERIC_PRECISION_RADIX, "NUMERIC_PRECISION_RADIX", blnRO, ds.ROUTINES.NUMERIC_PRECISION_RADIXColumn.MaxLength)
            DefineColumn(dgNUMERIC_SCALE, "NUMERIC_SCALE", blnRO, ds.ROUTINES.NUMERIC_SCALEColumn.MaxLength)
            DefineColumn(dgDATETIME_PRECISION, "DATETIME_PRECISION", blnRO, ds.ROUTINES.DATETIME_PRECISIONColumn.MaxLength)
            DefineColumn(dgINTERVAL_TYPE, "INTERVAL_TYPE", blnRO, ds.ROUTINES.INTERVAL_TYPEColumn.MaxLength)
            DefineColumn(dgINTERVAL_PRECISION, "INTERVAL_PRECISION", blnRO, ds.ROUTINES.INTERVAL_PRECISIONColumn.MaxLength)
            DefineColumn(dgTYPE_UDT_CATALOG, "TYPE_UDT_CATALOG", blnRO, ds.ROUTINES.TYPE_UDT_CATALOGColumn.MaxLength)
            DefineColumn(dgTYPE_UDT_SCHEMA, "TYPE_UDT_SCHEMA", blnRO, ds.ROUTINES.TYPE_UDT_SCHEMAColumn.MaxLength)
            DefineColumn(dgTYPE_UDT_NAME, "TYPE_UDT_NAME", blnRO, ds.ROUTINES.TYPE_UDT_NAMEColumn.MaxLength)
            DefineColumn(dgSCOPE_CATALOG, "SCOPE_CATALOG", blnRO, ds.ROUTINES.SCOPE_CATALOGColumn.MaxLength)
            DefineColumn(dgSCOPE_SCHEMA, "SCOPE_SCHEMA", blnRO, ds.ROUTINES.SCOPE_SCHEMAColumn.MaxLength)
            DefineColumn(dgSCOPE_NAME, "SCOPE_NAME", blnRO, ds.ROUTINES.SCOPE_NAMEColumn.MaxLength)
            DefineColumn(dgMAXIMUM_CARDINALITY, "MAXIMUM_CARDINALITY", blnRO, ds.ROUTINES.MAXIMUM_CARDINALITYColumn.MaxLength)
            DefineColumn(dgDTD_IDENTIFIER, "DTD_IDENTIFIER", blnRO, ds.ROUTINES.DTD_IDENTIFIERColumn.MaxLength)
            DefineColumn(dgROUTINE_BODY, "ROUTINE_BODY", blnRO, ds.ROUTINES.ROUTINE_BODYColumn.MaxLength)
            DefineColumn(dgROUTINE_DEFINITION, "ROUTINE_DEFINITION", blnRO, ds.ROUTINES.ROUTINE_DEFINITIONColumn.MaxLength)
            DefineColumn(dgEXTERNAL_NAME, "EXTERNAL_NAME", blnRO, ds.ROUTINES.EXTERNAL_NAMEColumn.MaxLength)
            DefineColumn(dgEXTERNAL_LANGUAGE, "EXTERNAL_LANGUAGE", blnRO, ds.ROUTINES.EXTERNAL_LANGUAGEColumn.MaxLength)
            DefineColumn(dgPARAMETER_STYLE, "PARAMETER_STYLE", blnRO, ds.ROUTINES.PARAMETER_STYLEColumn.MaxLength)
            DefineColumn(dgIS_DETERMINISTIC, "IS_DETERMINISTIC", blnRO, ds.ROUTINES.IS_DETERMINISTICColumn.MaxLength)
            DefineColumn(dgSQL_DATA_ACCESS, "SQL_DATA_ACCESS", blnRO, ds.ROUTINES.SQL_DATA_ACCESSColumn.MaxLength)
            DefineColumn(dgIS_NULL_CALL, "IS_NULL_CALL", blnRO, ds.ROUTINES.IS_NULL_CALLColumn.MaxLength)
            DefineColumn(dgSQL_PATH, "SQL_PATH", blnRO, ds.ROUTINES.SQL_PATHColumn.MaxLength)
            DefineColumn(dgSCHEMA_LEVEL_ROUTINE, "SCHEMA_LEVEL_ROUTINE", blnRO, ds.ROUTINES.SCHEMA_LEVEL_ROUTINEColumn.MaxLength)
            DefineColumn(dgMAX_DYNAMIC_RESULT_SETS, "MAX_DYNAMIC_RESULT_SETS", blnRO, ds.ROUTINES.MAX_DYNAMIC_RESULT_SETSColumn.MaxLength)
            DefineColumn(dgIS_USER_DEFINED_CAST, "IS_USER_DEFINED_CAST", blnRO, ds.ROUTINES.IS_USER_DEFINED_CASTColumn.MaxLength)
            DefineColumn(dgIS_IMPLICITLY_INVOCABLE, "IS_IMPLICITLY_INVOCABLE", blnRO, ds.ROUTINES.IS_IMPLICITLY_INVOCABLEColumn.MaxLength)
            DefineColumn(dgCREATED, "CREATED", blnRO, ds.ROUTINES.CREATEDColumn.MaxLength)
            DefineColumn(dgLAST_ALTERED, "LAST_ALTERED", blnRO, ds.ROUTINES.LAST_ALTEREDColumn.MaxLength)
            DefineColumn(dgCOMPLETE_NAME, "COMPLETE_NAME", blnRO, ds.ROUTINES.COMPLETE_NAMEColumn.MaxLength)
            PutColumnsInGrid()
            AdjustDataGridWidth(blnAdjustWidth)
            RefreshCombos()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
Public Overrides Sub RefreshCombos()
MyBase.RefreshCombos()
dg.CancelEdit()
iComboCount = 0
End sub
#Region "Filter"
Public Overrides Sub CreateFilterBoxes(ByVal _Controls As Control.ControlCollection)
MyBase.CreateFilterBoxes(_Controls)
CreateAFilterBox(tbSPECIFIC_CATALOGFind, "SPECIFIC_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSPECIFIC_SCHEMAFind, "SPECIFIC_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSPECIFIC_NAMEFind, "SPECIFIC_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbROUTINE_CATALOGFind, "ROUTINE_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbROUTINE_SCHEMAFind, "ROUTINE_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbROUTINE_NAMEFind, "ROUTINE_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbROUTINE_TYPEFind, "ROUTINE_TYPE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbMODULE_CATALOGFind, "MODULE_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbMODULE_SCHEMAFind, "MODULE_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbMODULE_NAMEFind, "MODULE_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbUDT_CATALOGFind, "UDT_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbUDT_SCHEMAFind, "UDT_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbUDT_NAMEFind, "UDT_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbDATA_TYPEFind, "DATA_TYPE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCHARACTER_MAXIMUM_LENGTHFind, "CHARACTER_MAXIMUM_LENGTH", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCHARACTER_OCTET_LENGTHFind, "CHARACTER_OCTET_LENGTH", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCOLLATION_CATALOGFind, "COLLATION_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCOLLATION_SCHEMAFind, "COLLATION_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCOLLATION_NAMEFind, "COLLATION_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCHARACTER_SET_CATALOGFind, "CHARACTER_SET_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCHARACTER_SET_SCHEMAFind, "CHARACTER_SET_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCHARACTER_SET_NAMEFind, "CHARACTER_SET_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbNUMERIC_PRECISIONFind, "NUMERIC_PRECISION", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbNUMERIC_PRECISION_RADIXFind, "NUMERIC_PRECISION_RADIX", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbNUMERIC_SCALEFind, "NUMERIC_SCALE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbDATETIME_PRECISIONFind, "DATETIME_PRECISION", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbINTERVAL_TYPEFind, "INTERVAL_TYPE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbINTERVAL_PRECISIONFind, "INTERVAL_PRECISION", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbTYPE_UDT_CATALOGFind, "TYPE_UDT_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbTYPE_UDT_SCHEMAFind, "TYPE_UDT_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbTYPE_UDT_NAMEFind, "TYPE_UDT_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSCOPE_CATALOGFind, "SCOPE_CATALOG", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSCOPE_SCHEMAFind, "SCOPE_SCHEMA", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSCOPE_NAMEFind, "SCOPE_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbMAXIMUM_CARDINALITYFind, "MAXIMUM_CARDINALITY", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbDTD_IDENTIFIERFind, "DTD_IDENTIFIER", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbROUTINE_BODYFind, "ROUTINE_BODY", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbROUTINE_DEFINITIONFind, "ROUTINE_DEFINITION", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbEXTERNAL_NAMEFind, "EXTERNAL_NAME", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbEXTERNAL_LANGUAGEFind, "EXTERNAL_LANGUAGE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbPARAMETER_STYLEFind, "PARAMETER_STYLE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbIS_DETERMINISTICFind, "IS_DETERMINISTIC", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSQL_DATA_ACCESSFind, "SQL_DATA_ACCESS", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbIS_NULL_CALLFind, "IS_NULL_CALL", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSQL_PATHFind, "SQL_PATH", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbSCHEMA_LEVEL_ROUTINEFind, "SCHEMA_LEVEL_ROUTINE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbMAX_DYNAMIC_RESULT_SETSFind, "MAX_DYNAMIC_RESULT_SETS", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbIS_USER_DEFINED_CASTFind, "IS_USER_DEFINED_CAST", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbIS_IMPLICITLY_INVOCABLEFind, "IS_IMPLICITLY_INVOCABLE", AddressOf tbFind_TextChanged, _Controls)
CreateAFilterBox(tbCREATEDFind, "CREATED", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbLAST_ALTEREDFind, "LAST_ALTERED", AddressOf tbFind_TextChanged, _Controls)
        CreateAFilterBox(tbCOMPLETE_NAMEFind, "COMPLETE_NAME", AddressOf tbFind_TextChanged, _Controls)
End Sub
Friend WithEvents tbSPECIFIC_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbSPECIFIC_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbSPECIFIC_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbROUTINE_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbROUTINE_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbROUTINE_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbROUTINE_TYPEFind As System.Windows.Forms.TextBox
Friend WithEvents tbMODULE_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbMODULE_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbMODULE_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbUDT_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbUDT_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbUDT_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbDATA_TYPEFind As System.Windows.Forms.TextBox
Friend WithEvents tbCHARACTER_MAXIMUM_LENGTHFind As System.Windows.Forms.TextBox
Friend WithEvents tbCHARACTER_OCTET_LENGTHFind As System.Windows.Forms.TextBox
Friend WithEvents tbCOLLATION_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbCOLLATION_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbCOLLATION_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbCHARACTER_SET_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbCHARACTER_SET_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbCHARACTER_SET_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbNUMERIC_PRECISIONFind As System.Windows.Forms.TextBox
Friend WithEvents tbNUMERIC_PRECISION_RADIXFind As System.Windows.Forms.TextBox
Friend WithEvents tbNUMERIC_SCALEFind As System.Windows.Forms.TextBox
Friend WithEvents tbDATETIME_PRECISIONFind As System.Windows.Forms.TextBox
Friend WithEvents tbINTERVAL_TYPEFind As System.Windows.Forms.TextBox
Friend WithEvents tbINTERVAL_PRECISIONFind As System.Windows.Forms.TextBox
Friend WithEvents tbTYPE_UDT_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbTYPE_UDT_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbTYPE_UDT_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbSCOPE_CATALOGFind As System.Windows.Forms.TextBox
Friend WithEvents tbSCOPE_SCHEMAFind As System.Windows.Forms.TextBox
Friend WithEvents tbSCOPE_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbMAXIMUM_CARDINALITYFind As System.Windows.Forms.TextBox
Friend WithEvents tbDTD_IDENTIFIERFind As System.Windows.Forms.TextBox
Friend WithEvents tbROUTINE_BODYFind As System.Windows.Forms.TextBox
Friend WithEvents tbROUTINE_DEFINITIONFind As System.Windows.Forms.TextBox
Friend WithEvents tbEXTERNAL_NAMEFind As System.Windows.Forms.TextBox
Friend WithEvents tbEXTERNAL_LANGUAGEFind As System.Windows.Forms.TextBox
Friend WithEvents tbPARAMETER_STYLEFind As System.Windows.Forms.TextBox
Friend WithEvents tbIS_DETERMINISTICFind As System.Windows.Forms.TextBox
Friend WithEvents tbSQL_DATA_ACCESSFind As System.Windows.Forms.TextBox
Friend WithEvents tbIS_NULL_CALLFind As System.Windows.Forms.TextBox
Friend WithEvents tbSQL_PATHFind As System.Windows.Forms.TextBox
Friend WithEvents tbSCHEMA_LEVEL_ROUTINEFind As System.Windows.Forms.TextBox
Friend WithEvents tbMAX_DYNAMIC_RESULT_SETSFind As System.Windows.Forms.TextBox
Friend WithEvents tbIS_USER_DEFINED_CASTFind As System.Windows.Forms.TextBox
Friend WithEvents tbIS_IMPLICITLY_INVOCABLEFind As System.Windows.Forms.TextBox
Friend WithEvents tbCREATEDFind As System.Windows.Forms.TextBox
    Friend WithEvents tbLAST_ALTEREDFind As System.Windows.Forms.TextBox
    Friend WithEvents tbCOMPLETE_NAMEFind As System.Windows.Forms.TextBox
Private Sub cbFind_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
MakeFilter(False)
End Sub
Private Sub tbFind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
 MakeFilter(False)
End Sub
#End Region
End Class
