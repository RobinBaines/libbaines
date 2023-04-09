'------------------------------------------------
'Name: MainDefinitions.vb.
'Function: Some public constants.
'Copyright Robin Baines 2008. All rights reserved.
'Modifications: 
'------------------------------------------------
Imports Microsoft.Win32
Imports System.Windows.Forms
Public Class MainDefinitions
    Public DEFAULTVERTICALRESOLUTION As Integer
    Public DEFAULTHORIZONTALRESOLUTION As Integer
    Private components As System.ComponentModel.IContainer
    Public DEFAULTROWHEADERSWIDTH As Integer

    Public dsFormLevels As New DataSet
    Dim ut As New Utilities

    Public MainForm As MainForm
    Private _User As String

    Public Sub New(ByVal _ParentForm As MainForm)
        MyBase.New()
        MainForm = _ParentForm
        _User = ""
        blnActiveFilters = True
        DEFAULTVERTICALRESOLUTION = 1024
        DEFAULTHORIZONTALRESOLUTION = 1280
        DEFAULTROWHEADERSWIDTH = 25
    End Sub

    'When user double clicks a field in the form ActiveFilter true means filter the form on that value.
    Private _blnActiveFilters As Boolean
    Public Property blnActiveFilters() As Boolean
        Get
            Return _blnActiveFilters
        End Get
        Set(ByVal value As Boolean)
            _blnActiveFilters = value
        End Set
    End Property

    Private _strIniFile As String
    Public Property strIniFile() As String
        Get
            Return _strIniFile
        End Get
        Set(ByVal value As String)
            _strIniFile = value
        End Set
    End Property

    Public Enum StatusValues
        OK
        NOK
    End Enum

    Public ReadOnly Property FRMNAME_DomainTables() As String
        Get
            Return "DomainTables"
        End Get
    End Property

    Public Overridable Function GetConnectionString() As String
        Return MainForm.GetConnectionString()    '(My.Settings.QualityDatabase)
    End Function

    Public Function GetConnectionString(ByVal blnQuality As Boolean) As String
        Return MainForm.GetConnectionString(blnQuality)
    End Function

#Region "Format"
    Public Overridable Function strGetFormat(ByVal strFieldType As String) As String
        'Enters with user type if defined.
        'else if a numeric with the column name.
        'else with sqltype.

        Dim strFormat As String
        strFormat = ""
        'Select Case strFieldType
        '    Case "TYP_QTY"
        '        strFormat = "0.00\%"
        '    Case "Float"
        '        strFormat = "N3"
        '    Case "pr_fixed_cost"
        '        strFormat = "N3"
        '    Case "pr_var_cost"
        '        strFormat = "N3"
        '    Case "pr_cost"
        '        strFormat = "N3"

        '    Case "fr_fixed_cost"
        '        strFormat = "N3"
        '    Case "fr_var_cost"
        '        strFormat = "N3"
        '    Case "fr_cost"
        '        strFormat = "N3"

        '    Case "margin_rsp"
        '        strFormat = "0\%"
        '    Case "margin_other"
        '        strFormat = "0\%"
        '    Case "total_price"
        '        strFormat = "N2"
        '    Case "total_price_adjusted"
        '        strFormat = "N2"
        '    Case "filling_subtotal"
        '        strFormat = "N3"

        'Case "TYP_DOSAGE"
        '    strFormat = "0.00\%" '"P2"
        'Case "TYP_SCRAP"
        '    strFormat = "0\%"
        'Case "TYP_MARGIN"
        '    strFormat = "0\%"
        'Case "TYP_LABOR"    'hours.
        '    strFormat = "N2"
        'Case "TYP_CALLSIZE"
        '    strFormat = "N0"
        'Case "TYP_LOTSIZE"
        '    strFormat = "N0"
        'Case "TYP_MONEY"
        '    strFormat = "N3"
        'Case "Money"
        '    strFormat = "N2"
        'Case "TYP_VER"
        '    strFormat = "N0"
        'Case "qc_cost"
        '    strFormat = "N3"
        'Case "xqc_cost"
        '    strFormat = "N3"
        'Case "total_cost"
        '    strFormat = "N2"
        'Case "mat_cost"
        '    strFormat = "N2"
        'Case "scrap_cost"
        '    strFormat = "N3"
        'Case "prod_cost"
        '    strFormat = "N3"
        'Case "inter_prod_cost"
        '    strFormat = "N3"

        'Case "filling_cost"
        '    strFormat = "N3"


        'Case "fixed_prod_cost"
        '    strFormat = "N3"
        'Case "var_prod_cost"
        '    strFormat = "N3"
        'Case "fixed_inter_prod_cost"
        '    strFormat = "N3"
        'Case "var_inter_prod_cost"
        '    strFormat = "N3"
        'Case "packmat_cost"
        '    strFormat = "N3"
        'Case "fixed_filling_cost"
        '    strFormat = "N3"
        'Case "var_filling_cost"
        '    strFormat = "N3"
        'Case "qc_cost"
        '    strFormat = "N3"
        'Case "total_cost"
        '    strFormat = "N2"
        'Case "trans_cost"
        '    strFormat = "N3"
        'Case "use_price_low"
        '    strFormat = "N2"
        'Case "use_price_high"
        '    strFormat = "N2"
        'Case "total_price"
        '    strFormat = "N2"
        'Case "fixed_cost"
        '    strFormat = "N2"
        'Case "var_cost"
        '    strFormat = "N2"
        'Case "inter_cost"
        '    strFormat = "N3"

        ' End Select

        Return strFormat
    End Function

    Public Function strGetColumnHeader(ByVal strFieldName As String) As String
        Return ""
    End Function

    Public Overridable Function strGetTableText(ByVal strFieldName As String) As String
        Return strFieldName
    End Function

    Public Function strGetPrintString(ByVal strTableName As String, ByVal strFieldName As String) As String
        Return ""
    End Function

    Public Function blnGetvisibility(ByVal strTableName As String, ByVal strFieldName As String) As Boolean
        Return True
    End Function

    Public Overridable Function strGetDisplayMember(ByVal strTableName As String, ByVal strFieldName As String) As String
        Return strFieldName
    End Function

    Public Function blnGetRO(ByVal blnRO As Boolean, ByVal strFieldName As String) As Boolean
        Dim blnRet = blnRO
        Return blnRet
    End Function

#End Region

#Region "PrintDefines"
    Public ReadOnly Property DONOTPRINT() As String
        Get
            Return "P"
        End Get
    End Property
#End Region

#Region "FieldWidths"
    
    'Public ReadOnly Property PLANTWIDTH() As Integer
    '    Get
    '        Return 35 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property FLOATWIDTH() As Integer
    '    Get
    '        Return 65 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property CLASSWIDTH() As Integer
    '    Get
    '        Return 45 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property FLOATWIDTHWIDE() As Integer
    '    Get
    '        Return 75 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property RATEWIDTH() As Integer
    '    Get
    '        Return 30 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property MONTHWIDTH() As Integer
    '    Get
    '        Return 40 * GetFieldFactor()
    '    End Get
    'End Property
    'Public ReadOnly Property ERRORWIDTH() As Integer
    '    Get
    '        Return 20 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property MATERIALWIDTH() As Integer
    '    Get
    '        Return 90 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property GENWIDTH() As Integer
    '    Get
    '        Return 55 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property SMALLWIDTH() As Integer
    '    Get
    '        Return 30 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property DATEWIDTH() As Integer
    '    Get
    '        Return 125 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property REMARKWIDTH() As Integer
    '    Get
    '        Return 200 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property SMALLREMARKWIDTH() As Integer
    '    Get
    '        Return 150 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property Numeric() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property Money() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property Int() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_M_COLUMN() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property VarChar() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property NVarCharMax() As Integer
    '    Get
    '        Return REMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_OPTIVA() As Integer
    '    Get
    '        Return 90 * GetFieldFactor()
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_SCENARIO() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_USER() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_EMAIL() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_BOOL() As Integer
    '    Get
    '        Return SMALLWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_SCRAP() As Integer
    '    Get
    '        Return FLOATWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property Float() As Integer
    '    Get
    '        Return FLOATWIDTHWIDE
    '    End Get
    'End Property
    'Public ReadOnly Property NVarChar() As Integer
    '    Get
    '        Return FLOATWIDTHWIDE
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_MARGIN() As Integer
    '    Get
    '        Return FLOATWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_DESCR() As Integer
    '    Get
    '        Return REMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_PLANT() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_QTY() As Integer
    '    Get
    '        Return FLOATWIDTHWIDE
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_RCPGRP() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_LOTSIZE() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_CALLSIZE() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_CUR() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_LABOR() As Integer
    '    Get
    '        Return FLOATWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_MONEY() As Integer
    '    Get
    '        Return FLOATWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_COSTCENTER() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_MATGRP() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_ACTTYP() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_CAT() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_SUBCAT() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_APPL() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_SUBCAT2() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_REGION() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_LEGIS() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_MAT() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_PRODCODE() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_CUSTCODE() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_CSF() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_PRJ() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_EST() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_DOSAGE() As Integer
    '    Get
    '        Return FLOATWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_CUST() As Integer
    '    Get
    '        Return SMALLREMARKWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_STRING() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_PAYS() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_ZONE() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_INCO() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property Bit() As Integer
    '    Get
    '        Return GENWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_VOLUME() As Integer
    '    Get
    '        Return FLOATWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_CUSTSEG() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_VER() As Integer
    '    Get
    '        Return SMALLWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_STATUS() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_DATE() As Integer
    '    Get
    '        Return DATEWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_USR() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property
    'Public ReadOnly Property TYP_MARKET() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public ReadOnly Property TYP_ANNVOL() As Integer
    '    Get
    '        Return MATERIALWIDTH
    '    End Get
    'End Property

    'Public Function GetFieldFactor() As Double
    '    Dim dFactor As Double = 1
    '    dFactor = My.Computer.Screen.Bounds.Width / DEFAULTHORIZONTALRESOLUTION
    '    If dFactor > 1 Then
    '        dFactor = 1 + (dFactor - 1) * 0.6
    '    End If
    '    Return dFactor
    'End Function

    'Public Function GetFactor(ByVal iAdjustable As Integer) As Double
    '    Dim dFactor As Double = 1
    '    Dim dEffectiveHeight As Double = My.Computer.Screen.Bounds.Height * DEFAULTHORIZONTALRESOLUTION / My.Computer.Screen.Bounds.Width
    '    If dEffectiveHeight < (DEFAULTVERTICALRESOLUTION - 100) Then
    '        dFactor = 1 + (dEffectiveHeight - DEFAULTVERTICALRESOLUTION) / iAdjustable
    '    End If
    '    Return dFactor
    'End Function
#End Region
End Class
