'------------------------------------------------
'Name: Module for statics.vb.
'Function: 
'Copyright Robin Baines 2010. All rights reserved.
'Created May 2010.
'Notes: 
'Modifications: 
'20170501 Try/catch around adding a key to hash tables because some values can be duplicated. This happened when m_txt was being filled with application data.
'20211112 Added try/catch with msgbox during initialisation calls Init() and ModuleStatics.InitConnection().
'------------------------------------------------
Imports System
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data.SqlTypes
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Security.Principal

Namespace statics
    Public Module ModuleStatics

        'If the usr name has rights to 'All'  the forms then every menu is active. 
        Const constALL = "All"

        Public Const GroupBoxRelativeVerticalLocation = 32

        'Create hash tables for the main lookup tables. These tables are read multiple times and for this reason they are in memory.
        Dim List_m_form_tble_column__visibility As New Hashtable()
        Dim List_m_tble_column As New Hashtable()
        Dim List_m_tble_column_header As New Hashtable()
        Dim List_m_txt As New Hashtable()
        Dim List_m_txt_header As New Hashtable()

        'create hash tables to speed up searching.
        Sub CreateHashTables()
            List_m_form_tble_column__visibility.Clear()
            For Each row As TheDataSet.m_form_tble_column__visibilityRow In m_form_tble_column__visibility
                Dim iHash = (row.form.ToUpper() + "*" + row.tble.ToUpper() + "*" + row.colmn.ToUpper()).GetHashCode()
                List_m_form_tble_column__visibility.Add(iHash, row)
            Next

            List_m_tble_column.Clear()
            For Each row As TheDataSet.m_tble_columnRow In m_tble_column
                Dim iHash = (row.tble.ToUpper() + "*" + row.colmn.ToUpper()).GetHashCode()

                '201705 Try/catch around adding a key to hash tables because some values can be duplicated. This happened when m_txt was being filled with application data.
                Try
                    List_m_tble_column.Add(iHash, row)
                Catch ex As Exception
                End Try

            Next

            List_m_tble_column_header.Clear()
            For Each row As TheDataSet.v_tble_column_headerRow In m_tble_column_header
                Dim iHash = (row.tble.ToUpper() + "*" + row.colmn.ToUpper()).GetHashCode()

                '201705 Try/catch around adding a key to hash tables because some values can be duplicated. This happened when m_txt was being filled with application data.
                Try
                    List_m_tble_column_header.Add(iHash, row)
                Catch ex As Exception
                End Try
            Next

            List_m_txt.Clear()
            For Each row As TheDataSet.m_txtRow In m_txt
                Dim iHash = (row.txt.ToUpper()).GetHashCode()

                '201705 Try/catch around adding a key to hash tables because some values can be duplicated. This happened when m_txt was being filled with application data.
                Try
                    List_m_txt.Add(iHash, row)
                Catch ex As Exception
                End Try
            Next

            List_m_txt_header.Clear()
            For Each row As TheDataSet.m_txt_headerRow In m_txt_header
                Dim iHash = (row.txt.ToUpper()).GetHashCode()

                '201705 Try/catch around adding a key to hash tables because some values can be duplicated. This happened when m_txt was being filled with application data.
                Try
                    List_m_txt_header.Add(iHash, row)
                Catch ex As Exception
                End Try
            Next
        End Sub

#Region "Properties"
        Public Enum StatusValues
            OK
            NOK
        End Enum

        'this property is used in dbColumns when the groupbox is created and is used in frmStandard.ReadOnlyGroupBox to prevent this groupbox
        'being disabled when a form is read only.
        Public ReadOnly Property FilterGroupBoxName() As String
            Get
                Return "gbForFiltersGroupBox"
            End Get
        End Property

        Private _DatabaseVersion As Integer
        Public ReadOnly Property DatabaseVersion() As Integer
            Get
                Return _DatabaseVersion
            End Get
        End Property

        Private _DomainAndUserName As String = ""
        Public ReadOnly Property DomainAndUserName() As String
            Get
                Return _DomainAndUserName
            End Get
        End Property

        Private _UserName As String = ""
        Public ReadOnly Property UserName() As String
            Get
                Return _UserName
            End Get
        End Property

        Private _Lang As String = ""
        Public ReadOnly Property Lang() As String
            Get
                Return _Lang
            End Get
        End Property

        Private _Blocked As Boolean = False
        Public ReadOnly Property Blocked() As Boolean
            Get
                Return _Blocked
            End Get
        End Property

        'A user is authorized if they are in the m_usr table.
        Private _IsAuthorized As Boolean = False
        Public ReadOnly Property IsAuthorized() As Boolean
            Get
                Return _IsAuthorized
            End Get
        End Property

        'the group the user is in.
        Private _Group As String = ""
        Public ReadOnly Property Group() As String
            Get
                Return _Group
            End Get
        End Property

        'provide bools to indicate if database has tables and supports associated functionality.
        Private _blnSupportsMultiLang As Boolean = False
        Public ReadOnly Property blnSupportsMultiLang() As Boolean
            Get
                Return _blnSupportsMultiLang
            End Get
        End Property

        Private _blnSupportsUserLog As Boolean = False
        Public ReadOnly Property blnSupportsUserLog() As Boolean
            Get
                Return _blnSupportsUserLog
            End Get
        End Property

        Private _blnSupportsOptions As Boolean = False
        Public ReadOnly Property blnSupportsOptions() As Boolean
            Get
                Return _blnSupportsOptions
            End Get
        End Property

        Private _blnSupportsAppLog As Boolean = False
        Public ReadOnly Property blnSupportsAppLog() As Boolean
            Get
                Return _blnSupportsAppLog
            End Get
        End Property

        Private _blnSupportsColorOptions As Boolean = False
        Public ReadOnly Property blnSupportsColorOptions() As Boolean
            Get
                Return _blnSupportsColorOptions
            End Get
        End Property

        Dim lMinutesLoggedOn As Long

        Friend WithEvents m_versionTableAdapter As TheDataSetTableAdapters.m_versionTableAdapter
        Friend WithEvents v_usr_grpTableAdapter As TheDataSetTableAdapters.v_usr_grpTableAdapter
        Friend WithEvents v_usr_grp_groupboxTableAdapter As TheDataSetTableAdapters.v_usr_grp_groupboxTableAdapter
        Friend WithEvents v_usr_blockedTableAdapter As TheDataSetTableAdapters.v_usr_blockedTableAdapter
        Friend WithEvents m_txtTableAdapter As TheDataSetTableAdapters.m_txtTableAdapter
        Friend WithEvents m_txt_headerTableAdapter As TheDataSetTableAdapters.m_txt_headerTableAdapter
        Friend WithEvents m_form_tble_column__visibilityTableAdapter As TheDataSetTableAdapters.m_form_tble_column__visibilityTableAdapter
        Friend WithEvents v_tble_column_headerTableAdapter As TheDataSetTableAdapters.v_tble_column_headerTableAdapter
        Friend WithEvents m_usr_logTableAdapter As TheDataSetTableAdapters.m_usr_logTableAdapter

        'tables for updating.
        Friend WithEvents m_formTableAdapter As TheDataSetTableAdapters.m_formTableAdapter
        Friend WithEvents m_tbleTableAdapter As TheDataSetTableAdapters.m_tbleTableAdapter
        Friend WithEvents m_form_tbleTableAdapter As TheDataSetTableAdapters.m_form_tbleTableAdapter
        Friend WithEvents m_tble_columnTableAdapter As TheDataSetTableAdapters.m_tble_columnTableAdapter
        Friend WithEvents AppParametersTableAdapter As TheDataSetTableAdapters.b_app_parameterTableAdapter
        Friend WithEvents AppColorsTableAdapter As TheDataSetTableAdapters.b_app_colorTableAdapter
        Friend WithEvents m_app_logTableAdapter As TheDataSetTableAdapters.m_app_logTableAdapter

        Private _strLastForm As String
        Public Property strLastForm() As String
            Get
                Return _strLastForm
            End Get
            Set(ByVal value As String)
                _strLastForm = value
            End Set
        End Property

        Private _strLastTble As String
        Public Property strLastTble() As String
            Get
                Return _strLastTble
            End Get
            Set(ByVal value As String)
                _strLastTble = value
            End Set
        End Property
#End Region

#Region "Tables"
        Dim m_version As TheDataSet.m_versionDataTable
        Dim v_usr_grps As TheDataSet.v_usr_grpDataTable
        Dim v_usr_grp_groupbox As TheDataSet.v_usr_grp_groupboxDataTable
        Dim m_usr_grp_groupbox As TheDataSet.v_usr_grp_groupboxDataTable
        Dim m_form As TheDataSet.m_formDataTable
        Dim v_usr_blocked As TheDataSet.v_usr_blockedDataTable
        Dim m_txt As TheDataSet.m_txtDataTable
        Dim m_txt_header As TheDataSet.m_txt_headerDataTable
        Dim m_tble_column As TheDataSet.m_tble_columnDataTable
        Dim m_tble_column_header As TheDataSet.v_tble_column_headerDataTable
        Dim m_form_tble_column__visibility As TheDataSet.m_form_tble_column__visibilityDataTable
        Dim ParentForm As MainForm
#End Region

#Region "Initialise "

        Private Sub SetConnectionStrings()
            m_versionTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            v_usr_grpTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            v_usr_grp_groupboxTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            v_usr_blockedTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_usr_logTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_txtTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_txt_headerTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            v_tble_column_headerTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_formTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_tbleTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_form_tbleTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_tble_columnTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_form_tble_column__visibilityTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            AppParametersTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            AppColorsTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
            m_app_logTableAdapter.Connection.ConnectionString = ParentForm.GetConnectionString()
        End Sub
        Public Sub Init(ByVal _ParentForm As MainForm)
            Dim logonToken As IntPtr = statics.LogonUser()
            Dim windowsIdentity As New WindowsIdentity(logonToken)

            'set DomainAndUsername
            _DomainAndUserName = windowsIdentity.Name

            'set UserName
            Dim i As Integer = windowsIdentity.Name.IndexOf("\")
            If i > 0 Then
                _UserName = windowsIdentity.Name.Substring(i + 1)
            Else
                _UserName = windowsIdentity.Name
            End If

            ParentForm = _ParentForm
            m_versionTableAdapter = New TheDataSetTableAdapters.m_versionTableAdapter
            v_usr_grpTableAdapter = New TheDataSetTableAdapters.v_usr_grpTableAdapter
            v_usr_grp_groupboxTableAdapter = New TheDataSetTableAdapters.v_usr_grp_groupboxTableAdapter
            v_usr_blockedTableAdapter = New TheDataSetTableAdapters.v_usr_blockedTableAdapter
            m_usr_logTableAdapter = New TheDataSetTableAdapters.m_usr_logTableAdapter
            m_txtTableAdapter = New TheDataSetTableAdapters.m_txtTableAdapter
            m_txt_headerTableAdapter = New TheDataSetTableAdapters.m_txt_headerTableAdapter
            v_tble_column_headerTableAdapter = New TheDataSetTableAdapters.v_tble_column_headerTableAdapter
            m_formTableAdapter = New TheDataSetTableAdapters.m_formTableAdapter
            m_tbleTableAdapter = New TheDataSetTableAdapters.m_tbleTableAdapter
            m_form_tbleTableAdapter = New TheDataSetTableAdapters.m_form_tbleTableAdapter
            m_tble_columnTableAdapter = New TheDataSetTableAdapters.m_tble_columnTableAdapter
            m_form_tble_column__visibilityTableAdapter = New TheDataSetTableAdapters.m_form_tble_column__visibilityTableAdapter
            AppParametersTableAdapter = New TheDataSetTableAdapters.b_app_parameterTableAdapter
            AppColorsTableAdapter = New TheDataSetTableAdapters.b_app_colorTableAdapter
            m_app_logTableAdapter = New TheDataSetTableAdapters.m_app_logTableAdapter

        End Sub

        Public Sub InitConnection()
            Try
                SetConnectionStrings()
            Catch ex As Exception
                MsgBox("SetConnectionStrings " & ex.Message)
            End Try

            Dim iOptions As Integer
            Try
                iOptions = AppParametersTableAdapter.ParameterCount()
                If iOptions > 0 Then
                    _blnSupportsOptions = True
                End If
            Catch ex As Exception
                MsgBox("AppParametersTableAdapter " & ex.Message)
            End Try

            Try
                iOptions = AppColorsTableAdapter.ParameterCount()
                If iOptions > 0 Then
                    _blnSupportsColorOptions = True
                End If
            Catch ex As Exception
                MsgBox("AppColorsTableAdapter " & ex.Message)
            End Try

            'Do loaddata to set Blocked before Updating the usr log.
            Try
                LoadData()
            Catch ex As Exception
                MsgBox("LoadData " & ex.Message)
            End Try

            Try
                'write to the user log 
                UpdateUsrLog(UserName, False, _Blocked)
                lMinutesLoggedOn = Now().Ticks
            Catch ex As Exception
                MsgBox("UpdateUsrLog " & ex.Message)
            End Try

            Try
                'Test whether the m_app_log table has been defined on the database.
                Dim m_app_logTable As TheDataSet.m_app_logDataTable = m_app_logTableAdapter.GetData()
                _blnSupportsAppLog = True
            Catch ex As Exception
                _blnSupportsAppLog = False
                MsgBox(" m_app_log table has NOT been defined on the database " & ex.Message)
            End Try
        End Sub

        'Get the Function/Sub name where the log text originated.
        Private Function GetSource() As String
            Dim st As StackTrace = New StackTrace()
            Dim sf As StackFrame = st.GetFrame(1)
            For intF As Integer = 0 To st.FrameCount - 1
                If st.GetFrames(intF).GetMethod.Name <> "GetSource" And _
                st.GetFrames(intF).GetMethod.Name <> "WriteLogText" And _
                st.GetFrames(intF).GetMethod.Name <> "UpdateAppLog" And _
                st.GetFrames(intF).GetMethod.Name <> "LogError" Then
                    Return st.GetFrames(intF).GetMethod.Name
                End If
            Next
            Return ""
        End Function

        Public Sub UpdateAppLog(ByVal strError As String, ByVal iPriority As Integer)
            If _blnSupportsAppLog = True Then
                Try
                    UpdateAppLog(strError, GetSource(), iPriority)
                Catch ex As Exception
                End Try
            End If
        End Sub

        '20110413 RPB added the application log table for TPITrack.
        Public Sub UpdateAppLog(ByVal strError As String, ByVal strRaisedIn As String, ByVal iPriority As Integer)
            If _blnSupportsAppLog = True Then
                Try
                    m_app_logTableAdapter.Insert(Application.ProductName, _UserName, strError, strRaisedIn, iPriority)
                Catch ex As Exception
                End Try
            End If
        End Sub

        '20110413 RPB added the application log table delete for TPITrack.
        'This procedure uses the b_app_parameter "delete_app_log" parameter.
        Public Sub p_delete_app_log()
            m_app_logTableAdapter.p_delete_app_log()
        End Sub

        'also called from frmManage when usr is unlocked.
        Public Sub UpdateUsrLog(ByVal Username As String, ByVal logout As Boolean, ByVal block As Boolean)
            m_usr_logTableAdapter.Insert(Application.ProductName, logout, DomainAndUserName, Username, 0, block)
        End Sub

        Public Sub LoadData()

            'get the database version and whether multilanguage.
            _blnSupportsMultiLang = False 'as default
            _DatabaseVersion = 1 'as default
            Try
                m_version = m_versionTableAdapter.GetData()
                For Each row As TheDataSet.m_versionRow In m_version
                    _blnSupportsMultiLang = row.multi_lang
                    _DatabaseVersion = row.version

                    _DatabaseVersion = 2
                    _blnSupportsMultiLang = True
                    Exit For
                Next
            Catch ex As Exception
                MsgBox("m_versionTableAdapter " & ex.Message)
            End Try

            '''''''''''''''''''''''''''
            'load the security form levels if necessary.
            v_usr_grps = v_usr_grpTableAdapter.GetDataBy(UserName)
            v_usr_grp_groupbox = v_usr_grp_groupboxTableAdapter.GetDataByUsr(UserName)
            _Lang = "EN" 'as default. This is done so that longer error messages also work.
            _Blocked = False

            'If the user log is supported then the v_usr_blocked table will contain the blocked setting of the user.
            _blnSupportsUserLog = False
            Try
                Dim strApp = Application.ProductName
                v_usr_blocked = v_usr_blockedTableAdapter.GetDataBy(UserName)   'RPB removed strApp because it fails the first time otherwise.
                If v_usr_blocked.Count = 0 Then
                    _IsAuthorized = False
                    _Group = ""
                Else
                    _IsAuthorized = True
                    For Each row As TheDataSet.v_usr_blockedRow In v_usr_blocked
                        _Lang = row.lang
                        _Group = row.grp
                        If row.blocked = 1 Then
                            _Blocked = True
                        End If
                        Exit For
                    Next
                End If
                _blnSupportsUserLog = True
            Catch ex As Exception
                MsgBox("v_usr_grps " & ex.Message)
            End Try

            If _blnSupportsMultiLang = True Then
                Try
                    m_txt_header = m_txt_headerTableAdapter.GetDataBy(Lang)
                Catch ex As Exception
                    MsgBox("m_txt_header " & ex.Message)
                End Try

                Try
                    m_form = m_formTableAdapter.GetData()
                Catch ex As Exception
                    MsgBox("m_form " & ex.Message)
                End Try

                Try
                    m_txt = m_txtTableAdapter.GetData()
                Catch ex As Exception
                    MsgBox("m_txt " & ex.Message)
                End Try

                Try
                    m_tble_column = m_tble_columnTableAdapter.GetData()
                Catch ex As Exception
                    MsgBox("m_tble_column " & ex.Message)
                End Try

                Try
                    m_tble_column_header = v_tble_column_headerTableAdapter.GetDataBy(Lang)
                Catch ex As Exception
                    MsgBox("m_tble_column_header " & ex.Message)
                End Try

                '20110411 The GetData is sorted on sequence for the ui. Have added a not sorted version to speed 
                'things up.
                Try
                    m_form_tble_column__visibility = m_form_tble_column__visibilityTableAdapter.GetDataNotSorted()
                Catch ex As Exception
                    MsgBox("m_form_tble_column__visibility " & ex.Message)
                End Try

                Try
                    CreateHashTables()
                Catch ex As Exception
                    MsgBox("CreateHashTables " & ex.Message)
                End Try

            End If
        End Sub

        Sub Closing()
            If _blnSupportsUserLog = True Then
                Dim elapsedSpan As New TimeSpan(Now().Ticks - lMinutesLoggedOn)
                m_usr_logTableAdapter.Insert(Application.ProductName, True, DomainAndUserName, UserName, elapsedSpan.Minutes, Blocked)
            End If
        End Sub
#End Region

#Region "Parameters"
        Public Function GetParameter(ByVal strName As String) As String
            Dim strRet As String = ""
            If _blnSupportsOptions = True Then
                AppParametersTableAdapter.p_get_app_parameter(strName, strRet)
            End If
            Return strRet
        End Function

        Public Function GetColor(ByVal strName As String) As String
            Dim strRet As String = ""
            If _blnSupportsOptions = True Then
                AppColorsTableAdapter.p_get_app_color(strName, strRet)
            End If
            Return strRet
        End Function

        '20130127 Added this function.
        Public Function GetColor(ByVal strName As String, DefaultColor As Color) As Color
            Dim strRet As String = ""
            Dim RetColor As Color = DefaultColor
            If _blnSupportsOptions = True Then
                strRet = GetColor(strName)
                If strRet <> "" Then
                    RetColor = Color.FromName(strRet)
                End If
            End If
            Return RetColor
        End Function

#End Region

#Region "Publics"
        ''' <summary>
        ''' return the RO flag for the user, form, groupbox.
        ''' </summary>
        ''' <param name="strForm"></param>
        ''' <param name="strGroupbox"></param>
        ''' <returns>true means read only. Returns false if the groupbox is not in the table. This happens the first time the groupbox
        ''' is shown in a form. Do this because the default in the table is read only = false.</returns>
        ''' <remarks></remarks>
        Public Function blnCheckGroupBoxLevel(ByVal strForm As String, strGroupbox As String) As Boolean

            strForm = strForm.Trim()
            strGroupbox = strGroupbox.Trim()
            Dim blnRet As Boolean = False
            For Each row As TheDataSet.v_usr_grp_groupboxRow In v_usr_grp_groupbox
                If row.form.ToString().ToUpper = strForm.Trim.ToUpper And row.groupbox.ToString().ToUpper = strGroupbox.Trim.ToUpper Then
                    blnRet = row.RO
                    Exit For
                End If
            Next
            Return blnRet
        End Function

        Private Function blnCheckLevel(ByVal strForm As String, ByRef blnRO As Boolean, ByVal blnAdd As Boolean) As Boolean
            Return blnCheckLevel(strForm, blnRO, blnAdd, False)
        End Function

        'added to get control over the blnMenu. Is used for Dialogs which should always be opened if the parent form can be opened.
        Public Function blnCheckLevel(ByVal strForm As String, ByRef blnRO As Boolean, ByVal blnAdd As Boolean, blnMenu As Boolean) As Boolean

            '20100105 RPB modified blnCheckLevel by Trimming strForm
            strForm = strForm.Trim()
            Dim blnRet As Boolean = False
            Dim blnIntable As Boolean = False

            '20100426 If a form is named then let it override ALL.
            'if form or all not found then blnRO = true
            blnRO = True
            Try
                For Each row As TheDataSet.v_usr_grpRow In v_usr_grps
                    If row.form.ToString().ToUpper = constALL.ToUpper() And blnRet = False Then
                        blnRet = True
                        blnRO = row.RO
                        'Exit For
                    End If
                    If row.form.ToString().ToUpper = strForm.Trim.ToUpper Then
                        blnRet = True
                        blnIntable = True
                        blnRO = row.RO
                        Exit For
                    End If
                Next
            Catch ex As Exception

            End Try

            'In some isolated cases (a dialog) the form may not be in m_form table.
            If blnAdd And blnIntable = False Then
                statics.put_v_form(strForm, False, blnMenu)
            End If

            Return blnRet
        End Function

        ''' <summary>
        ''' Return true if the user may use the form.
        ''' </summary>
        ''' <param name="strForm"></param>
        ''' <param name="blnRO"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' do add forms to m_form if they come this way.
        ''' </remarks>
        Public Function blnCheckLevel(ByVal strForm As String, ByRef blnRO As Boolean) As Boolean
            Dim blnRet As Boolean = blnCheckLevel(strForm, blnRO, True)
            Return blnRet
        End Function

        'dont add forms to m_form if they come this way.
        Public Function blnCheckLevel(ByVal strForm As String) As Boolean
            Dim blnRO As Boolean
            Dim blnRet As Boolean = blnCheckLevel(strForm, blnRO, False)
            Return blnRet
        End Function

        '20100217 RPB IsExcel2003 always returns true.
        Public Function IsExcel2003(ByVal strConnection As String) As Boolean
            Return True
        End Function
        Public Function IsBlocked(ByVal strUser As String) As Boolean
            Dim blnRet = True
            If _blnSupportsUserLog = True Then

            End If
            Return blnRet
        End Function
#End Region

#Region "get for multilang"
        '20111104 Added version with Descr and Type when Word template function needed identify new Bookmark texts.
        Public Function get_txt_header(ByVal txt As String) As String
            Return get_txt_header(txt, "Added by application.", "Menu text")
        End Function

        Public Function get_txt_header(ByVal txt As String, ByVal strDescr As String, ByVal strType As String) As String
            Dim strHeader As String = txt
            If Not strType Is Nothing And Not strDescr Is Nothing Then
                txt = txt.Trim()

                Dim RecordAvailable = False
                Dim rowT As TheDataSet.m_txtRow = List_m_txt((txt.ToUpper()).GetHashCode())
                If Not rowT Is Nothing Then
                    RecordAvailable = True
                End If

                If RecordAvailable = True Then
                    Dim rowTH As TheDataSet.m_txt_headerRow = List_m_txt_header((txt.ToUpper()).GetHashCode())
                    If Not rowTH Is Nothing Then
                        If Not rowTH.IsheaderNull Then
                            strHeader = rowTH.header
                        End If
                    End If
                End If

                If RecordAvailable = False Then
                    Try
                        m_txtTableAdapter.Insert(txt, strDescr, strType)
                    Catch ex As Exception
                    End Try
                End If
            Else
                strHeader = get_txt_header(txt)
            End If
            Return strHeader
        End Function

        Public Function get_v_tble_column_header(ByVal strTableU As String, ByVal strColumnU As String, _
        ByRef strHeader As String) As Boolean
            Dim blnRet As Boolean = False
            'And row.lang = Lang is not needed because table is filled with only the lang of the user.
            Dim rowCH As TheDataSet.v_tble_column_headerRow = List_m_tble_column_header((strTableU + "*" + strColumnU).GetHashCode())
            If Not rowCH Is Nothing Then
                If Not rowCH.IsheaderNull Then
                    strHeader = rowCH.header
                End If
                blnRet = True
            End If
            Return blnRet
        End Function

        Public Function get_v_tble_column_format(ByVal strTableU As String, ByVal strColumnU As String, _
                ByRef strFormat As String) As Boolean
            Dim blnRet As Boolean = False
            Dim rowTC As TheDataSet.m_tble_columnRow = List_m_tble_column((strTableU + "*" + strColumnU).GetHashCode())
            If Not rowTC Is Nothing Then
                If rowTC.IsformatNull = True Then
                    strFormat = ""
                Else
                    strFormat = rowTC.format
                    blnRet = True
                End If
            End If
            Return blnRet
        End Function

        Public Function get_v_tble_column_width(ByVal strTableU As String, ByVal strColumnU As String, _
              ByRef iWidth As Integer) As Boolean
            Dim blnRet As Boolean = False
            Dim rowTC As TheDataSet.m_tble_columnRow = List_m_tble_column((strTableU + "*" + strColumnU).GetHashCode())
            If Not rowTC Is Nothing Then
                iWidth = rowTC.width
                blnRet = True
            End If
            Return blnRet
        End Function

        Public Function get_v_form_tble_column_visible(ByVal strForm As String, ByVal strTable As String, ByVal strColumn As String, _
            ByRef blnVisible As Boolean) As Boolean

            Dim strTableU = strTable.Trim.ToUpper()
            Dim strColumnU = strColumn.Trim.ToUpper()
            Dim strFormU = strForm.Trim.ToUpper()
            Dim rowL As TheDataSet.m_form_tble_column__visibilityRow = List_m_form_tble_column__visibility((strFormU + "*" + strTableU + "*" + strColumnU).GetHashCode())
            If Not rowL Is Nothing Then
                blnVisible = rowL.visible
                Return True
            Else
                blnVisible = False
            End If
            Return False
        End Function

        Public Sub get_v_form_tble_column(ByVal strForm As String, ByVal strTable As String, ByVal strColumn As String, _
        ByRef strHeader As String, ByRef strFormat As String, ByRef iWidth As Integer, _
        ByRef blnVisible As Boolean, ByRef blnPrnt As Boolean, ByRef blnBold As Boolean, _
        ByRef iSequence As Integer)

            Dim strDefault_filter As String = ""

            'Look for the header of lang. If not found use the column name.
            Dim HeaderAvailable = False
            Dim strOriginalHeader As String = strHeader
            strHeader = strColumn

            Dim strTableU = strTable.Trim.ToUpper()
            Dim strColumnU = strColumn.Trim.ToUpper()
            Dim strFormU = strForm.Trim.ToUpper()
            HeaderAvailable = get_v_tble_column_header(strTableU, strColumnU, strHeader)

            'Get table column dependencies.
            Dim RecordAvailable = False
            Dim rowTC As TheDataSet.m_tble_columnRow = List_m_tble_column((strTableU + "*" + strColumnU).GetHashCode())
            If Not rowTC Is Nothing Then
                If rowTC.IsformatNull = True Then
                    strFormat = ""
                Else
                    strFormat = rowTC.format
                End If
                iWidth = rowTC.width
                RecordAvailable = True
            End If
            strFormat = strFormat.Trim()

            'if not in the database store default values.
            If RecordAvailable = False Then
                put_v_form_tble(strForm, strTable)
                Try
                    m_tble_columnTableAdapter.Insert(strTable, strColumn, strFormat, iWidth)
                Catch ex As Exception
                End Try
            End If

            'Get form, table, column dependencies.
            RecordAvailable = False
            Dim rowL As TheDataSet.m_form_tble_column__visibilityRow = List_m_form_tble_column__visibility((strFormU + "*" + strTableU + "*" + strColumnU).GetHashCode())
            If Not rowL Is Nothing Then
                blnVisible = rowL.visible
                blnPrnt = rowL.prnt
                blnBold = rowL.bold
                If rowL.IssequenceNull Then
                    iSequence = 0
                Else
                    iSequence = rowL.sequence
                End If

                If rowL.Isdefault_filterNull Then
                    strDefault_filter = ""
                Else
                    strDefault_filter = rowL.default_filter
                    If strDefault_filter Is Nothing Then
                        strDefault_filter = ""
                    End If
                End If
                RecordAvailable = True
            End If

            'if not in the database store default values.
            If RecordAvailable = False Then
                put_v_form_tble(strForm, strTable)
                Try
                    m_form_tble_column__visibilityTableAdapter.Insert(strForm, strTable, strColumn, _
                        blnVisible, blnPrnt, iSequence, blnBold, strDefault_filter)
                Catch ex As Exception
                End Try
            End If

            'Insert if the header coming from DefineColumn doesn't equal the column name and is not defined in the data.
            'This occurs only when converting old code whare the Header was defined by hand in the code.
            If HeaderAvailable = False And strOriginalHeader.Length > 0 And strOriginalHeader.ToUpper() <> strColumnU Then
                Try
                    strHeader = strOriginalHeader
                    v_tble_column_headerTableAdapter.Insert(strTable, strColumn, Lang, strOriginalHeader)
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If

        End Sub

        Public Function strGetDefault_Filter(ByVal strForm As String, ByVal strTable As String, ByVal strColumn As String) As String
            Dim row As TheDataSet.m_form_tble_column__visibilityRow = List_m_form_tble_column__visibility((strForm.ToUpper() + "*" + strTable.ToUpper() + "*" + strColumn.ToUpper()).GetHashCode())
            Dim strDefault_filter As String = ""
            If Not row Is Nothing Then
                If Not row.default_filter Is Nothing Then
                    strDefault_filter = row.default_filter.Trim
                End If
            End If
            Return strDefault_filter
        End Function

#End Region

#Region "Put Functions"
        Public Sub put_v_form_tble(ByVal form As String, ByVal tble As String)
            If strLastForm <> form Then
                Try
                    '20120614 When DELETING Forms from Security form remember that the Forms only get created on initialisation so 
                    'it is necessary to restart the application to re-create.
                    'It cannot be done here because we do not what type of menu it is and therefore do not know the 
                    'Insert( parameters.
                    'm_formTableAdapter.Insert(form)
                Catch ex As Exception
                End Try
                strLastForm = form
            End If

            If strLastTble <> tble Then
                Try
                    m_tbleTableAdapter.Insert(tble)
                Catch ex As Exception
                End Try
                Try
                    m_form_tbleTableAdapter.Insert(form, tble)
                Catch ex As Exception
                End Try
                strLastTble = tble
            End If
        End Sub

#End Region

#Region "Private Functions"
        'Check whether table and/or form are already in the database.

        Private Function LogonUser() As IntPtr
            Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
            Return accountToken
        End Function

        Public Sub put_v_form(ByVal form As String, ByVal blnMenu As Boolean, ByVal blnMenu_entry As Boolean)

            'Check whether this is a new form.
            Dim blnNoForm = True
            Dim blnDifferent = True
            If Not m_form Is Nothing Then
                For Each row As TheDataSet.m_formRow In m_form
                    If row.form.ToUpper = form.ToUpper Then
                        blnNoForm = False
                        If row.menu = blnMenu And row.menu_entry = blnMenu_entry Then
                            blnDifferent = False
                        End If
                        Exit For
                    End If
                Next
            End If

            Try
                If blnNoForm = True Then
                    m_formTableAdapter.Insert(form, blnMenu, blnMenu_entry)
                Else
                    If blnDifferent = True Then
                        m_formTableAdapter.Update(form, blnMenu, blnMenu_entry, form)
                    End If
                End If
            Catch ex As Exception
            End Try
        End Sub

        Public Sub put_v_form_groupbox(ByVal strForm As String, ByVal strGroupbox As String)

            'Check whether this is a new form.
            Dim blnNoForm = True
            If Not v_usr_grp_groupbox Is Nothing Then
                For Each row As TheDataSet.v_usr_grp_groupboxRow In v_usr_grp_groupbox
                    If row.form.ToUpper = strForm.ToUpper And row.groupbox.ToUpper = strGroupbox.ToUpper Then
                        blnNoForm = False
                        Exit For
                    End If
                Next
            End If

            Try
                If blnNoForm = True Then
                    Dim m_usr_grp_groupboxTA = New TheDataSetTableAdapters.m_form_grp_groupboxTableAdapter
                    m_usr_grp_groupboxTA.Connection.ConnectionString = ParentForm.GetConnectionString()

                    'insert the group with read only false.
                    If _Group <> "" Then
                        m_usr_grp_groupboxTA.Insert(_Group, strForm, strGroupbox, False)
                    End If

                End If
            Catch ex As Exception
            End Try

        End Sub
#End Region

#Region "Text colour"
        Public Function GetTextColor(backc As Color) As Color

            'Dim c As Color = Color.FromArgb(backc.ToArgb Xor &HFFFFFF)
            'Return c

            Dim brightness As Integer = _
                CInt(backc.R) + _
                backc.G + _
                backc.B
            If brightness > 350 Then
                Return Color.Black
            Else
                Return Color.White
            End If

        End Function

        Public Function GetFieldColor(blnRO As Boolean) As System.Drawing.Color
            If blnRO Then
                Return ReadOnlyBackGroundColor()
            End If

            Return Drawing.Color.White
        End Function

        Public ReadOnly Property ReadOnlyBackGroundColor() As System.Drawing.Color
            Get
                Return Drawing.Color.FromKnownColor(System.Drawing.KnownColor.Control)
            End Get
        End Property

#End Region

    End Module
End Namespace

