'------------------------------------------------
'Name: Module ExcelFormat.vb.
'Function: 
'Copyright Robin Baines 2007. All rights reserved.
'Created March 2007.
'Notes: 
'Modifications: 
'------------------------------------------------
Public Class ExcelFormat
    Friend iName As ExcelStringFormats

    Private strName As String
    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
        End Set
    End Property
    Private blnBold As Boolean
    Public Property Bold() As Boolean
        Get
            Return blnBold
        End Get
        Set(ByVal value As Boolean)
            blnBold = value
        End Set
    End Property
    Private blnItalic As Boolean
    Public Property Italic() As Boolean
        Get
            Return blnItalic
        End Get
        Set(ByVal value As Boolean)
            blnItalic = value
        End Set
    End Property
    Private blnWordWrap As Boolean
    Public Property WordWrap() As Boolean
        Get
            Return blnWordWrap
        End Get
        Set(ByVal value As Boolean)
            blnWordWrap = value
        End Set
    End Property
    Private blnRightAligned As Boolean
    Public Property RightAligned() As Boolean
        Get
            Return blnRightAligned
        End Get
        Set(ByVal value As Boolean)
            blnRightAligned = value
        End Set
    End Property
    Private strColour As String
    Public Property Colour() As String
        Get
            Return strColour
        End Get
        Set(ByVal value As String)
            strColour = value
        End Set
    End Property
    Private strBorderTop As String
    Public Property BorderTop() As String
        Get
            Return strBorderTop
        End Get
        Set(ByVal value As String)
            strBorderTop = value
        End Set
    End Property
    Private strBorderLeft As String
    Public Property BorderLeft() As String
        Get
            Return strBorderLeft
        End Get
        Set(ByVal value As String)
            strBorderLeft = value
        End Set
    End Property

    Private strBorderBottom As String
    Public Property BorderBottom() As String
        Get
            Return strBorderBottom
        End Get
        Set(ByVal value As String)
            strBorderBottom = value
        End Set
    End Property
    Private strBorderRight As String
    Public Property BorderRight() As String
        Get
            Return strBorderRight
        End Get
        Set(ByVal value As String)
            strBorderRight = value
        End Set
    End Property

    Private iBorderWeight As String
    Public Property BorderWeight() As BorderWeight
        Get
            Return iBorderWeight
        End Get
        Set(ByVal value As BorderWeight)
            iBorderWeight = value
        End Set
    End Property
    Private strFormat As String
    Public Property Format() As String
        Get
            Return strFormat
        End Get
        Set(ByVal value As String)
            strFormat = value
        End Set
    End Property
    Private iFontSize As Integer
    Public Property FontSize() As Integer
        Get
            Return iFontSize
        End Get
        Set(ByVal value As Integer)
            iFontSize = value
        End Set
    End Property
    Private strUnderlined As String
    Public Property Underlined() As String
        Get
            Return strUnderlined
        End Get
        Set(ByVal value As String)
            strUnderlined = value
        End Set
    End Property

    Private strType As String
    Public Property Type() As String
        Get
            Return strType
        End Get
        Set(ByVal value As String)
            strType = value
        End Set
    End Property
    Private strFont As String
    Public Property Font() As String
        Get
            Return strFont
        End Get
        Set(ByVal value As String)
            strFont = value
        End Set
    End Property
    Private strBackGround As String
    Public Property BackGround() As String
        Get
            Return strBackGround
        End Get
        Set(ByVal value As String)
            strBackGround = value
        End Set
    End Property

    Public Sub New(ByVal _iName As ExcelStringFormats, ByVal _strName As String, ByVal _blnBold As Boolean, _
    ByVal _strColour As String, _
    ByVal _BorderTop As Boolean, _
    ByVal _BorderLeft As Boolean, _
    ByVal _BorderBottom As Boolean, _
    ByVal _BorderRight As Boolean, _
    ByVal _iBorderWeight As BorderWeight, _
    ByVal _strFormat As String, _
    ByVal _strType As String, _
    ByVal _iFontSize As Integer, _
    ByVal _strUnderlined As String, _
    ByVal _strFont As String, _
    ByVal _strBackGround As String)

        iName = _iName
        strName = _strName
        blnBold = _blnBold
        blnItalic = False
        strColour = _strColour
        BorderTop = _BorderTop
        BorderLeft = _BorderLeft
        BorderBottom = _BorderBottom
        BorderRight = _BorderRight
        iBorderWeight = _iBorderWeight
        strFormat = _strFormat
        strType = _strType
        iFontSize = _iFontSize
        strUnderlined = _strUnderlined
        strFont = _strFont
        BackGround = _strBackGround
        '  WordWrap = False
    End Sub
    
    Public Sub New(ByVal _iName As ExcelStringFormats, ByVal _strName As String, ByVal _blnBold As Boolean, _
    ByVal _blnItalic As Boolean, _
    ByVal _blnRightAligned As Boolean, _
    ByVal _strFormat As String, _
    ByVal _strType As String, _
    ByVal _iFontSize As Integer, _
    ByVal _strFont As String)

        iName = _iName
        strName = _strName
        blnBold = _blnBold
        blnRightAligned = _blnRightAligned
        blnItalic = _blnItalic
        strColour = ""
        BorderTop = False
        BorderLeft = False
        BorderBottom = False
        BorderRight = False
        iBorderWeight = BorderWeight.None
        strFormat = _strFormat
        strType = _strType
        iFontSize = _iFontSize
        strUnderlined = ""
        strFont = _strFont
        BackGround = ""

        ' WordWrap = False
    End Sub
    Public Sub New(ByVal _iName As ExcelStringFormats, ByVal _strName As String, ByVal _blnBold As Boolean, _
    ByVal _blnItalic As Boolean, _
    ByVal _blnRightAligned As Boolean, _
    ByVal _strFormat As String, _
    ByVal _strType As String, _
    ByVal _iFontSize As Integer, _
    ByVal _strFont As String, _
    ByVal _blnWordWrap As Boolean)

        iName = _iName
        strName = _strName
        blnBold = _blnBold
        blnRightAligned = _blnRightAligned
        blnItalic = _blnItalic
        strColour = ""
        BorderTop = False
        BorderLeft = False
        BorderBottom = False
        BorderRight = False
        iBorderWeight = BorderWeight.None
        strFormat = _strFormat
        strType = _strType
        iFontSize = _iFontSize
        strUnderlined = ""
        strFont = _strFont
        BackGround = ""
        blnWordWrap = _blnWordWrap
    End Sub

    Friend XMLExcelInterface
End Class
