Option Strict On
Imports System.Collections.Generic


Public Class Helper_FeierTage
    Private Shared ReadOnly feiertageList As New List(Of Helper_FeierTag)()
    Public Structure HolidayDef
        Friend isHoliday As Boolean
        Friend Datum As Date
        Friend Country As Land()
        Friend Feiertag As String
        Sub New(Optional ByVal _isHoliday As Boolean = False,
                       Optional ByVal _Datum As Date = Nothing,
                       Optional ByVal _Country As Land() = Nothing,
                       Optional ByVal _Feiertag As String = "")
            isHoliday = _isHoliday
            Datum = _Datum
            Country = _Country
            Feiertag = _Feiertag
        End Sub
    End Structure
    Public Enum Land
        Baden_Würtenberg
        Bayern
        Berlin
        Brandenburg
        Bremen
        Hamburg
        Hessen
        Mecklenburg_Vorpommern
        Niedersachsen
        Nordrhein_Westfalen
        Rheinland_Pfalz
        Saarland
        Sachsen
        Sachsen_Anhalt
        Schleswig_Holstein
        Thüringen
    End Enum
    Public Enum FeiertagsArt
        Fester_Feiertag
        Bewegliche_Feiertag
    End Enum
    Public Structure TypeFeiertag
        Public NameA As String
        Public DatumA As Date
        Public arbeitsfrei As Boolean
    End Structure

    'Public Shared Sub New()
    '    Initial()
    'End Sub

    Public Shared Function GetFeiertag(datum As DateTime, bundesland As Land) As String
        ' Liste der Feiertage durchgehen
        For Each f As Helper_FeierTag In feiertageList
            If datum.ToShortDateString().Equals(f.GetDatum(GetOstersonntag(datum.Year)).ToShortDateString()) Then
                ' Prüfen ob das Land enthalten ist
                For Each l As Land In f.Laender
                    If bundesland = l Then
                        Return f.Feiertag
                    End If
                Next
            End If
        Next
        Return ""
    End Function
    Public Shared Function GetFeiertagAsHolidayDef(datum As DateTime, bundesland As Land) As HolidayDef    'As [String]
        Dim Holiday As New HolidayDef
        ' Liste der Feiertage durchgehen
        For Each f As Helper_FeierTag In feiertageList
            If datum.ToShortDateString().Equals(f.GetDatum(GetOstersonntag(datum.Year)).ToShortDateString()) Then
                ' Prüfen ob das Land enthalten ist
                For Each l As Land In f.Laender
                    If bundesland = l Then
                        Holiday.Datum = f.Datum
                        Holiday.Feiertag = f.Feiertag
                        Holiday.Country = f.Laender
                        Holiday.isHoliday = True
                        Return Holiday
                        'Return f.Feiertag
                    End If
                Next
            End If
        Next
        Return Holiday
    End Function

    Public Shared Function IsFeiertag(dateval As DateTime, Bundesland As Land) As Boolean
        Return GetFeiertag(dateval, Bundesland).Length > 0
    End Function
    Public Shared Function IsFeiertag(dateval As DateTime, Bundesland As String) As Boolean
        Dim nFeierTageLand As Helper_FeierTage.Land
        nFeierTageLand = CType(System.Enum.Parse(nFeierTageLand.GetType(), Bundesland), Helper_FeierTage.Land)
        Return GetFeiertag(dateval, nFeierTageLand).Length > 0
    End Function
    Public Shared Function IsFeiertagAsHolidayDef(dateval As DateTime, Bundesland As Land) As HolidayDef
        Dim Holiday As New HolidayDef
        Holiday = GetFeiertagAsHolidayDef(dateval, Bundesland)
        Return Holiday
    End Function

    Public Shared Function GetOstersonntag(jahr As Double) As DateTime
        Dim c As Double
        Dim i As Double
        Dim j As Double
        Dim k As Double
        Dim l As Double
        Dim n As Double
        Dim OsterTag As Double
        Dim OsterMonat As Double

        c = jahr / 100
        n = jahr - 19 * Helper_VarConvert.ConvertToDouble(jahr / 19, 0)
        k = (c - 17) / 25
        i = c - c / 4 - (Helper_VarConvert.ConvertToDouble(c - k, 0) / 3) + 19 * n + 15
        i = i - 30 * Helper_VarConvert.ConvertToDouble(i / 30, 0)
        i = i - (i / 28) * (Helper_VarConvert.ConvertToDouble(1 - (i / 28)) * Helper_VarConvert.ConvertToDouble(29 / (i + 1)) * (Helper_VarConvert.ConvertToDouble(21 - n, 0) / 11))
        j = jahr + (Helper_VarConvert.ConvertToDouble(jahr, 0) / 4) + i + 2 - c + (Helper_VarConvert.ConvertToDouble(c, 0) / 4)
        j = j - 7 * Helper_VarConvert.ConvertToDouble(j / 7, 0)
        l = i - j

        OsterMonat = 3 + (Helper_VarConvert.ConvertToDouble(l + 40, 0) / 44)
        OsterTag = l + 28 - 31 * (Helper_VarConvert.ConvertToDouble(OsterMonat, 0) / 4)

        Return System.Convert.ToDateTime(Helper_VarConvert.ConvertToInteger(jahr).ToString & "-" & Helper_VarConvert.ConvertToInteger(OsterMonat).ToString & "-" & Helper_VarConvert.ConvertToInteger(OsterTag).ToString)
    End Function


    Public Shared Sub Initial()
        Dim alle As Land() = {Land.Baden_Würtenberg, Land.Bayern, Land.Berlin, Land.Brandenburg, Land.Bremen, Land.Hamburg,
            Land.Hessen, Land.Mecklenburg_Vorpommern, Land.Niedersachsen, Land.Nordrhein_Westfalen, Land.Rheinland_Pfalz, Land.Saarland,
            Land.Sachsen, Land.Sachsen_Anhalt, Land.Schleswig_Holstein, Land.Thüringen}

        feiertageList.Add(New Helper_FeierTag("Neujahr", "01.01", FeiertagsArt.Fester_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Heiligen Drei Könige", "06.01", FeiertagsArt.Fester_Feiertag, {Land.Baden_Würtenberg, Land.Bayern, Land.Sachsen_Anhalt}))
        feiertageList.Add(New Helper_FeierTag("Karfreitag", -2, FeiertagsArt.Bewegliche_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Ostersonntag", 0, FeiertagsArt.Bewegliche_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Ostermontag", 1, FeiertagsArt.Bewegliche_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Tag der Arbeit", "01.05", FeiertagsArt.Fester_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Christi Himmelfahrt", 39, FeiertagsArt.Bewegliche_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Pfingstsonntag", 49, FeiertagsArt.Bewegliche_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Pfingstmontag", 50, FeiertagsArt.Bewegliche_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Fronleichnam", 60, FeiertagsArt.Bewegliche_Feiertag, {Land.Baden_Würtenberg, Land.Bayern, Land.Hessen, Land.Nordrhein_Westfalen, Land.Rheinland_Pfalz, Land.Saarland}))
        feiertageList.Add(New Helper_FeierTag("Mariä Himmelfahrt", "15.08", FeiertagsArt.Fester_Feiertag, {Land.Saarland}))
        feiertageList.Add(New Helper_FeierTag("Tag der dt. Einheit", "03.10", FeiertagsArt.Fester_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("Allerheiligen", "01.11", FeiertagsArt.Fester_Feiertag, {Land.Baden_Würtenberg, Land.Bayern, Land.Nordrhein_Westfalen, Land.Rheinland_Pfalz, Land.Saarland}))
        feiertageList.Add(New Helper_FeierTag("1. Weinachtstag", "25.12", FeiertagsArt.Fester_Feiertag, alle))
        feiertageList.Add(New Helper_FeierTag("2. Weinachtstag", "26.12", FeiertagsArt.Fester_Feiertag, alle))

    End Sub
End Class

Public Class Helper_FeierTag
    Private _art As Helper_FeierTage.FeiertagsArt
    Public Property art As Helper_FeierTage.FeiertagsArt
        Get
            Return _art
        End Get
        Set(ByVal value As Helper_FeierTage.FeiertagsArt)
            _art = value
        End Set
    End Property
    Private _tageHinzu As Integer
    Public Property tageHinzu As Integer
        Get
            Return _tageHinzu
        End Get
        Set(ByVal value As Integer)
            _tageHinzu = value
        End Set
    End Property
    Private _testDatum As String
    Public Property testDatum As String
        Get
            Return _testDatum
        End Get
        Set(ByVal value As String)
            _testDatum = value
        End Set
    End Property
    Private m_feiertag As String
    Public Property Feiertag As String
        Get
            Return m_feiertag
        End Get
        Set(ByVal value As String)
            m_feiertag = value
        End Set

    End Property
    Private m_datum As New DateTime
    Public Property Datum As DateTime
        Get
            Return m_datum
        End Get
        Set(ByVal value As DateTime)
            m_datum = value
        End Set
    End Property
    Private m_laender As Helper_FeierTage.Land()
    Public Property Laender As Helper_FeierTage.Land()
        Get
            Return m_laender
        End Get
        Set(ByVal value As Helper_FeierTage.Land())
            m_laender = value
        End Set
    End Property

    Sub New(ByVal feiertag__1 As String, ByVal ftestDatum As String, ByVal fart As Helper_FeierTage.FeiertagsArt, ByVal länder As Helper_FeierTage.Land())
        m_feiertag = feiertag__1
        _testDatum = ftestDatum
        _tageHinzu = 0
        _art = fart
        m_laender = länder
    End Sub
    Sub New(ByVal feiertag__1 As String, ByVal ftageHinzu As Integer, ByVal fart As Helper_FeierTage.FeiertagsArt, ByVal länder As Helper_FeierTage.Land())
        m_feiertag = feiertag__1
        _tageHinzu = ftageHinzu
        _art = fart
        m_laender = länder
    End Sub

    Public Function GetDatum(ByVal osterSonntag As DateTime) As DateTime
        If art <> Helper_FeierTage.FeiertagsArt.Fester_Feiertag Then
            m_datum = osterSonntag.AddDays(_tageHinzu)
        Else
            m_datum = DateTime.Parse(_testDatum & "." & osterSonntag.Year)
        End If

        Return DateTime.Parse(osterSonntag.Year & "-" & m_datum.Month.ToString() & "-" & m_datum.Day.ToString())
    End Function
End Class

