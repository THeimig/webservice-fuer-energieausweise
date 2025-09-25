#Region "Imports"

#End Region
'----------------------------------------------------------------------------------
Public Class Klasse_Webservice
    '----------------------------------------------------------------------------------
#Region "Variablen"
    '----------------------------------------------------------------------------------
    Dim DLL_ID_DIBT As String
    Dim DLL_PWD_DIBT As String
    Dim DLL_Ausstellungsdatum As Date
    Dim DLL_Bundesland As String
    Dim DLL_Postleitzahl As String
    Dim DLL_Gesetzesgrundlage As String
    Dim DLL_Gebaeudeart As String
    Dim DLL_Berechnungsart As String
    Dim DLL_Neubau As Integer
    Dim DLL_Registriernummer As String
    Dim DLL_Kontrolldatei As String
    Dim DLL_Sandbox As Boolean
    '----------------------------------------------------------------------------------
    Public DLL_Ergebnis_Datenregistratur As String
    '----------------------------------------------------------------------------------
    Public DLL_Ergebnis_KontrolldateiPruefen As Boolean
    Public DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Anzahl As Integer
    Public DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_ID(200) As String
    Public DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Kurztext(200) As String
    Public DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Langtext(200) As String
    '----------------------------------------------------------------------------------
    Public DLL_Ergebnis_Restkontingent_Anzahl As Integer
    Public DLL_Ergebnis_Restkontingent_Fehlermeldung As String
    '----------------------------------------------------------------------------------
    Public DLL_Ergebnis_ZusatzdatenErfassung As String
    '----------------------------------------------------------------------------------
    Public DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl As Integer
    Public DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Registriernummer(1000) As String
    Public DLL_Ergebnis_OffeneKontrolldateien_Ausweis_NummerErzeugtAm(1000) As Date
    Public DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Aussteller(1000) As String
    Public DLL_Ergebnis_OffeneKontrolldateien_Fehler As String
    '----------------------------------------------------------------------------------
#End Region
    '----------------------------------------------------------------------------------
#Region "New"
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieses "New" wird angesteuert wenn der Webservice eine Regriernummer anfordert.
    ''' </summary>
    Sub New(ByVal ID_DIBT As String, ByVal PWD_DIBT As String, ByVal Ausstellungsdatum As Date, ByVal Bundesland As String, ByVal Postleitzahl As String, ByVal Gesetzesgrundlage As String, ByVal Gebaeudeart As String, ByVal Berechnungsart As String, ByVal Neubau As Integer, ByVal Sandbox As Boolean)
        '----------------------------------------------------------------------------------
        DLL_ID_DIBT = ID_DIBT
        DLL_PWD_DIBT = PWD_DIBT
        DLL_Ausstellungsdatum = Ausstellungsdatum
        DLL_Bundesland = Bundesland
        DLL_Postleitzahl = Postleitzahl
        DLL_Gesetzesgrundlage = Gesetzesgrundlage
        DLL_Gebaeudeart = Gebaeudeart
        DLL_Berechnungsart = Berechnungsart
        DLL_Neubau = Neubau
        DLL_Sandbox = Sandbox
        '----------------------------------------------------------------------------------
        If DLL_Gesetzesgrundlage = "ENEV-2014" Or DLL_Gesetzesgrundlage = "ENEV-2016" Then
            DLL_Gesetzesgrundlage = "ENEV"
        End If
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieses "New" wird angesteuert wenn die Kontrolldatei überprüft werden soll und beim Hochladen der Kontrolldatei.
    ''' </summary>
    Sub New(ByVal ID_DIBT As String, ByVal PWD_DIBT As String, ByVal Registriernummer As String, ByVal Kontrolldatei As String, ByVal Sandbox As Boolean)
        '----------------------------------------------------------------------------------
        DLL_ID_DIBT = ID_DIBT
        DLL_PWD_DIBT = PWD_DIBT
        DLL_Registriernummer = Registriernummer
        DLL_Kontrolldatei = Kontrolldatei
        DLL_Sandbox = Sandbox
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' Dieses "New" wird angesteuert wenn das Restkontingent abgefragt werden soll.
    ''' </summary>
    Sub New(ByVal ID_DIBT As String, ByVal PWD_DIBT As String, ByVal Sandbox As Boolean)
        '----------------------------------------------------------------------------------
        DLL_ID_DIBT = ID_DIBT
        DLL_PWD_DIBT = PWD_DIBT
        DLL_Sandbox = Sandbox
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
#End Region
    '----------------------------------------------------------------------------------
#Region "Webservice Abfragen"
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Webservice angesteuert und eine Regriernummer angefordert.
    ''' </summary>
    Sub Webservice_Datenregistratur()
        '----------------------------------------------------------------------------------
        Try
            '----------------------------------------------------------------------------------
            Dim Inhalt_XML As String = ""
            '----------------------------------------------------------------------------------
            Inhalt_XML &= "<?xml version='1.0' encoding='UTF-8'?>"
            Inhalt_XML &= "<root xmlns = 'https://energieausweis.dibt.de/schema/SchemaDatenErfassungGEG.xsd' >"
            '----------------------------------------------------------------------------------
            Inhalt_XML &= "<Authentifizierung>"
            Inhalt_XML &= "<Aussteller_ID_DIBT>" & DLL_ID_DIBT & "</Aussteller_ID_DIBT>"
            Inhalt_XML &= "<Aussteller_PWD_DIBT>" & DLL_PWD_DIBT & "</Aussteller_PWD_DIBT>"
            Inhalt_XML &= "</Authentifizierung>"
            '----------------------------------------------------------------------------------
            Inhalt_XML &= "<Nachweis-Daten>"
            Inhalt_XML &= "<Ausstellungsdatum>" & DLL_Ausstellungsdatum & "</Ausstellungsdatum>"
            Inhalt_XML &= "<Bundesland>" & DLL_Bundesland & "</Bundesland>"
            Inhalt_XML &= "<Postleitzahl>" & DLL_Postleitzahl & "</Postleitzahl>"
            Inhalt_XML &= "<Gesetzesgrundlage>" & DLL_Gesetzesgrundlage & "</Gesetzesgrundlage>"
            Inhalt_XML &= "</Nachweis-Daten>"
            '----------------------------------------------------------------------------------
            Inhalt_XML &= "<Energieausweis-Daten>"
            Inhalt_XML &= "<Gebaeudeart>" & DLL_Gebaeudeart & "</Gebaeudeart>"
            Inhalt_XML &= "<Art>" & DLL_Berechnungsart & "</Art>"
            Inhalt_XML &= "<Neubau>" & DLL_Neubau & "</Neubau>"
            Inhalt_XML &= "</Energieausweis-Daten>"
            '----------------------------------------------------------------------------------
            Inhalt_XML &= "</root>"
            '----------------------------------------------------------------------------------
            Dim xml As New XmlDocument()
            Dim xmlresult As New XmlDocument()
            '----------------------------------------------------------------------------------
            xml.LoadXml(Inhalt_XML)
            '----------------------------------------------------------------------------------
            Dim xe As XmlElement = xmlresult.CreateElement("root")
            '----------------------------------------------------------------------------------
            If DLL_Sandbox = True Then
                '----------------------------------------------------------------------------------
                Dim Sandbox As New ServiceReferenceSandbox.DibtEnergieAusweisServiceSoapClient()
                xe.InnerXml = (DirectCast(Sandbox.Datenregistratur(xml.DocumentElement), XmlElement)).InnerXml
                '----------------------------------------------------------------------------------
            Else
                '----------------------------------------------------------------------------------
                Dim Published As New ServiceReferencePublished.DibtEnergieAusweisServiceSoapClient()
                xe.InnerXml = (DirectCast(Published.Datenregistratur(xml.DocumentElement), XmlElement)).InnerXml
                '----------------------------------------------------------------------------------
            End If
            '----------------------------------------------------------------------------------
            xmlresult.AppendChild(xe)
            '----------------------------------------------------------------------------------
            DLL_Ergebnis_Datenregistratur = xmlresult.InnerXml
            '----------------------------------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Webservice angesteuert um das Restkontingent abzufragen.
    ''' </summary>
    Sub Webservice_Restkontingent()
        Try
            '----------------------------------------------------------------------------------
            If DLL_Sandbox = True Then
                '----------------------------------------------------------------------------------
                Dim Sandbox As New ServiceReferenceSandbox.DibtEnergieAusweisServiceSoapClient()
                Dim Antwort As ServiceReferenceSandbox.AntwortRestkontingent
                '----------------------------------------------------------------------------------
                Antwort = Sandbox.Restkontingent(DLL_ID_DIBT, DLL_PWD_DIBT)
                '----------------------------------------------------------------------------------
                DLL_Ergebnis_Restkontingent_Anzahl = Antwort.Kontingent
                DLL_Ergebnis_Restkontingent_Fehlermeldung = Antwort.Fehler
                '----------------------------------------------------------------------------------
            Else
                '----------------------------------------------------------------------------------
                Dim Published As New ServiceReferencePublished.DibtEnergieAusweisServiceSoapClient()
                Dim Antwort As ServiceReferencePublished.AntwortRestkontingent
                '----------------------------------------------------------------------------------
                Antwort = Published.Restkontingent(DLL_ID_DIBT, DLL_PWD_DIBT)
                '----------------------------------------------------------------------------------
                DLL_Ergebnis_Restkontingent_Anzahl = Antwort.Kontingent
                DLL_Ergebnis_Restkontingent_Fehlermeldung = Antwort.Fehler
                '----------------------------------------------------------------------------------
            End If
            '----------------------------------------------------------------------------------
        Catch ex As Exception
            '----------------------------------------------------------------------------------
            Fehlerfenster(ex)
            '----------------------------------------------------------------------------------
        End Try

    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Webservice angesteuert und die Kontrolldatei wird geprüft.
    ''' </summary>
    Sub Webservice_KontrolldateiPruefen()
        '----------------------------------------------------------------------------------
        Try
            '----------------------------------------------------------------------------------
            Dim Kontrolldatei_XML As New XmlDocument()
            Kontrolldatei_XML.Load(DLL_Kontrolldatei)
            '----------------------------------------------------------------------------------
            If DLL_Sandbox = True Then
                '----------------------------------------------------------------------------------
                Dim Sandbox As New ServiceReferenceSandbox.DibtEnergieAusweisServiceSoapClient()
                Dim Fehlerliste As ServiceReferenceSandbox.Evaluation_Result_List
                '----------------------------------------------------------------------------------
                Fehlerliste = Sandbox.KontrolldateiPruefen(DLL_ID_DIBT, DLL_PWD_DIBT, DLL_Registriernummer, Kontrolldatei_XML.DocumentElement, True, False)
                '----------------------------------------------------------------------------------
                DLL_Ergebnis_KontrolldateiPruefen = Fehlerliste.Errors_detected
                '----------------------------------------------------------------------------------
                DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Anzahl = Fehlerliste.Evaluation_List.Count
                '----------------------------------------------------------------------------------
                For WertX = 0 To (Fehlerliste.Evaluation_List.Count - 1)
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_ID(WertX) = Fehlerliste.Evaluation_List(WertX).Key
                    DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Kurztext(WertX) = Fehlerliste.Evaluation_List(WertX).Error_Warning
                    DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Langtext(WertX) = Fehlerliste.Evaluation_List(WertX).Result
                    '----------------------------------------------------------------------------------
                Next
                '----------------------------------------------------------------------------------
            Else
                '----------------------------------------------------------------------------------
                Dim Published As New ServiceReferencePublished.DibtEnergieAusweisServiceSoapClient()
                Dim Fehlerliste As ServiceReferencePublished.Evaluation_Result_List
                '----------------------------------------------------------------------------------
                Fehlerliste = Published.KontrolldateiPruefen(DLL_ID_DIBT, DLL_PWD_DIBT, DLL_Registriernummer, Kontrolldatei_XML.DocumentElement, True, True)
                '----------------------------------------------------------------------------------
                DLL_Ergebnis_KontrolldateiPruefen = Fehlerliste.Errors_detected
                '----------------------------------------------------------------------------------
                DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Anzahl = Fehlerliste.Evaluation_List.Count
                '----------------------------------------------------------------------------------
                For WertX = 0 To (Fehlerliste.Evaluation_List.Count - 1)
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_ID(WertX) = Fehlerliste.Evaluation_List(WertX).Key
                    DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Kurztext(WertX) = Fehlerliste.Evaluation_List(WertX).Error_Warning
                    DLL_Ergebnis_KontrolldateiPruefen_Fehlerliste_Langtext(WertX) = Fehlerliste.Evaluation_List(WertX).Result
                    '----------------------------------------------------------------------------------
                Next
                '----------------------------------------------------------------------------------
            End If
            '----------------------------------------------------------------------------------
        Catch ex As Exception
            Fehlerfenster(ex)
        End Try
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Webservice angesteuert um die Zusatzdaten beim DIBt Server hochzuladen.
    ''' </summary>
    Sub Webservice_ZusatzdatenErfassung()
        '----------------------------------------------------------------------------------
        Try
            '----------------------------------------------------------------------------------
            Dim XML As New XmlDocument()
            Dim xmlresult As New XmlDocument()
            '----------------------------------------------------------------------------------
            XML.Load(DLL_Kontrolldatei)
            '----------------------------------------------------------------------------------
            Dim xe As XmlElement = xmlresult.CreateElement("root")
            '----------------------------------------------------------------------------------
            If DLL_Sandbox = True Then
                '----------------------------------------------------------------------------------
                Dim Sandbox As New ServiceReferenceSandbox.DibtEnergieAusweisServiceSoapClient()
                xe.InnerXml = (DirectCast(Sandbox.ZusatzdatenErfassung(XML.DocumentElement, DLL_Registriernummer, DLL_ID_DIBT, DLL_PWD_DIBT), XmlElement)).InnerXml
                '----------------------------------------------------------------------------------
            Else
                '----------------------------------------------------------------------------------
                Dim Published As New ServiceReferencePublished.DibtEnergieAusweisServiceSoapClient()
                xe.InnerXml = (DirectCast(Published.ZusatzdatenErfassung(XML.DocumentElement, DLL_Registriernummer, DLL_ID_DIBT, DLL_PWD_DIBT), XmlElement)).InnerXml
                '----------------------------------------------------------------------------------
            End If
            '----------------------------------------------------------------------------------
            xmlresult.AppendChild(xe)
            '----------------------------------------------------------------------------------
            DLL_Ergebnis_ZusatzdatenErfassung = xmlresult.InnerXml
            '----------------------------------------------------------------------------------
        Catch ex As Exception
            '----------------------------------------------------------------------------------
            Fehlerfenster(ex)
            '----------------------------------------------------------------------------------
        End Try
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' In dieser Anweisung wird der Webservice angesteuert um die Liste der offenen Kontrolldateien abzufragen.
    ''' </summary>
    Sub Webservice_OffeneKontrolldateien()
        '----------------------------------------------------------------------------------
        Try
            '----------------------------------------------------------------------------------
            If DLL_Sandbox = True Then
                '----------------------------------------------------------------------------------
                Dim Sandbox As New ServiceReferenceSandbox.DibtEnergieAusweisServiceSoapClient()
                Dim Antwort As ServiceReferenceSandbox.AntwortOffeneKontrolldateien
                Dim Ausweise(1000) As ServiceReferenceSandbox.Ausweis
                '----------------------------------------------------------------------------------
                Antwort = Sandbox.OffeneKontrolldateien(DLL_ID_DIBT, DLL_PWD_DIBT, False)
                '----------------------------------------------------------------------------------
                If Antwort.Ausweise IsNot Nothing Then
                    '----------------------------------------------------------------------------------
                    Ausweise = Antwort.Ausweise
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl = Ausweise.Count
                    '----------------------------------------------------------------------------------
                    For WertX = 0 To Ausweise.Count - 1
                        '----------------------------------------------------------------------------------
                        DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Registriernummer(WertX) = Ausweise(WertX).Registriernummer
                        DLL_Ergebnis_OffeneKontrolldateien_Ausweis_NummerErzeugtAm(WertX) = Ausweise(WertX).NummerErzeugtAm
                        DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Aussteller(WertX) = Ausweise(WertX).Aussteller
                        '----------------------------------------------------------------------------------
                    Next
                    '----------------------------------------------------------------------------------
                Else
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_OffeneKontrolldateien_Fehler = Antwort.Fehler
                    '----------------------------------------------------------------------------------
                    If DLL_Ergebnis_OffeneKontrolldateien_Fehler = "" Then DLL_Ergebnis_OffeneKontrolldateien_Fehler = "Es liegt ein Fehler vor. Es konnten keine Daten geladen werden."
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl = 0
                    '----------------------------------------------------------------------------------
                End If
                '----------------------------------------------------------------------------------
            Else
                '----------------------------------------------------------------------------------
                Dim Published As New ServiceReferencePublished.DibtEnergieAusweisServiceSoapClient()
                Dim Antwort As ServiceReferencePublished.AntwortOffeneKontrolldateien
                Dim Ausweise(1000) As ServiceReferencePublished.Ausweis
                '----------------------------------------------------------------------------------
                Antwort = Published.OffeneKontrolldateien(DLL_ID_DIBT, DLL_PWD_DIBT, False)
                '----------------------------------------------------------------------------------
                If Antwort.Ausweise IsNot Nothing Then
                    '----------------------------------------------------------------------------------
                    Ausweise = Antwort.Ausweise
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl = Ausweise.Count
                    '----------------------------------------------------------------------------------
                    For WertX = 0 To Ausweise.Count - 1
                        '----------------------------------------------------------------------------------
                        DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Registriernummer(WertX) = Ausweise(WertX).Registriernummer
                        DLL_Ergebnis_OffeneKontrolldateien_Ausweis_NummerErzeugtAm(WertX) = Ausweise(WertX).NummerErzeugtAm
                        DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Aussteller(WertX) = Ausweise(WertX).Aussteller
                        '----------------------------------------------------------------------------------
                    Next
                    '----------------------------------------------------------------------------------
                Else
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_OffeneKontrolldateien_Fehler = Antwort.Fehler
                    '----------------------------------------------------------------------------------
                    If DLL_Ergebnis_OffeneKontrolldateien_Fehler = "" Then DLL_Ergebnis_OffeneKontrolldateien_Fehler = "Es liegt ein Fehler vor. Es konnten keine Daten geladen werden."
                    '----------------------------------------------------------------------------------
                    DLL_Ergebnis_OffeneKontrolldateien_Ausweis_Anzahl = 0
                    '----------------------------------------------------------------------------------
                End If
                '----------------------------------------------------------------------------------
            End If
            '----------------------------------------------------------------------------------
        Catch ex As Exception
            '----------------------------------------------------------------------------------
            Fehlerfenster(ex)
            '----------------------------------------------------------------------------------
        End Try
        '----------------------------------------------------------------------------------
    End Sub
    '----------------------------------------------------------------------------------
#End Region
    '----------------------------------------------------------------------------------
#Region "Fehlertext"
    '----------------------------------------------------------------------------------
    Function Fehlerfenster(ByVal ex As Exception)

        Dim sError As String = "Unerwarteter Anwendungsfehler: " & vbCrLf & ex.Message.ToString & vbCrLf & vbCrLf & "Source: " & ex.StackTrace.ToString & vbCrLf & vbCrLf & "Bitte kontaktieren Sie den Entwickler unter (https://www.energieausweis.support)"

        MsgBox(sError, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Fehler...")

        Fehlerfenster = sError

    End Function
    '----------------------------------------------------------------------------------
#End Region
    '----------------------------------------------------------------------------------
End Class
