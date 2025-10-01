
# BBSR Energieausweis-Druckapplikation
Status: Aktiv  
Lizenz: GPLv3 Lizenz  
Sprachen: VB Net  
Entwickeler: Visionworld GmbH (im Auftrag des BBSR)  

## Über dieses Projekt
Die **BBSR Energieausweis-Druckapplikation** ist ein kostenfrei verfügbares Softwareinstrument zur **Erstellung von Energieausweisen** gemäß GEG
Sie kann als **eigenständige Anwendung** (z.B. für Verbrauchsausweise) oder in Verbindung mit **kommerzieller Berechnungssoftware** (z.B. fürBedarfsausweise) genutzt werden.
Die Druckapplikation des BBSR wurde 2014 als kostenfrei verfügbares DV-Instrument zur Erstellung von Energieausweisen nach der EnEV 2013 geschaffen, sowohl als eigenständige Anwendung z.B. für Verbrauchsausweise oder auch in Verbindung mit kommerzieller Berechnungssoftware für z.B.
Bedarfsausweise.
Von 2014 bis 2024 gab es die Druckapplikation, umgesetzt und gepflegt von der Firma LMIS in Java. Aufgrund des stetig gewachsenen und in die Jahre gekommenen Quelltextes wurde eine Auffrischung, in Form einer Neuprogrammierung, notwendig. Seit 2024 bietet das BBSR deshalb eine neue
Druckapplikation an, umgesetzt und gepflegt von der Firma Visionworld GmbH in VB.Net (IDE: Visual Studio 2019/2022 Professionell).
Die neue Druckapplikation wird hier nun als Open Source Produkt angeboten.
Ziel ist es, die Anwendung einer breiten Nutzergruppe zugänglich zu machen und der Softwarebranche eine Grundlage für eine eigenständige Weiterentwicklung der Druckapplikation zu geben.

## Technische Hintergrundinformationen

### Die neue Generation (seit 2024)
Aufgrund des stetig gewachsenen und in die Jahre gekommenen Quelltextes der ursprünglichen Java-Implementierung wurde eine **Neuprogrammierung** notwendig.
- Entwickelt in: Visual Studio 2019/2022 Professionell
- Umsetzung: Visionworld GmbH (im Auftrag des BBSR)
Diese Version ist die aktuelle Open Source Basis der Druckapplikation.

### Historie
- 2014 – 2024: Die erste Version der Druckapplikation, umgesetzt und gepflegt von der Firma LMIS in Java.
- 2024 – 2025: Neuentwicklung der Druckapplikation, umgesetzt und gepflegt von der Firma Visionworld GmbH in VB.Net

## Einbindung in das Kontrollsystem (DIBt)
Die Druckapplikation spielt eine zentrale Rolle in der Kommunikation mit dem **Webserver des DIBt (Deutsches Institut für Bautechnik)**, über den das EU-rechtlich obligatorische Kontrollsystem für Energieausweise organisiert wird.

### Datenaustausch
- Die Applikation erzeugt und übermittelt **Kontrolldatensätze** nach einem vom DIBt vorgegebenen **XML-Schema**.
- Dieses **XML-Schema** dient auch als **Kommunikationsschnittstelle** zwischen kommerzieller Berechnungssoftware und der Druckapplikation, wenn die Energieausweise im Verbund erstellt werden.

## Technische Informationen
Das Handbuch finden Sie unter dem nachfolgenden Link: [Handbuch](BBSR-Druckapplikation_Handbuch.pdf)  
Die Dokumentation finden sie unter dem nachfolgenden Link: [Dokumentation](BBSR-Druckapplikation_Dokumentation.pdf)

## Mitmachen und Entwicklung
Wir freuen uns über jeden Beitrag zur Weiterentwicklung, Fehlerbehebung und Optimierung dieser Anwendung!