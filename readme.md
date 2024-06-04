# Inhalt  
Die in diesem Repository enthaltenen Dateien stehen in Zusammenhang mit meiner Bachelorarbeit zum Thema "Entwicklung einer Lösung zur bidirektionalen Outlook-Synchronisation mit einem Terminbuchungssystem".  
Der gesamte Quellcode ist für sich alleine nicht lauffähig, da er im Umfeld des SmartCX-Terminbuchungssystems entwickelt wurde.  
Auch sind alle Dateien, die Quellcode der Firma SmartCJM enthielten auf meinen für die Bachelorarbeit selbst verfassten Quellcode reduziert.  
  
- Die [OutlookHelper.cs](/OutlookHelper.cs) enthält die Implementation der Methoden, die für die direkte Kommunikation mit Microsofts Graph bzw. mit SmartCX notwendig sind.  
- Die [RegisterOrRenewSubscription.cs](/RegisterOrRenewSubscription.cs) enthält die Methode zum Erstellen oder Erneuern von Abonnements.  
- Die [BackgroundTaskJob.cs](/BackgroundTaskJob.cs) enthält den Backgroundjob zum asynchronen Verarbeiten der Abonnementbenachrichtigungen.  
- Und die [LiquidBeispiel.html](/LiquidBeispiel.html) enthält den Beispielcode zur ausgabe von Termindaten.
