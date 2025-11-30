# Projektbeschreibung: Kanzlei Hartmut Ahting Zeiterfassug

## Ausgangssituation

s
## Zielsetzung

## Umsetzungsidee


## Vorteile


## TECHNIK
## Thema: Ribbon-Menü, das ausschließlich in der Steuerzentrale.xlsm erscheint und sonst nirgends

### Problem
Wenn man ein neues Ribbon-Menü über *Menüband anpassen* anlegt, wird dies global von Excel gespeichert außerhalb der aktuell geöffneten .xlsm Datei. Das ist störend, wenn man andere Excel-Dateien öffnet, in denen man das Ribbon nicht braucht oder ein rerenziertes Makro sogar nicht vorhanden ist.

### Lösung
Es ist möglich, ein dateispezifisches Ribbon anzulegen, das nur erscheint, wenn man eine bestimmte Datei öffnet.
Dies ist jedoch nur mit einem bei github verfügbaren Zusatzool möglich, das als ZIP herunteladbar ist. Die Originalversion stammt von Microsoft, jedoch gibt es einen weiterentwickelten Fork. Der Link dazu lautet: https://github.com/fernandreu/office-ribbonx-editor/releases/latest

**Vorgehensweise:**
- RibbonXEditor starten
- Eigene .xlsm Datei öffnen
- Rechtsklick auf die Datei > "Insert Office 2007 Custom UI Part" (oder ähnlich).
- Dieses XML einfügen (Beispiel, hier für 2 Buttons):

```XML
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab id="LMVToolsId" label="LVM Tools">
        <group id="SteuerzentraleId" label="Steuerzentrale">
          <button id="ImportXMLButton" label="Import Vertrag" imageMso="XmlImport" onAction="ImportXML.ImportVertrag" />
          <button id="ExportVertragsvarianten" label="Erzeuge Test-XMLs" imageMso="XmlExport" onAction="ExportTestXMLs.CreateTestXMLs" />
          <button id="StartButton" label="Aufruf Stefan-Tools" imageMso="MacroPlay" onAction="StarteBefuellungen.StarteDialog" />
          <button id="MakeCSV" label="CSV erzeugen" imageMso="ExportTextFile" onAction="SheetGeVoBfr.CreateCsv" />
          <button id="NeuerVertrag" label="Speichern &amp;&amp; Neu" imageMso="SlideNew" onAction="ImportXML.NeuerVertrag" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

Hinweise: 
- Die aufgerufenen Methoden müssen ohne "()" angegeben werden. Falls Klammern angegeben werden, nimmt VBA an, dass es sich um eine Funktion handelt und erwartet einen Rückgabewert.
- Wenn - wie im Beispiel - rein der Methodenname im XML angegeben wird, so ruft Excel (offenbar) diese Methode implizit mit dem Sender als Parameter auf. Eine Methode *ImportVertragsCSV* sollte also folgerndermaßen deklariert werden, um Probleme zu vermeiden: 
  ```VBA
  Sub ImportVertragsCSV(control As IRibbonControl)
      ' Do something
  End Sub
  ```
- Bei Bedarf für andere Icons gibt es hier eine Seite mit Icons und deren internen Namen: https://bert-toolkit.com/imagemso-list.html