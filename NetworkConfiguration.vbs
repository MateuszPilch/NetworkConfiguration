If Not WScript.Arguments.Named.Exists("elevate") Then 'uruchamia nowe okno z uprawnieniami do modyfikacji interfejsu
CreateObject("Shell.Application").ShellExecute WScript.FullName _
, """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
WScript.Quit
End If

strComputer= "."
listID = 0        'id obecnego interfejsu sieciowego
x = 0             'id wybranego interfejsu sieciowego
adapterName = ""  'nazwa wybranego interfejsu sieciowego

propertyList = Array("Adres IPv4: ","Maska podsieci: ", "Brama domyslna: ", "Adres DNS: ")
orderList = Array("Podaj adres IPv4:","Podaj adres maski podsieci:", "Podaj adres bramy:", "Podaj adres DNS:")

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set adapterList = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE") 'zmienna przechowujaca informacje o konfiguracji interfejsow

Display

Function Display

  listID = 0
  displayText = ""

  For Each objItem In adapterList

    displayText = displayText & "[" & listID & "]" & Right(objItem.Caption,Len(objItem.Caption) - 10) & vbCrLf 'lista dostepnych interfejsow
    listID = listID + 1

  Next

  x = InputBox("Lista interfejsow sieciowych:" & vbCrLf & displayText,"Wprowadz wybrane ID interfejsu") 'okno wyboru interfejsu

  if IsNumeric(x) = False then
    msg = MsgBox ("Wpisana wartosc jest nieprawidlowa!",vbCritical,"Wystapil problem") 'komunikat o blednie zostanie wpisana litera
    Display

  end if 

  x = CInt(x)
  listID = 0
  displayText = ""
  adapterName = ""

  For Each objItem In adapterList 'okno z informacjami o wybranym interfejsie

      if(x = listID) then
      
        displayText = "Konfiguracja interfejsu: " & vbCrLf & _
        propertyList(0) & ListValueCheck(objItem.IPAddress(0)) & vbCrLf & _
        propertyList(1) & ListValueCheck(objItem.IPSubnet(0)) & vbCrLf & _
        propertyList(2) & ListValueCheck(objItem.DefaultIPGateway(0)) & vbCrLf & _
        propertyList(3) & ListValueCheck(objItem.DNSServerSearchOrder(0))

        adapterName = Right(objItem.Caption,Len(objItem.Caption) - 10)
      
      end if
      
      listID = listID + 1

  Next

  if x < 0 or x >= listID then

    msg = MsgBox("Wpisana wartosc jest nieprawidlowa!", vbCritical,"Wystapil problem") 'okno bledu w przypadku wpisania zlego id
    Display

  end if 

  msg = MsgBox(displayText,vbInformation, adapterName)
  msg = MsgBox ("Czy chcesz zmodyfikowac wybrany interfejs ?",vbYesNo+vbInformation, adapterName) 'okno wyboru dalszego dzialania lub wyjscia z programu

  if msg = vbYes Then

    Modify

  else 

    WScript.Quit()

  end if

End Function

Function Modify 'funkcja konfiguruje wybrany interfejs sieciowy

  listID = 0

  promptList = Array("","","","") ' ipv4, submask, gateway, dns
  
  For i = 0 To 3 Step 1

    msg = InputBox(propertyList(0) & promptList(0) & vbCrLf & _
      propertyList(1) & promptList(1) & vbCrLf & _
      propertyList(2) & promptList(2) & vbCrLf & _
      propertyList(3) & promptList(3) & vbCrLf  & vbCrLf & _
      orderList(i), adapterName)

    promptList(i) = msg

  Next

  errorCount = 0
  For Each objItem In adapterList

    if(x = listID) then
 
      a = objItem.EnableStatic(Array(promptList(0)), Array(promptList(1)))
      b = objItem.SetGateways(Array(promptList(2)))
      c = objItem.SetDNSServerSearchOrder(Array(promptList(3)))

      errorCount = a + b + c 'wartosc przechowuje sume kodow bledow ktore wystapily 

    end if
    listID = listID + 1

  Next

  if errorCount <> 0 then 'sprawdzenie czy wystapil blad

    msg = MsgBox ("Wystapil problem podczas konfiguracji interfejsu. Czy chcesz sprobowac ponownie ?", vbYesNo + vbCritical, adapterName) 'komunikat o wystapieniu bledu

      if msg = vbNo then

      WScript.Quit()

      else

        Modify

      end if    

  else

    msg = MsgBox ("Konfiguracja zakonczona pomyslnie! Czy kontynuowac dzialanie ?", vbYesNo + vbInformation ,"Sukces!") 'komunikat o pomyslym zakonczeniu konfiguracji

    if msg = vbNo then

      WScript.Quit()

      else

        Display

      end if

  end if

End Function

Function ListValueCheck(val) 'funkcja sprawdza poprawnosc wyswietlanych danych

    if IsNull(val) or IsEmpty(val) or VarType(val) = 10 then

    ListValueCheck = "Brak danych"

    else
    
    ListValueCheck = val
    
    end if

End Function