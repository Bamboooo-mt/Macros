Option Explicit
Attribute VB_Name = "addFootnotesToHackTools"
Sub addFootnotesToHackTools()

    Dim doc As Document
    Dim selRange As Range
    Set doc = ActiveDocument
    Set selRange = Selection.Range
    
    ' Объявляем массив инструментов (введите число инструментов от 0)
    ' We declare an array of tools (enter the number of tools from 0)
    Dim tools() As Variant
    ReDim tools(0 To 109)
    
    ' Строка записывается в формате: название инструмента, текст для сноски, include-условие, exclude-условие
    ' The line is recorded in the format: the name of the instrument, text for the footnote, the Include-melting, Exclude
    tools(0) = Array("adidnsdump", "https://github.com/dirkjanm/adidnsdump", "", "")
    tools(1) = Array("ADExplorerSnapshot", "https://github.com/c3c/ADExplorerSnapshot.py", "", "")
    tools(2) = Array("Apktool", "https://github.com/iBotPeaches/Apktool", "", "")
    tools(3) = Array("atexec", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/atexec.py", "", "")
    tools(4) = Array("BadPotato", "https://github.com/BeichenDream/BadPotato", "", "")
    tools(5) = Array("Bloodhound", "https://github.com/BloodHoundAD/BloodHound", "", "")
    tools(6) = Array("Bloodhound.py", "https://github.com/fox-it/BloodHound.py", "", "")
    tools(7) = Array("Certipy", "https://github.com/ly4k/Certipy", "", "")
    tools(8) = Array("CrackMapExec", "https://github.com/byt3bl33d3r/CrackMapExec", "", "")
    tools(9) = Array("cycript", "http://www.cycript.org/", "", "")
    tools(10) = Array("decrypt_chrome_password", "https://github.com/ohyicong/decrypt-chrome-passwords/blob/main/decrypt_chrome_password.py", "", "")
    tools(11) = Array("dirsearch", "https://github.com/maurosoria/dirsearch", "", "")
    tools(12) = Array("DNSlivery", "https://github.com/no0be/DNSlivery", "", "")
    tools(13) = Array("DSInternals", "https://github.com/MichaelGrafnetter/DSInternals", "", "")
    tools(14) = Array("dpapi", "https://github.com/fortra/impacket/blob/master/examples/dpapi.py?ysclid=lr7qm27cvf426269446", "", "")
    tools(15) = Array("EfsPotato", "https://github.com/zcgonvh/EfsPotato", "", "")
    tools(16) = Array("exchanger", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/exchanger.py", "", "")
    tools(17) = Array("firefox_decrypt", "https://github.com/unode/firefox_decrypt", "", "")
    tools(18) = Array("ffuf", "https://github.com/ffuf/ffuf", "", "")
    tools(19) = Array("Frida", "https://www.frida.re/", "", "")
    tools(20) = Array("Get-LoggedOn", "https://gist.github.com/GeisericII/6849bc86620c7a764d88502df5187bd0", "", "")
    tools(21) = Array("getST", "https://github.com/fortra/impacket/blob/master/examples/getST.py", "", "")
    tools(22) = Array("getTGT", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/getTGT.py", "", "")
    tools(23) = Array("GetUserSPNs", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/GetUserSPNs.py", "", "")
    tools(24) = Array("HAProxy", "https://github.com/haproxy/haproxy", "", "")
    tools(25) = Array("hashcat", "https://github.com/hashcat/hashcat", "", "")
    tools(26) = Array("http-ntlm-info", "https://nmap.org/nsedoc/scripts/http-ntlm-info.html", "", "")
    tools(27) = Array("httpx", "https://github.com/projectdiscovery/httpx", "", "")
    tools(28) = Array("hostapd-wpe", "https://github.com/aircrack-ng/aircrack-ng/tree/master/patches/wpe/hostapd-wpe", "", "")
    tools(29) = Array("Hydra", "https://github.com/vanhauser-thc/thc-hydra", "", "")
    tools(30) = Array("IIS Short Name Scanner", "https://github.com/irsdl/IIS-ShortName-Scanner", "", "")
    tools(31) = Array("iis_tilde_enum", "https://github.com/esabear/iis_tilde_enum", "", "")
    tools(32) = Array("incognito", "https://github.com/FSecureLABS/incognito", "", "")
    tools(33) = Array("Invoke Kerberoast", "https://github.com/EmpireProject/Empire/blob/master/data/module_source/credentials/Invoke-Kerberoast.ps1", "", "")
    tools(34) = Array("jmet", "https://github.com/matthiaskaiser/jmet", "", "")
    tools(35) = Array("ldap_shell", "https://github.com/PShlyundin/ldap_shell", "", "")
    tools(36) = Array("LDAPDomainDump", "https://github.com/dirkjanm/ldapdomaindump", "", "")
    tools(37) = Array("LDAPPER", "https://github.com/shellster/LDAPPER", "", "")
    tools(38) = Array("loubia", "https://github.com/metalnas/loubia", "", "")
    tools(39) = Array("LsassSilentProcessExit", "https://github.com/deepinstinct/LsassSilentProcessExit", "", "")
    tools(40) = Array("LLDB", "http://lldb.llvm.org/", "", "")
    tools(41) = Array("lyncsmash", "https://github.com/nyxgeek/lyncsmash", "", "")
    tools(42) = Array("Metasploit framework", "https://github.com/rapid7/metasploit-framework", "", "")
    tools(43) = Array("mimikatz", "https://github.com/gentilkiwi/mimikatz", "", "")
    tools(44) = Array("mitm6", "https://github.com/fox-it/mitm6", "", "")
    tools(45) = Array("mRemoteNG Decrypt", "https://github.com/kmahyyg/mremoteng-decrypt", "", "")
    tools(46) = Array("nanodump", "https://github.com/fortra/nanodump", "", "")
    tools(47) = Array("Neo-reGeorg", "https://github.com/L-codes/Neo-reGeorg", "", "")
    tools(48) = Array("Nmap", "https://github.com/nmap/nmap", "", "")
    tools(49) = Array("smb-os-discovery", "https://nmap.org/nsedoc/scripts/smb-os-discovery.html", "", "")
    tools(50) = Array("noPac", "https://github.com/Ridter/noPac", "", "")
    tools(51) = Array("psexec", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/psexec.py", "IMPACKET", "")
    tools(52) = Array("psexec", "https://docs.microsoft.com/en-us/sysinternals/downloads/psexec", "PSTOOLS", "")
    tools(53) = Array("Nuclei", "https://github.com/projectdiscovery/nuclei", "", "")
    tools(54) = Array("oathtool", "https://manpages.ubuntu.com/manpages/trusty/man1/oathtool.1.html", "", "")
    tools(55) = Array("Objection", "https://github.com/sensepost/objection", "", "")
    tools(56) = Array("Oracle Database Attacking Tool", "https://github.com/quentinhardy/odat", "", "")
    tools(57) = Array("patator", "https://github.com/lanjelot/patator", "", "")
    tools(58) = Array("PEAS", "https://github.com/WithSecureLabs/peas", "", "")
    tools(59) = Array("PHPGGC", "https://github.com/ambionics/phpggc", "", "")
    tools(60) = Array("PowerView", "https://github.com/aniqfakhrul/powerview.py", "", "")
    tools(61) = Array("ProcessHacker", "https://processhacker.sourceforge.io/downloads.php", "", "")
    tools(62) = Array("Proxifier", "https://proxifier.com/", "", "")
    tools(63) = Array("proxychains", "https://github.com/haad/proxychains", "", "")
    tools(64) = Array("PyInstaller", "https://github.com/pyinstaller/pyinstaller", "", "")
    tools(65) = Array("pypykatz", "https://github.com/skelsec/pypykatz", "", "")
    tools(66) = Array("rbcd", "https://github.com/fortra/impacket/blob/master/examples/rbcd.py", "", "")
    tools(67) = Array("reGeorg", "https://github.com/sensepost/reGeorg", "", "")
    tools(68) = Array("RegSave", "https://github.com/EncodeGroup/RegSave", "", "")
    tools(69) = Array("restorepassword", "https://github.com/dirkjanm/CVE-2020-1472/blob/master/restorepassword.py", "", "")
    tools(70) = Array("rpc2socks", "https://github.com/lexfo/rpc2socks", "", "")
    tools(71) = Array("rpcclient", "https://www.samba.org/samba/docs/current/man-html/rpcclient.1.html", "", "")
    tools(72) = Array("rpivot", "https://github.com/klsecservices/rpivot", "", "")
    tools(73) = Array("Rubeus", "https://github.com/GhostPack/Rubeus", "", "")
    tools(74) = Array("s5.go", "https://github.com/ring04h/s5.go", "", "")
    tools(75) = Array("secretsdump", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/secretsdump.py", "", "")
    tools(76) = Array("SharpSCCM", "https://github.com/Mayyhem/SharpSCCM", "", "")
    tools(77) = Array("SharpChrome", "https://github.com/GhostPack/SharpDPAPI/tree/master/SharpChrome", "", "")
    tools(78) = Array("SharpHound", "https://github.com/BloodHoundAD/SharpHound", "", "")
    tools(79) = Array("SharpSecretsdump", "https://github.com/laxa/SharpSecretsdump", "", "")
    tools(80) = Array("shortscan", "https://github.com/bitquark/shortscan", "", "")
    tools(81) = Array("SIET", "https://github.com/Sab0tag3d/SIET", "", "")
    tools(82) = Array("SigmaPotato", "https://github.com/tylerdotrar/SigmaPotato", "", "")
    tools(83) = Array("smbclient", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/smbclient.py", "", "")
    tools(84) = Array("smbserver", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/smbserver.py", "", "")
    tools(85) = Array("SMBExec", "https://github.com/fortra/impacket/blob/master/examples/smbexec.py", "", "")
    tools(86) = Array("smbspray", "https://github.com/absolomb/smbspray", "", "")
    tools(87) = Array("SSF", "https://securesocketfunneling.github.io/ssf/", "", "")
    tools(88) = Array("SSH Password logging via PAM", "https://github.com/cameron-gagnon/ssh_pass_logging", "", "")
    tools(89) = Array("ticketer", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/ticketer.py", "", "")
    tools(90) = Array("tsh", "https://github.com/creaktive/tsh", "", "")
    tools(91) = Array("tstool", "https://github.com/fortra/impacket/blob/master/examples/tstool.py", "", "")
    tools(92) = Array("Watson", "https://github.com/rasta-mouse/Watson", "", "")
    tools(93) = Array("WinSCP", "https://winscp.net/eng/index.php", "", "")
    tools(94) = Array("winscppasswd", "https://github.com/anoopengineer/winscppasswd", "", "")
    tools(95) = Array("wmiexec", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/wmiexec.py", "", "")
    tools(96) = Array("wpa_supplicant", "https://github.com/jmalinen/hostap/tree/master/wpa_supplicant", "", "")
    tools(97) = Array("ysoserial.net", "https://github.com/pwntester/ysoserial.net", "", "")
    tools(98) = Array("zerologon_exploit", "https://github.com/dirkjanm/CVE-2020-1472?ysclid=lslt5i9589405163212", "", "")
    tools(99) = Array("zerologon_tester", "https://github.com/SecuraBV/CVE-2020-1472/blob/master/zerologon_tester.py", "", "")
    tools(100) = Array("Эмулятор MySQL-сервера", "https://github.com/allyshka/Rogue-MySql-Server", "", "")
    tools(101) = Array("Coercer", "https://github.com/p0dalirius/Coercer", "", "")
    tools(102) = Array("Gost", "https://github.com/ginuerzh/gost", "", "")
    tools(103) = Array("rsocks", "https://github.com/brimstone/rsocks", "", "")
    tools(104) = Array("pre2k", "https://github.com/garrettfoster13/pre2k", "", "")
    tools(105) = Array("ntlmrelayx", "https://github.com/SecureAuthCorp/impacket/blob/master/examples/ntlmrelayx.py", "", "")
    tools(106) = Array("KeeThief", "https://github.com/GhostPack/KeeThief", "", "")
    tools(107) = Array("adexplorer", "https://learn.microsoft.com/en-us/sysinternals/downloads/adexplorer", "", "")
    tools(108) = Array("sucrack", "https://github.com/hemp3l/sucrack", "", "")
    tools(109) = Array("changepasswd", "https://github.com/fortra/impacket/blob/master/examples/changepasswd.py", "", "")
    
    ' Словарь для отслеживания уже добавленных инструментов (по индексу)
    ' Dictionary for tracking already added tools (by index)
    Dim addedToolsDict As Object
    Set addedToolsDict = CreateObject("Scripting.Dictionary")
    
    Dim sentence As Range
    Dim sentenceText As String, upperSentenceText As String
    Dim toolEntry As Variant
    Dim keyword As String, link As String, includeCond As String, excludeCond As String
    Dim posInSentence As Long, posInsert As Long, absPos As Long
    Dim foundRange As Range
    Dim conditionsOK As Boolean
    Dim i As Long
    Dim part As Variant
    Dim incParts As Variant, excParts As Variant
    Dim foundPos As Long, candidate As String
    
    ' Обходим каждое предложение в выделенной области
    ' We go around each proposal in the allocated area
    For Each sentence In selRange.Sentences
        sentenceText = sentence.text
        upperSentenceText = UCase(sentenceText)
        
        ' Обходим все инструменты
        ' We go around all the tools
        For i = LBound(tools) To UBound(tools)
            ' Если инструмент уже добавлен, пропускаем его
            ' If the tool is already added, we miss it
            If Not addedToolsDict.Exists(CStr(i)) Then
                toolEntry = tools(i)
                keyword = UCase(toolEntry(0))
                link = toolEntry(1)
                includeCond = UCase(toolEntry(2))
                excludeCond = UCase(toolEntry(3))
                
                ' Другой вариант объявления исключений
                ' Another option for announcing exceptions
                If keyword = "REGEORG" And InStr(upperSentenceText, "NEO-REGEORG") > 0 Then GoTo NextTool
                If keyword = "ADEXPLORER" And InStr(upperSentenceText, "SNAPSHOT") > 0 Then GoTo NextTool
                
                posInSentence = InStr(1, upperSentenceText, keyword, vbTextCompare)
                If posInSentence > 0 Then
                    conditionsOK = True
                    ' Проверка include-условий
                    ' Verification of Include consequences
                    If includeCond <> "" Then
                        incParts = Split(includeCond, ",")
                        For Each part In incParts
                            If InStr(upperSentenceText, Trim(part)) = 0 Then
                                conditionsOK = False
                                Exit For
                            End If
                        Next part
                    End If
                    ' Проверка exclude-условий
                    ' Checking Exclude
                    If conditionsOK And excludeCond <> "" Then
                        excParts = Split(excludeCond, ",")
                        For Each part In excParts
                            If InStr(upperSentenceText, Trim(part)) > 0 Then
                                conditionsOK = False
                                Exit For
                            End If
                        Next part
                    End If
                    
                    If conditionsOK Then
                        ' Используем метод Find с настройкой для поиска целых слов
                        ' We use the Find method with setting to search for whole words
                        Set foundRange = sentence.Duplicate
                        With foundRange.Find
                            .text = toolEntry(0)
                            .MatchCase = False
                            .MatchWholeWord = True
                            .Execute
                        End With
                        If Not foundRange Is Nothing Then
                            Dim isSup As Boolean
                            isSup = False
                            Dim ch As Range
                            For Each ch In foundRange.Characters
                                If ch.Font.Superscript Then
                                    isSup = True
                                    Exit For
                                End If
                            Next ch
                            If Not isSup Then
                                ' Используем позицию найденного совпадения
                                ' We use the position of the coincidence found
                                foundPos = foundRange.Start - sentence.Start + 1
                                candidate = Mid(sentenceText, foundPos, Len(keyword) + 3)
                                ' Спец. обработка для инструментов с .py
                                ' Specialist.Tools for tools with .py
                                If StrComp(candidate, keyword & ".py", vbTextCompare) = 0 Then
                                    posInsert = foundPos + Len(keyword) + 3 - 1
                                Else
                                    posInsert = foundPos + Len(keyword) - 1
                                End If
                                absPos = sentence.Start + posInsert
                                doc.Footnotes.Add Range:=doc.Range(absPos, absPos), text:=link
                                addedToolsDict.Add CStr(i), True
                            End If
                        End If
                    End If
                End If
            End If
NextTool:
        Next i
    Next sentence
End Sub


