VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form2"
   ScaleHeight     =   0
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim mData As Recordset
Dim nData As Recordset

Private Sub Form_Load()
Set WS = DBEngine.Workspaces(0)
DbFile = (AppPath & "DATA.PAV")
PwdString = "htgtalcmdltnsc"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)

Set mData = db.OpenRecordset("STRI", dbOpenTable)

mData.AddNew
mData.Fields("String") = "<decription>AutoIT"
mData.Fields("VirusName") = "AutoIT File"
mData.Update


mData.AddNew
mData.Fields("String") = "bpk.exe"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpkun.exe"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpkvw.exe"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpkhk.dll"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpki.dll"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpkwb.dll"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpkr.exe"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "pk.bin"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "bpkch.dat"
mData.Fields("VirusName") = "Trojan.Perfect-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "¡èÈýºÃÑ§Éú¡è"
mData.Fields("VirusName") = "Virus.W32.Explorer.exe"
mData.Update


mData.AddNew
mData.Fields("String") = "C:\windows\system32\drivers\qciuw.exe"
mData.Fields("VirusName") = "Virus.W32.Explorer.exe"
mData.Update


mData.AddNew
mData.Fields("String") = "ü>HüFlüûï\ü`ýÇ|þ>xþýÇÐþõT"
mData.Fields("VirusName") = "Trojan.W32.Spy007"
mData.Update


mData.AddNew
mData.Fields("String") = "hÿôÿpfÿlpÿäDTÿ%lpÿ"
mData.Fields("VirusName") = "Trojan.W32.Spy007.b"
mData.Update


mData.AddNew
mData.Fields("String") = "jýëªÿuë¥ÿu‰W0VWè"
mData.Fields("VirusName") = "Trojan.Ardamax.Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "šiûì³uÆv÷š‘û²#M·îvÚñdU^5rÿët‡ð/"
mData.Fields("VirusName") = "Trojan.Ardamax.Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "http://AbsoluteKeyLogger.com"
mData.Fields("VirusName") = "AbsoluteKeyLogger"
mData.Update


mData.AddNew
mData.Fields("String") = "ä÷ÿÿèbÉÿÿ‹…ä÷ÿÿè#÷ÿPh"
mData.Fields("VirusName") = "Virus.W32.Active.Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "KMiNT21.SpyArsenal.FamilyKeyLogger"
mData.Fields("VirusName") = "FamilyKeyLogger"
mData.Update


mData.AddNew
mData.Fields("String") = "KMiNT21\GoldenKeylogger"
mData.Fields("VirusName") = "GoldenKeylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "http://www.kerneltek.com/downloads/kl21download.htm"
mData.Fields("VirusName") = "Virus.Keylover.Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "CompanyName.ProductName.MiniKeyLog"
mData.Fields("VirusName") = "MiniKeyLog"
mData.Update


mData.AddNew
mData.Fields("String") = "http://www.amplusnet.com/products/stealthkeylogger/overview.htm"
mData.Fields("VirusName") = "StealthKeylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "A+ Stealth KeyLogger"
mData.Fields("VirusName") = "A+ Stealth KeyLogger"
mData.Update


mData.AddNew
mData.Fields("String") = "ÂtŠHÿ:Êt€ùÿt;ú‹ÎuLþHÿRèPúÿÿ2ÀëG;úu8T"
mData.Fields("VirusName") = "Trojan.SCKeylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "‹MðƒÁ@ÿ%P‹MðƒÁPÿ%P‹M"
mData.Fields("VirusName") = "Trojan.SCKeylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "S‹øÿÖ…ÿ‹ðts…ötoÇD$"
mData.Fields("VirusName") = "Virus.Tiny.Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "http://www.xp-tools.com/xpadvancedkeylogger/index.htm"
mData.Fields("VirusName") = "Virus.XP_Advandce_Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "http://www.toolsanywhere.com/advanced-invisible-keylogger.html"
mData.Fields("VirusName") = "Virus.Advandce_Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "ßÛÔÒÓ†BâÔñø³ÝhnƒÚÍÿÿÿ¾"
mData.Fields("VirusName") = "Keylogger.Trojan.OinFP.exe"
mData.Update

mData.AddNew
mData.Fields("String") = "http://campaigns.outerinfo.net/client_settings_3.bin"
mData.Fields("VirusName") = "Trojan.Keylogger.OuterInfo.exe"
mData.Update


mData.AddNew
mData.Fields("String") = "http://www.ardamax.com/order_akl.html"
mData.Fields("VirusName") = "Trojan.Aradamax-Keylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "ÞÒ‡@LnÆ4Q§4^ÓÒØB×#Ü4W4òæÒø×{"
mData.Fields("VirusName") = "Virus.stealthkeylogger"
mData.Update


mData.AddNew
mData.Fields("String") = "ïnñìªëÅ©6MøYr:å½ÏÉxt–;¬º„Cži{'ü'dd˜TþÁˆ$"
mData.Fields("VirusName") = "Worm.Autorun.Picture"
mData.Update


mData.AddNew
mData.Fields("String") = "prÿkrÿô@ßkrÿôZÕÄkrÿôaßkrÿôzÕÄÅkrÿô0ßkrÿô9ÕÄÅkrÿô_ÆÅkrÿô.ÆÅY"
mData.Fields("VirusName") = "Worm.Autorun.Brontok"
mData.Update


End Sub
