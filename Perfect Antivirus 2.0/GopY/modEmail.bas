Attribute VB_Name = "modEmail"
Public Function SendMail(xEmail, xNoiDung)
Dim Flds
Dim iMsg As New CDO.Message
Dim iConf As New CDO.Configuration
Set Flds = iConf.Fields

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "PerfectAV2009"
Flds.Item(schema & "sendpassword") = "htgtalcmdltnsc"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update
 
With iMsg
.To = xEmail
.CC = ""
.BCC = ""
.From = "<perfectav2009@gmail.com>"
.Subject = "Danh gia ve chuong trinh Perfect Antivirus 2009"
.HTMLBody = xNoiDung
.TextBody = xNoiDung
.Sender = "perfectav2009@gmail.com"
.Organization = "<PerfectAV2009>"
.ReplyTo = "perfectav2009@gmail.com"
Set .Configuration = iConf
.Send
End With
 
Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing
End Function
