Option Explicit
Dim RunTimer As Date
Sub Scheduler()
'
RunTimer = TimeValue("00:00:00")

Application.OnTime RunTimer, "Scheduler"
'MsgBox "Take a break.", vbInformation
Call Test

End Sub
'outlook16
Sub Test()

Dim olApp As Outlook.Application
Dim olNs As Namespace
Dim Fldr As MAPIFolder
Dim olMail As Variant
Dim olAtt As Outlook.Attachment
Dim mailitem As Outlook.mailitem
Dim strFname As String
Dim strTempFolder As String
strTempFolder = Environ("Temp") & Chr(92)

Set olApp = New Outlook.Application
Set olNs = olApp.GetNamespace("MAPI")
Set Fldr = olNs.GetDefaultFolder(olFolderInbox)

Dim rectime As Date
Dim theDate As Date
theDate = Now()
Dim i As Variant
Dim ti As String
Dim va As Variant
Dim va1 As Variant

Dim sumcai_h As Integer
Dim sumcai_i As Long
Dim sumaft_h As Integer
Dim sumaft_i As Long
Dim sumpsb_h As Integer
Dim sumpsb_i As Long
Dim sumreu_h As Integer
Dim sumreu_i As Long

Dim zipFileName As String
Dim unzipFolderName As String
Dim objZipItems As FolderItems
Dim objZipItem As FolderItem

Dim wShApp As Shell
Set wShApp = CreateObject("Shell.Application")


For Each olMail In Fldr.Items
    If InStr(olMail.Subject, "_0800_Passerelle_CORP_PROD_Fichiers_POCs_FR_et_RI de la veille") <> 0 Then
        ti = olMail.ReceivedTime

        If Left(ti, 10) = Left(Now, 10) Then
            olMail.Display

            'Workbooks.OpenText Filename:="C:\Users\xxx\Desktop\Monitoring_svk\Supervision Downstream SK.xlsx"
            'ActiveWorkbook.Sheets("ARC_ETL").Activate
            'Range("A1:I1").Select
            'Range(Selection, Selection.End(xlDown)).Select
            'Selection.ClearContents
            'Range("A1").Select
            
            'ActiveWorkbook.Sheets("ARC_PAYS").Activate
            'Range("A1:I1").Select
            'Range(Selection, Selection.End(xlDown)).Select
            'Selection.ClearContents
            'Range("A1").Select
            
            For Each olAtt In olMail.Attachments
                If Right(olAtt.Filename, 3) = "zip" Then
                    MsgBox "Le mail du jour est: " & olAtt.Filename
                    olAtt.SaveAsFile "C:\Users\xxx\Desktop\Monitoring_svk" & "\" & olAtt.Filename '
                    
                    zipFileName = "C:\Users\xxx\Desktop\Monitoring_svk" & "\" & olAtt.Filename '
                    unzipFolderName = "C:\Users\xxx\Desktop\Monitoring_svk" '
                    Set objZipItems = wShApp.Namespace(zipFileName).Items
                    
                    For Each objZipItem In wShApp.Namespace(zipFileName).Items
                        If InStr(objZipItem.Name, "SC2") <> 0 Then
                            wShApp.Namespace(unzipFolderName).CopyHere objZipItem
                        
                            Workbooks.OpenText Filename:= _
                            "C:\Users\xxx\Desktop\Monitoring_svk\" & objZipItem '
                            
                            Columns("A:A").Select
                            'Application.CutCopyMode = False
                            
                            Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                            :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
                            Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1)), TrailingMinusNumbers:=True

                            'va = Left(objZipItem, 16)
                            'ActiveWorkbook.Sheets(va).Activate
                            'ActiveWorkbook.Close SaveChanges:=False
                            
                            Columns("H:I").Select
                            Selection.NumberFormat = "General"
                            
                            Range("H1") = Application.WorksheetFunction.Sum(Range("H:H"))
                            Range("I1") = Application.WorksheetFunction.Sum(Range("I:I"))
                            
                            If InStr(objZipItem.Name, "CAI") <> 0 Then
                                'MsgBox "CAI"
                                sumcai_h = Range("H1").Value
                                'sumcai_i = Range("I1").Value
                            ElseIf InStr(objZipItem.Name, "AFT") <> 0 Then
                                'MsgBox "AFT"
                                sumaft_h = Range("H1").Value
                                sumaft_i = Range("I1").Value
                            ElseIf InStr(objZipItem.Name, "PSB") <> 0 Then
                                'MsgBox "PSB"
                                sumpsb_h = Range("H1").Value
                                sumpsb_i = Range("I1").Value
                            Else
                                'MsgBox "REU"
                                sumreu_h = Range("H1").Value
                                sumreu_i = Range("I1").Value
                                
                            End If
                            

                        End If
                    Next
                End If
            Next olAtt
            
            'MsgBox sumcai_h
            'MsgBox sumaft_h
            'MsgBox sumpsb_h
            'MsgBox sumreu_h

        End If
    End If

Next olMail

End Sub

