Attribute VB_Name = "SignaturePath"
Sub UpdateSignaturePath()
    ' Camron Walker 2019-12-10
    ' adds email signature path to named range Email_Signature_Path
    
    Dim startEmailSignPath As String
    Dim strFileToOpen As String
    
    startEmailSignPath = Environ("appdata") & _
                "\Microsoft\Signatures\"
            
    ChDir startEmailSignPath
    ChDrive startEmailSignPath
    
    strFileToOpen = Application.GetOpenFilename _
        (Title:="Please select the email signature you want to use.", _
        FileFilter:="Document Files *.htm* (*.htm*),")
        
    If Not srtFileToOpen = False Then Range("Email_Signature_Path").Value = strFileToOpen
    
End Sub


