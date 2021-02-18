Attribute VB_Name = "Misc"
Sub RestoreSettingsToDefaults()
    Range("Dev_Mode").Value = False
    Range("Logging").Value = True
    Range("Custom_File_Location").Value = False
    Range("SENDorDISPLAYemail").Value = "DISPLAY"
    Range("Email_Table_Filter").Value = "<>Closeout"
    Range("Email_Hide_Closed").Value = "SHOW"

End Sub
