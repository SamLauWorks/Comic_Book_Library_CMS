Attribute VB_Name = "Member"
Sub membersload()
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path + "\BookD.mdb")
Set Members = db.OpenRecordset("·|­û")
End Sub

Function Ending()
End
End Function
