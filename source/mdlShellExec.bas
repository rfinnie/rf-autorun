Attribute VB_Name = "mdlShellExec"
      Option Explicit

      Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

      Private Declare Function GetDesktopWindow Lib "user32" () As Long

      Const SW_SHOWNORMAL = 1

      Const SE_ERR_FNF = 2&
      Const SE_ERR_PNF = 3&
      Const SE_ERR_ACCESSDENIED = 5&
      Const SE_ERR_OOM = 8&
      Const SE_ERR_DLLNOTFOUND = 32&
      Const SE_ERR_SHARE = 26&
      Const SE_ERR_ASSOCINCOMPLETE = 27&
      Const SE_ERR_DDETIMEOUT = 28&
      Const SE_ERR_DDEFAIL = 29&
      Const SE_ERR_DDEBUSY = 30&
      Const SE_ERR_NOASSOC = 31&
      Const ERROR_BAD_FORMAT = 11&


      Public Function MyShellExec(cmd)
          Dim r As Long, msg As String
          Dim Scr_hDC As Long
          
          Scr_hDC = GetDesktopWindow()
          r = ShellExecute(Scr_hDC, "Open", cmd, "", App.Path, SW_SHOWNORMAL)
          If r <= 32 Then
              'There was an error
              Select Case r
                  Case SE_ERR_FNF
                      msg = "File not found"
                  Case SE_ERR_PNF
                      msg = "Path not found"
                  Case SE_ERR_ACCESSDENIED
                      msg = "Access denied"
                  Case SE_ERR_OOM
                      msg = "Out of memory"
                  Case SE_ERR_DLLNOTFOUND
                      msg = "DLL not found"
                  Case SE_ERR_SHARE
                      msg = "A sharing violation occurred"
                  Case SE_ERR_ASSOCINCOMPLETE
                      msg = "Incomplete or invalid file association"
                  Case SE_ERR_DDETIMEOUT
                      msg = "DDE Time out"
                  Case SE_ERR_DDEFAIL
                      msg = "DDE transaction failed"
                  Case SE_ERR_DDEBUSY
                      msg = "DDE busy"
                  Case SE_ERR_NOASSOC
                      msg = "No association for file extension"
                  Case ERROR_BAD_FORMAT
                      msg = "Invalid EXE file or error in EXE image"
                  Case Else
                      msg = "Unknown error"
              End Select
              MsgBox msg
          End If
      End Function

