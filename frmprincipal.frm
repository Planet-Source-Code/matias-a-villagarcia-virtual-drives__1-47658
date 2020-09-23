VERSION 5.00
Begin VB.Form frmprincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Virtual Drive Creator"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmprincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkiniciarconwindows 
      Caption         =   "Create Virtual Drive at Windows startup"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton cmdborrardiscovirtual 
      Caption         =   "&Delete Virt. Drive"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdinformacion 
      Caption         =   "&Information..."
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdcrear 
      Caption         =   "&Create"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdseleccionar 
      Caption         =   "&Select..."
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtdiscovirtual 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblcarpeta 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "None"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Virtual Drive Folder"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmprincipal.frx":0442
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6810
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RutaCarpeta As String 'RutaCarpeta=FolderPath
Dim DiscoVirtual As String 'DiscoVirtual=VirtualDrive
'---
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Function BuscarCarpeta() As String
'KPD-Team 1998
'URL: http://www.allapi.net/
'KPDTeam@Allapi.net
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo
With udtBI
    'Set the owner window
    .hWndOwner = Me.hWnd
    'lstrcat appends the two strings and returns the memory address
    .lpszTitle = lstrcat(LetraHD(App.Path), "")
    'Return only if the user selected a directory
    .ulFlags = BIF_RETURNONLYFSDIRS
End With
'Show the 'Browse for folder' dialog
lpIDList = SHBrowseForFolder(udtBI)
If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    'Get the path from the IDList
    SHGetPathFromIDList lpIDList, sPath
    'free the block of memory
    CoTaskMemFree lpIDList
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
        sPath = Left$(sPath, iNull - 1)
    End If
End If
BuscarCarpeta = sPath
End Function

Private Sub cmdborrardiscovirtual_Click()
'This will delete a virtual drive
If MsgBox("This will delete the current Virtual Drive, do you want to procced?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete Virtual Drive") = vbYes Then
    'check the app path, if the len of app.path
    'is 3 then it must be X:\ where X is your
    'hard-disk. I know it's no best way to do
    'this but works!!!
    If Len(App.Path) = 3 Then
        'txtdiscovirtual=txtvirtualdisk
        ret = Shell("subst /d " & txtdiscovirtual.Text, vbHide)
        'chkiniciarconwindows=chkrunatwindowsstartup
        If chkiniciarconwindows.Value = 1 Then
            Kill App.Path & "CreateVirtualDrive.bat"
            DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "VirtualDrive"
        End If
    Else
        'see above
        ret = Shell("subst /d " & txtdiscovirtual.Text, vbHide)
        If chkiniciarconwindows.Value = 1 Then
            Kill App.Path & "\CreateVirtualDrive.bat"
            DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "VirtualDrive"
        End If
    End If
End If
End Sub

Private Sub cmdcrear_Click()
'this will create the virtual drive
Dim ArchivoNumero As Integer 'ArchivoNumero=FileNumber
'get a free file number
ArchivoNumero = FreeFile
'if checked i'll create the virtual drive at
'windows startup
'chkiniciarconwindows=chkrunatwindowsstartup
If chkiniciarconwindows.Value = 1 Then
    'i get the app path, see above
    If Len(App.Path) = 3 Then
        'since the method i use to create the
        'virtual drive is origally for DOS I use
        'a BAT file to create the virtual drive
        'when windows starts, otherwise when the
        'computer is shutdown the virtua drive
        'will disappear.
        Open App.Path & "CreateVirtualDrive.Bat" For Append As #ArchivoNumero
            'txtdiscovirtual=txtvirtualdrive
            Print #ArchivoNumero, "subst " & txtdiscovirtual.Text & " " & RutaCarpeta
        Close #ArchivoNumero
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "VirtualDrive", App.Path & "CreateVirtualDrive.bat", REG_SZ
    Else
        Open App.Path & "\CreateVirtualDrive.Bat" For Append As #ArchivoNumero
            Print #ArchivoNumero, "subst " & txtdiscovirtual.Text & " " & RutaCarpeta
        Close #ArchivoNumero
        SetKeyValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "VirtualDrive", App.Path & "\CreateVirtualDrive.bat", REG_SZ
    End If
End If
'creates the virtual drive
ret = Shell("subst " & txtdiscovirtual.Text & " " & RutaCarpeta, vbHide)
MsgBox "The virtual drive should be created.", vbInformation, "Virtual Drive"
End Sub

Private Sub cmdinformacion_Click()
'Shows information about me!, read please!!! :-)
frminformacion.Show 1
'frminformacion=frminformation (something like about me)
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdseleccionar_Click()
'RutaCarpeta=FolderPath
'BuscarCarpeta=FindFolder
RutaCarpeta = BuscarCarpeta
lblcarpeta.Caption = RutaCarpeta
End Sub

Function LetraHD(AppPath As String) As String
'this function return the hard-disk letter,
'for example C:\
'LetraHD=Hard-Disk Letter
LetraHD = Mid(AppPath, 1, 3)
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Dim IniciarConWindows As String
'IniciarConWindows=Run at windows startup
'frmprincipal=frmmain
frmprincipal.Caption = "Virtual Drive Creator Version: " & App.Major & "." & App.Minor & "." & App.Revision & " Created by Matías A. Villagarcía"
IniciarConWindows = QueryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "VirtualDrive")
If IniciarConWindows <> "" Then chkiniciarconwindows.Value = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Do you really want to exit?", vbQuestion + vbYesNo + vbDefaultButton2, "Salir") <> vbYes Then Cancel = 1
End Sub
