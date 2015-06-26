VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMemoryPatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' CMemoryPatcher.cls
' Revision 1

Option Explicit

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_ALL_ACCESS As Long = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF

Private Const PAGE_READWRITE As Long = &H4

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                                                     ByVal bInheritHandle As Long, _
                                                     ByVal dwProcessID As Long) As Long
Private Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, _
                                                          ByVal lpAddress As Long, _
                                                          ByVal dwSize As Long, _
                                                          ByVal flNewProtect As Long, _
                                                          lpflOldProtect As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
                                                           ByVal lpBaseAddress As Long, _
                                                           lpBuffer As Any, _
                                                           ByVal nSize As Long, _
                                                           lpNumberOfBytesRead As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, _
                                                            ByVal lpBaseAddress As Long, _
                                                            lpBuffer As Any, _
                                                            ByVal nSize As Long, _
                                                            lpNumberOfBytesWritten As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, _
                                                                lpdwProcessID As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long


Private m_dwPatchAddress As Long
Private m_strPatchBuffer As String
Private m_strOldBuffer As String
Private m_dwPatchLength As Long

Private m_dwProcessID As Long
Private m_hWindowHandle As Long

Public Sub SetPatch(ByVal strPatch As String, ByVal dwPatchLength As Long)
    m_dwPatchLength = dwPatchLength
    m_strPatchBuffer = strPatch
End Sub

Public Function ApplyPatch(ByVal strWindowTitle) As Boolean
    ApplyPatch = Patch(strWindowTitle, True)
End Function

Public Function RemovePatch(ByVal strWindowTitle) As Boolean
    RemovePatch = Patch(strWindowTitle, False)
End Function

Public Sub SetPatchAddress(ByVal dwPatchAddress As Long)
    m_dwPatchAddress = dwPatchAddress
End Sub

Public Function FindAddress(ByVal strWindowTitle As String, ByVal strSearch As String, ByVal dwStartAddress As Long, ByVal dwEndAddress As Long, ByVal dwInterval As Long, Optional ByVal dwOptionalOffset As Long = 0) As Long
    Dim hProcessHandle As Long
    Dim strReadContents As String
    Dim dwBytesRead As Long

    Call GetProcessByName(strWindowTitle)

    If m_dwProcessID Then
        hProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, 0, m_dwProcessID)

        If hProcessHandle Then
            Dim dwSearchAddress As Long
            
            For dwSearchAddress = dwStartAddress To dwEndAddress Step dwInterval
                Call ReadProcessMemory(hProcessHandle, (dwSearchAddress + dwOptionalOffset), StrPtr(strReadContents), Len(strSearch), dwBytesRead)
                
                If strReadContents = strSearch Then
                    FindAddress = dwSearchAddress + dwOptionalOffset
                    Exit Function
                End If

                dwSearchAddress = dwSearchAddress + dwInterval
            Next dwSearchAddress

            Call CloseHandle(hProcessHandle)
        End If
    End If
    
    FindAddress = 0
End Function

Private Function GetProcessByName(ByVal strWindowTitle As String) As Long
    m_hWindowHandle = FindWindow(vbNullString, strWindowTitle)
    
    If m_hWindowHandle > 0 Then
        GetProcessByName = GetWindowThreadProcessId(m_hWindowHandle, m_dwProcessID)
        Exit Function
    End If

    GetProcessByName = 0
End Function

Public Function ReadDwordFromMemory(ByVal strWindowTitle As String, ByVal dwBaseAddress As Long) As Long
    Dim tempString As String
    tempString = ReadBytesFromMemory(strWindowTitle, dwBaseAddress, 4)
    'Debug.Print "DWORD as String -> " & tempString
    'Debug.Print DebugOutput(tempString)
    ReadDwordFromMemory = GetDWORD(tempString)
End Function

Public Function ReadWordFromMemory(ByVal strWindowTitle As String, ByVal dwBaseAddress As Long) As Long
    Dim tempString As String
    tempString = ReadBytesFromMemory(strWindowTitle, dwBaseAddress, 2)
    'Debug.Print "DWORD as String -> " & tempString
    'Debug.Print DebugOutput(tempString)
    ReadWordFromMemory = GetWORD(tempString)
End Function

Public Function ReadByteFromMemory(ByVal strWindowTitle As String, ByVal dwBaseAddress As Long) As Long
    Dim tempString As String
    tempString = ReadBytesFromMemory(strWindowTitle, dwBaseAddress, 1)
    'Debug.Print "DWORD as String -> " & tempString
    'Debug.Print DebugOutput(tempString)
    If Len(tempString) = 0 Then tempString = Chr(0)
    ReadByteFromMemory = Asc(tempString)
End Function

Public Function ReadBytesFromMemory(ByVal strWindowTitle As String, ByVal dwBaseAddress As Long, ByVal dwLength As Long) As String
    Dim hProcessHandle As Long
    Dim returnString As String
    
    Call GetProcessByName(strWindowTitle)
    
    If m_dwProcessID Then
        hProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, 0, m_dwProcessID)
        
        If hProcessHandle Then
            Dim dwOldMemoryAttributes As Long
            
            If VirtualProtectEx(hProcessHandle, dwBaseAddress, dwLength, PAGE_READWRITE, dwOldMemoryAttributes) Then
                returnString = String(dwLength, 0)
                Call ReadProcessMemory(hProcessHandle, dwBaseAddress, ByVal returnString, dwLength, 0&)
                
                ReadBytesFromMemory = returnString
                Exit Function
            End If
        End If
    End If
End Function

Public Function ReadBytesFromMemoryUntilChar(ByVal strWindowTitle As String, ByVal dwBaseAddress As Long, ByVal strDelimiter As String, Optional ByVal IncludeDelimiter As Boolean = False) As String
    Dim hProcessHandle As Long

    Dim i As Long
    Dim tempByte As String * 1
    Dim returnString As String
    
    Call GetProcessByName(strWindowTitle)
    
    If m_dwProcessID Then
        hProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, 0, m_dwProcessID)
        
        If hProcessHandle Then
            Do
                Dim dwOldMemoryAttributes As Long
            
                If VirtualProtectEx(hProcessHandle, dwBaseAddress + i, 1, PAGE_READWRITE, dwOldMemoryAttributes) Then
                    Call ReadProcessMemory(hProcessHandle, dwBaseAddress + i, ByVal tempByte, 1, 0&)
                    returnString = returnString + tempByte
                    i = i + 1
                End If
            Loop While tempByte <> strDelimiter
            
            If (IncludeDelimiter = False) And (Len(returnString) > 0) Then returnString = Left$(returnString, Len(returnString) - 1)
            
            ReadBytesFromMemoryUntilChar = returnString
            Exit Function
        End If
    End If
End Function

Private Function Patch(ByVal strWindowTitle As String, ByVal bApply As Boolean) As Boolean
    Patch = False
    Dim hProcessHandle As Long
    Dim bCompleted As Long
    
    Call GetProcessByName(strWindowTitle)
    
    If m_dwProcessID Then
        hProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, 0, m_dwProcessID)
        
        If hProcessHandle Then
            Dim dwOldMemoryAttributes As Long
            
            If VirtualProtectEx(hProcessHandle, m_dwPatchAddress, m_dwPatchLength, PAGE_READWRITE, dwOldMemoryAttributes) Then
                Dim dwWrittenBytes As Long
                Dim dwBytesRead As Long
                
                If bApply Then
                    m_strOldBuffer = String(m_dwPatchLength, 0)
                    Call ReadProcessMemory(hProcessHandle, m_dwPatchAddress, ByVal m_strOldBuffer, m_dwPatchLength, dwBytesRead)
                Else
                    m_strPatchBuffer = String(m_dwPatchLength, 0)
                    Call ReadProcessMemory(hProcessHandle, m_dwPatchAddress, ByVal m_strPatchBuffer, m_dwPatchLength, dwBytesRead)
                End If
                
                If bApply Then
                    If WriteProcessMemory(hProcessHandle, m_dwPatchAddress, ByVal m_strPatchBuffer, m_dwPatchLength, dwWrittenBytes) Then
                        bCompleted = True
                    End If
                Else
                    If WriteProcessMemory(hProcessHandle, m_dwPatchAddress, ByVal m_strOldBuffer, m_dwPatchLength, dwWrittenBytes) Then
                        bCompleted = True
                    End If
                End If

                Call VirtualProtectEx(hProcessHandle, m_dwPatchAddress, m_dwPatchLength, dwOldMemoryAttributes, dwOldMemoryAttributes)
            End If
            
            Call CloseHandle(hProcessHandle)
        End If
    End If
    
    If bCompleted Then Patch = True
End Function

Private Function DebugOutput(ByVal sIn As String) As String
   Dim x1 As Long, y1 As Long
   Dim iLen As Long, iPos As Long
   Dim sB As String, sT As String
   Dim sOut As String
   
   'build random string to display
   'y1 = 384
   'sIn = String(y1, 0)
   'For x1 = 1 To y1
   '    Mid(sIn, x1, 1) = Chr(255 * Rnd())
   'Next x1
   
   iLen = Len(sIn)
   If iLen = 0 Then Exit Function
   sOut = ""
   For x1 = 0 To ((iLen - 1) \ 16)
       sB = String(48, " ")
       sT = "................"
       For y1 = 1 To 16
           iPos = 16 * x1 + y1
           If iPos > iLen Then Exit For
           Mid(sB, 3 * (y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
           Select Case Asc(Mid(sIn, iPos, 1))
           Case 32 To 255
               Mid(sT, y1, 1) = Mid(sIn, iPos, 1)
           End Select
       Next y1
       If Len(sOut) > 0 Then sOut = sOut & vbCrLf
       sOut = sOut & sB & "  " & sT
   Next x1
   DebugOutput = sOut
End Function

Private Function GetDWORD(Data As String) As Long
    'Dim tmp As String
    'tmp = ToHex(Data)
    'Dim A As String, b As String, c As String, d As String
    'A = Mid(tmp, 7, 2)
    'b = Mid(tmp, 5, 2)
    'c = Mid(tmp, 3, 2)
    'd = Mid(tmp, 1, 2)
    'tmp = A & b & c & d
    'GetDWORD = Val("&H" & tmp)
    CopyMemory GetDWORD, ByVal Data, 4
End Function

Private Function GetWORD(Data As String) As Long
    'Dim tmp As String
    'tmp = ToHex(Data)
    'Dim A As String, b As String
    'A = Mid(tmp, 3, 2)
    'b = Mid(tmp, 1, 2)
    'tmp = A & b
    'GetWORD = Val("&H" & tmp)
    CopyMemory GetWORD, ByVal Data, 2
End Function

Private Function ToHex(Data As String) As String
    Dim i As Integer
    For i = 1 To Len(Data)
        ToHex = ToHex & Right("00" & Hex(Asc(Mid(Data, i, 1))), 2)
    Next i
End Function

Private Sub Class_Initialize()
    m_dwProcessID = 0
End Sub

