Attribute VB_Name = "Mod_Function_Ptr"
'Module to call any function with any parameters
'(as long as they are from the type long (dword))
'by its address, which is needed for example if you want to
'code an plugin for a C/C++ based program

'This code uses assembler, its easy to debug and see whats going on
'basically it just uses CallWindowProc to execute the machine code and save
'the return value. The return value is directly stored after the first function(code)
'as that was the only place i could think of to store it savely



Private Declare Function CallWindowProc Lib "user32.dll" Alias _
  "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
  ByVal hWnd As Long, ByVal Msg As Long, _
  ByVal WParam As Long, ByVal lParam As Long) As Long
  
Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal Destination As Long, _
    ByVal Source As Long, ByVal Length As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" _
    (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

'ret_func=ptr to function to deliver ret value (callback)
'addr=ptr to function
Public Function CallByAddress(addr As Long, ByRef params() As Long) As Long
On Error GoTo err

Dim offset As Long
Dim code As Long
Dim code2 As Long
Dim ptr_return_value As Long

code = 0
code2 = 0
offset = 0
ptr_return_value = 0

    hHeap = GetProcessHeap()
    
    code = HeapAlloc(hHeap, 0, (UBound(params) - LBound(params) + 1) * 5 + 22)
    
    'Pushing arguments on the stack...
    
    i = UBound(params)
    
    For i = LBound(params) To UBound(params)
            
        CopyMemory code + offset, VarPtr(104), 1 'PUSH
        
            offset = offset + 1
    
        CopyMemory code + offset, VarPtr(params(i)), 4 'Next param
        
            offset = offset + 4
            
    Next i
 
err:    'Jump over parameters (I know, bad coding...)

On Error Resume Next

    If code = 0 Then code = HeapAlloc(hHeap, 0, 22)
    
    'Preparing CALL...
    
    CopyMemory code + offset, VarPtr(184), 1  'MOV EAX
    
        offset = offset + 1
    
    CopyMemory code + offset, VarPtr(addr), 4 'addr to call
    
        offset = offset + 4
    
    
    
    CopyMemory code + offset, VarPtr(255), 1 '...
    
        offset = offset + 1
    
    CopyMemory code + offset, VarPtr(208), 1 'CALL EAX
    
        offset = offset + 1
    
    
    
    
    'Preparing return value...
    
    hHeap = GetProcessHeap()
    
    code2 = HeapAlloc(hHeap, 0, 10)
    
    
    CopyMemory code + offset, VarPtr(139), 1  '...
    
        offset = offset + 1
        
    CopyMemory code + offset, VarPtr(216), 1  'MOV EBX,EAX
    
        offset = offset + 1
 
    CopyMemory code + offset, VarPtr(184), 1  'MOV EAX,...
    
        offset = offset + 1
           
    CopyMemory code + offset, VarPtr(code2), 4 '...addr to call
    
        offset = offset + 4
        
       
    CopyMemory code + offset, VarPtr(255), 1 '...
    
        offset = offset + 1
    
    CopyMemory code + offset, VarPtr(208), 1 'CALL EAX
    
        offset = offset + 1
        
    CopyMemory code + offset, VarPtr(195), 1 'RET
    
        offset = offset + 1
        
      ptr_return_value = code + offset
        
    'Setting up 2nd functio which will handle the return value
    
    offset = 0
    
    'getting address from where the call happened
    CopyMemory code2 + offset, VarPtr(88), 1 'POP EAX
    
        offset = offset + 1
        
    CopyMemory code2 + offset, VarPtr(80), 1 'PUSH EAX
    
        offset = offset + 1
       
    CopyMemory code2 + offset, VarPtr(137), 1 '...
    
        offset = offset + 1
    
    CopyMemory code2 + offset, VarPtr(88), 1 '...
    
        offset = offset + 1
    
    CopyMemory code2 + offset, VarPtr(1), 1 'MOV [EAX+1],EBX
    
        offset = offset + 1
    
    CopyMemory code2 + offset, VarPtr(195), 1 'RET
    
        offset = offset + 1
    
    'return value successfully copied!
        
       
    CallWindowProc code, 0&, 0&, 0&, 0&
    
    'Set the return value
    CopyMemory VarPtr(CallByAddress), ptr_return_value, 4
    
    HeapFree hHeap, 0&, code
    HeapFree hHeap, 0&, code2


End Function

Public Function blub(ByVal lng As Long, ByVal lng1 As Long, ByVal lng2 As Long) As Long

If lng = 1 Then

    blub = lng1 + lng2
Else

    blub = lng1 * lng2
    
End If

End Function
