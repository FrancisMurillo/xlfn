Attribute VB_Name = "FnBuffer"
' Fn Lambda Buffers
' -----------------
'
' This is a private module to support pseudo lambdas for Fn.
' So none of the functions here should be invoked manually. It should be done with Fn.Invoke
'
' Since there is no lambda functions, we can simulate it by creating a huge premade functions
' and giving them parameters. Basically, a pseudo lambda with precreated buffers.

' ## Consants
'
' Error constants
Private Const ERR_SOURCE As String = "FnBuffer"
Private Const ERR_OFFSET As Long = 2100

Private Const BUFFER_COUNT As Long = 10
Public Const BUFFER_PREFIX As String = "Buffer_"

Public Const CURRY_METHOD As String = "Curry_"
Public Const COMPOSE_METHOD As String = "Compose_"
Public Const REINVOKE_METHOD As String = "Reinvoke_"
Public Const LAMBDA_METHOD As String = "Lambda_"

Private gIsBufferReady As Boolean
Private gBufferIndex As Long
Private gBufferArgs As Variant

'# Gets the next buffer index available at the same time incrementing it
Public Function GetNextBufferIndex() As Long
    GetNextBufferIndex = gBufferIndex
    gBufferIndex = gBufferIndex + 1
    If gBufferIndex >= BUFFER_COUNT Then _
        gBufferIndex = gBufferIndex - BUFFER_COUNT
End Function

'# Moves the buffer one step back
'# This is an advanced function to prevent the buffer from being overloaded
'# Not to be used unless understood
Public Function ReleaseCurrentBuffer()
    gBufferArgs(gBufferIndex) = Array()
    gBufferIndex = gBufferIndex - 1
    If gBufferIndex < 0 Then _
        gBufferIndex = BUFFER_COUNT - 1
End Function

' ## Function Buffers CRUD
'
' These functions provide the mechanism to combine functions

'# Prepares the buffers for action
Public Sub InitializeBuffers()
    If Not gIsBufferReady Then _
        gBufferArgs = ArrayUtil.CreateWithSize(BUFFER_COUNT)
    gIsBufferReady = True
End Sub

'# Resets the flag so that the buffer arg can be updated.
'# Useful when there is a new buffer to update but not supposed to be used during prod
Public Sub ResetBuffers()
    gIsBufferReady = False
End Sub

'# Checks if the lambda functions are used before they are prepared
Private Sub CheckIfReady()
    If Not gIsBufferReady Then _
        Err.Raise vbObjectError + ERR_OFFSET, ERR_SOURCE, "Tried to access the buffer function manually."
End Sub

Public Sub SetBuffer(Args As Variant, Index As Long)
    gBufferArgs(Index) = Args
End Sub
Public Sub SetClosureBufferArgs(BufferArgs As Variant, Index As Long)
    gBufferArgs(Index)(3) = BufferArgs
End Sub

' ## The Function Buffers
'
' The actual buffers that pass the invokation

Private Sub Buffer_0(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(0)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_1(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(1)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_2(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(2)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_3(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(3)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_4(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(4)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_5(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(5)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_6(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(6)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_7(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(7)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_8(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(8)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub

Private Sub Buffer_9(Optional Args As Variant = Empty)
    CheckIfReady
    Dim BufferArgs As Variant
    BufferArgs = gBufferArgs(9)
    If IsMissing(Args) Then _
        Args = ArrayUtil.CreateEmptyArray()
    Fn.Result = Fn.Invoke(CStr(BufferArgs(0)), Array(BufferArgs(1), BufferArgs(2), Args, BufferArgs(3), BufferArgs(4)))
End Sub



