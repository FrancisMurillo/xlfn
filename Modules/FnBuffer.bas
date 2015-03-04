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

Private Const BUFFER_COUNT As Long = 100
Public Const BUFFER_PREFIX As String = "Buffer_"


' Main buffer constants
Public Const BUFFER_MAIN_METHOD As String = "FnBuffer.__Buffer__"
Public Const BUFFER_MAIN_DELIMITER As String = "-"

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
        InitializeBuffers
End Sub

Public Sub SetBuffer(Args As Variant, Index As Long)
    CheckIfReady
    gBufferArgs(Index) = Args
End Sub
Public Sub SetClosureBufferArgs(BufferArgs As Variant, Index As Long)
    gBufferArgs(Index)(3) = BufferArgs
End Sub

' ## The Function Buffers
'
' The actual buffers that pass the invokation

' Public Sub BufferMain(BufferFs As String, InnerFs As String, PreArgs As Variant, Args As Variant, ClosureVars As Variant)
Public Sub BufferMain(MainArgs As Variant, BufferIndex As Long)
    Dim BufferArgs_ As Variant
    Dim Lambda_Fs As String, Inner_Fs As String, PreArgs As Variant, ClosureVars As Variant
    BufferArgs_ = gBufferArgs(BufferIndex)
    Lambda_Fs = BufferArgs_(0)
    Inner_Fs = BufferArgs_(1)
    PreArgs = BufferArgs_(2)
    ClosureVars = BufferArgs_(3)
    
    Fn.Result = Fn.Invoke(Lambda_Fs, Array(Inner_Fs, PreArgs, MainArgs, ClosureVars, BufferIndex))
End Sub

'# Generates the correct buffer lambda for invokation
Public Function GenerateBufferLambda(LambdaFs As String, InnerFs As String, PreArgs As Variant, ClosureVars As Variant) As String
    Dim BufferFs As String, BufferIndex_ As Long
    BufferIndex_ = GetNextBufferIndex
    BufferFs = Join(Array(BUFFER_MAIN_METHOD, BufferIndex_), BUFFER_MAIN_DELIMITER)
    SetBuffer Array(LambdaFs, InnerFs, PreArgs, ClosureVars), BufferIndex_
    GenerateBufferLambda = BufferFs
End Function
