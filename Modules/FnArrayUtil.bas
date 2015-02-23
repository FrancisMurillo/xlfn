Attribute VB_Name = "FnArrayUtil"
' Functional Array Utility
' ------------------------
'
' These utilize the pseudo functions of Fn, rather this gives you the reason to use it.
'
' Stemming from ArrayUtil and Functional Programming, it follows its convention and nuance along with Fn's.
' These methods are a collection of well known Functional methods, happy programming
'
' # Module Contract
'
' MethodNames should be fully qualified as there might be conflict if there is another with the same name.
' Likewise, said methods should follow the argument and return restriction although they are variant
' The type notation is [Arg1, Arg2, ...] -> [Ret], this could be [Var] -> [Int]

' ## Core Functions
'
' The core of functional programs: Map, Reduce and Filter

'# This applies a new array with each element applied to a function
'P MethodName: A function of [Var]->[Var]
'R Retains Base
Public Function Map_(MethodName As String, Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        Map_ = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Arr_ As Variant, Index As Long, Elem_ As Variant
    Arr_ = ArrayUtil.CloneSize(Arr)

    For Index = LBound(Arr_) To UBound(Arr_)
        Elem_ = Arr(Index)
        Arr_(Index) = Fn.Invoke(MethodName, Array(Elem_))
    Next
    
    Map_ = Arr_
End Function

'# This returns a new subarray from an array that satisfies a condition
'P MethodName: A predicate function of [Var]->[Bool], this dictates who gets drafted
'R Zero Base
Public Function Filter_(MethodName As String, Arr As Variant)
    If ArrayUtil.IsEmptyArray(Arr) Then
        Filter_ = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Arr_ As Variant, Index As Long, Elem_ As Variant
    Arr_ = ArrayUtil.CreateWithSize(ArrayUtil.Size(Arr))
    For Each Elem_ In Arr
        If Fn.Invoke(MethodName, Array(Elem_)) Then
            Arr_(Index) = Elem_
            Index = Index + 1
        End If
    Next
    
    If Index = 0 Then
        Arr_ = ArrayUtil.CreateEmptyArray()
    Else
        ReDim Preserve Arr_(0 To Index - 1)
    End If
    
    Filter_ = Arr_
End Function


'# This computes a total for an array
'# This is foldl in functional literature
'P MethodName: An operator function [Var(Acc), Var(Elem)] -> [Var]
'P             Where Acc is the accumulator and elem is the element in question
'P Initial: An optionall value indicating a start value,
'P          if this is empty, the accumulator starts with the first element in the array and starts counting at the second;
'P          otherwise with this
'R Zero Base
Public Function Reduce_(MethodName As String, Arr As Variant, Optional Initial As Variant = Empty) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        Reduce_ = Empty
        Exit Function
    End If
    
    Dim Acc_ As Variant, Index As Long, StartIndex As Long, Elem_ As Variant, IsFirst As Boolean, UseFirst As Boolean
    UseFirst = IsEmpty(Initial)
    Acc_ = IIf(UseFirst, Arr(0), Initial)
    StartIndex = LBound(Arr) + IIf(UseFirst, 1, 0)
    For Index = StartIndex To UBound(Arr)
        Elem_ = Arr(Index)
        Acc_ = Fn.Invoke(MethodName, Array(Acc_, Elem_))
    Next
    
    Reduce_ = Acc_
End Function

' ## Functional Methods
'
' These methods are implementations of well known functional methods such as Zip, Take, Drop and so on.

'# Zip combines several arrays into an arrays of tuple
'C Base Independent
'R Zero Base
Public Function Zip(Arrs As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arrs) Then
        Zip = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If
    
    Dim MinSize As Long, Arr_ As Variant, Size_ As Long
    MinSize = Size(Arrs(0))
    For Each Arr_ In Arrs
        If IsEmptyArray(Arr_) Then
            Zip = ArrayUtil.CreateEmptyArray()
            Exit Function
        Else
            Size_ = Size(Arr_)
            MinSize = IIf(Size_ < MinSize, Size_, MinSize)
        End If
    Next
    
    Dim ZArr As Variant, Tuple As Variant, Index As Long, TIndex As Long, ElemArr As Variant
    ZArr = CreateWithSize(MinSize)
    Tuple = CreateWithSize(UBound(Arrs) + 1)
    
    For Index = 0 To UBound(ZArr)
        For TIndex = 0 To UBound(Arrs)
            Tuple(TIndex) = Arrs(TIndex)(LBound(Arrs(TIndex)) + Index)
        Next
        ZArr(Index) = Tuple
    Next
    
    Zip = ZArr
End Function

'# Zips arrays and applies a function on each element.
'# Zip + Map basically with the same requirement
Public Function ZipWith_(MethodName As String, Arrs As Variant) As Variant
    ZipWith_ = Map_(MethodName, Zip(Arrs))
End Function

'# Joins arrays into one bigger array, pretty much join on each
'# Although not really Functional Method, it falls under the iterators of Python
'C Base Independent
'R Zero Base
Public Function Chain(Arr As Variant) As Variant
    If ArrayUtil.IsEmptyArray(Arr) Then
        Chain = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim CArr As Variant, TSize As Long, Arr_ As Variant, CIndex As Long, Elem_ As Variant
    TSize = 0
    For Each Arr_ In Arr
        TSize = TSize + Size(Arr_)
    Next
    
    If TSize = 0 Then
        Chain = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If
    
    CArr = CreateWithSize(TSize)
    CIndex = 0
    For Each Arr_ In Arr
        For Each Elem_ In Arr_
            CArr(CIndex) = Elem_
            CIndex = CIndex + 1
        Next
    Next
    Chain = CArr
End Function

'# Takes the first N items in an array.
'# If it exceeds the number of elements, it returns all elements
'C Base Independent
'R Zero Base
Public Function TakeN(N As Long, Arr As Variant) As Variant
    If IsEmptyArray(Arr) Then
        TakeN = ArrayUtil.CreateEmptyArray()
        Exit Function
    End If

    Dim Arr_ As Variant, Index As Long, Ctr As Long, Offset As Long, Size_ As Long, MaxSize As Long
    Offset = LBound(Arr)
    MaxSize = Size(Arr)
    Size_ = N
    
    Size_ = IIf(MaxSize < Size_, MaxSize, Size_)
    Arr_ = CreateWithSize(Size_)
    If Size_ = 0 Then Exit Function
    
    For Ctr = 0 To Size_ - 1
        Arr_(Ctr) = Arr(Offset + Ctr)
    Next

    TakeN = Arr_
End Function


