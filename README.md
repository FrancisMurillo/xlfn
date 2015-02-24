xlfn
----

A pseudo functional programming library for Excel VBA.

After working for some time in an limited Windows environment without Python or Haskell and making Excel VBA modules, wouldn't it be nice to have a small piece of functional programming on VBA to ease the development pain? So with a little magic from <a href="https://msdn.microsoft.com/en-us/library/office/ff197132.aspx">Application.Run</a> and inspiration from Python's <a href="https://github.com/kachayev/fn.py">fn.py</a>, here is a quasi functional programming library for and done purely in VBA.

**To God I plant this seed, may it turn into a forest**

### Introduction

Since VBA doesn't have lambdas or closure or the first class functions, you can't declare function variables. A little cheap but the other workaround for this is that you can declare String variables that contain the function name you want to invoke, so this is the route this library takes. This is done by **Application.Run** which serious flaws. One serious weak point is that functions invoked by **Application.Run** cannot return value. So how do we go about the return mechanism? Cheap but it can be done by declaring a global variable as the return holder and returning this value after the function invokation which is defined here as **Fn.Result**. 

So in short, we have to refit our functions so that they can be "invokable" by **Application.Run**. Take this simple sample addition function in the module **MyModule**.

```VB.net
Public Function Add(A As Long, B as Long) as Long
  Add = A + B
End Function
```

The newly reflavored function with the above workaround

```VB.net
Public Sub Add_(A as Long, B as Long)
  Fn.Result = A + B
End Sub
```

Not much of a difference except the return mechanism and the function header without the return type. Now to invoke this quasi function is done through Fn.Invoke seen here.

```VB.net
  Debug.Print MyModule.Add(1, 2) 
  Debug.Print Fn.Invoke("MyModule.Add_", Array(1, 2))
  Debug.Print Fn.InvokeTwoArgs("MyModule.Add_", 1, 2)
```

Note the way functions are invoked using their full name([Module Name].[Method Name]) and how the arguments are wrapped in the **Array()** function, a little cumbersome but a necessary evil since you can't call the functions straight. Now with this function mechanism, the core functional method Filter, Reduce and Map is now fully available at our disposal as seen here.

```VB.net
Public Sub IsOdd_(Val as Long) 
  Fn.Result = ((Val Mod 2) = 1)
End Sub

Public Sub FilteringUsingFP()
  Dim MyVals as Variant, OddVals as Variant
  MyVals = Array(1, 2, 4, 5, 7, 8, 10)
  
  OddVals = FnArrayUtil.Filter("MyModule.IsOdd_", MyVals)
  Debug.Print ArrayUtil.Print_(OddVals) ' Returns [1, 5, 7]
End Sub

Public Sub FilteringWithoutFP()
  Dim MyVals as Variant
  MyVals = Array(1, 2, 4, 5, 7, 8 , 10)
  
  Dim OddVals as Variant, ValIndex as Long, MyValIndex as Long, MyVal as Long
  OddVals = ArrayUtil.CloneSize(MyVals)
  ValIndex = 0
  
  For MyValIndex = 0 to UBound(MyVals)
    MyVal = MyVals(MyValIndex)
    If ((MyVal Mod 2 ) = 1) Then
      OddVals(ValIndex) = MyVal
      ValIndex = ValIndex + 1
    End if
  Next
  
  If ValIndex = 0 Then
    OddVals = Array()
  Else
    ReDim Preserve OddVals(0 to ValIndex - 1)
  End if
  
  Debug.Print ArrayUtil.Print_(OddVals) ' Same as above
End Sub
```

Compare the code, without using FP there would be some boilerplate just to filter a simple array although it can be still shortened. The mechanism of lambdas here are somewhat cumbersome but with the ability of **Map**, **Filter** and **Reduce** at the ready, it's a small price to pay for these three functional functions. There are others such as **ZipWith**, **Sort**, and so on just to make this worthwhile.

Just a word of warning, these functions might run slower than the longer versions since there is the overhead of Application.Run as well as the transfer mechanisms involved although Python can get the same flak. But if performance is not an issue, then this library is good for you and your sanity.

### Quick Start

This is a <a href="https://github.com/FrancisMurillo/xlchip">chip</a> project, so you can download this via *Chip.ChipOnFromRepo "Fn"* or if you want to install it via importing module. Just import these four modules in your project.

1. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/Fn.bas">Fn.bas</a>
2. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnBuffer.bas">FnBuffer.bas</a>
3. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnUtil.bas">FnArrayUtil.bas</a>
4. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnArrayUtil.bas">FnArrayUtil.bas</a>
5. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnPredicate.bas">FnPredicate.bas</a>
6. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnFunction.bas">FnFunction.bas</a>
7. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnOperator.bas">FnOperator.bas</a>

And include in your project references the following.

1. **Microsoft Visual Basic for Applications Extensibility 5.3** - Any version would do but it has been tested with version 5.3
2. **Microsoft Scripting Runtime** - Just for compatibility sake

So to see if it's working, run in the Intermediate Window or what I call the *terminal*.

```VB.net
  Fn.Hello()
```

You should see in the window output **"Hello Fn: The Pseudo Functional Programming Library for VBA"** in the intermediate window.
