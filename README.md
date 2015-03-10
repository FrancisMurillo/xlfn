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

Not much of a difference except the return mechanism and the function header without the return type. Now to invoke this quasi function is done through **Fn.Invoke** seen here.

```VB.net
  Debug.Print MyModule.Add(1, 2) 
  Debug.Print Fn.Invoke("MyModule.Add_", Array(1, 2))
  Debug.Print Fn.InvokeTwoArgs("MyModule.Add_", 1, 2)
```

Note the way functions are invoked using their full name([Module Name].[Method Name]) and how the arguments are wrapped in the **Array()** function, a little cumbersome but a necessary evil since you can't call the functions straight. Now with this function mechanism, the core functional method Filter, Reduce and Map can be implemented nicely. A sample of **FnArrayUtil.Filter** is shown here.

```VB.net
' Lambda definition
Public Sub IsOdd_(Val as Long) 
  Fn.Result = ((Val Mod 2) = 1)
End Sub

' Functional Programming Style
Public Sub FilteringUsingFP()
  Dim MyVals as Variant, OddVals as Variant
  MyVals = Array(1, 2, 4, 5, 7, 8, 10)
  
  OddVals = FnArrayUtil.Filter("MyModule.IsOdd_", MyVals) ' Filter at work
  Debug.Print ArrayUtil.Print_(OddVals) ' Returns [1, 5, 7]
End Sub

' Vanilla VBA with ArrayUtil
Public Sub FilteringWithoutFP()
  Dim MyVals as Variant
  MyVals = Array(1, 2, 4, 5, 7, 8 , 10)
  
  ' Boilerplate code to filter an array
  Dim OddVals as Variant, ValIndex as Long, MyValIndex as Long, MyVal as Long
  OddVals = ArrayUtil.CloneSize(MyVals) ' Create an array with the same LBound and UBound as MyVals
  ValIndex = 0
  
  ' Looping the array boilerplate
  For MyValIndex = 0 to UBound(MyVals)
    MyVal = MyVals(MyValIndex)
    If ((MyVal Mod 2 ) = 1) Then ' Filtering here
      OddVals(ValIndex) = MyVal ' Adding an array the hard way
      ValIndex = ValIndex + 1 
    End if
  Next
  
  ' Trimming the array size, can be put in a method as well but more efficient in this form
  If ValIndex = 0 Then
    OddVals = Array()
  Else
    ReDim Preserve OddVals(0 to ValIndex - 1)
  End if
  
  Debug.Print ArrayUtil.Print_(OddVals) ' Same as above
End Sub
```

Compare the code, without using FP there would be some boilerplate just to filter a simple array although it can be still shortened. The mechanism of lambdas here are somewhat cumbersome but with the ability of **Map**, **Filter** and **Reduce** at the ready, it's a small price to pay for these three functional functions. There are others such as **ZipWith**, **Sort**, and so on just to make this worthwhile.

Just a word of warning, these functions might run slower than the longer versions since there is the overhead of **Application.Run** as well as the transfer mechanisms involved although Python can get the same flak. But if performance is not an issue, then this library is good for you and your sanity.

### Quick Start

This is a <a href="https://github.com/FrancisMurillo/xlchip">chip</a> project, so you can download this via *Chip.ChipOnFromRepo "Fn"* or if you want to install it via importing module. Just import these four modules in your project.

Dependency

- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlbutil/master/Modules/ArrayUtil.bas">ArrayUtil.bas</a> - Since **xlfn** is built-on ArrayUtil, this module is it's only dependency as well as to build on the Map, Reduce and Filter using Arrays. It is recommended to get <a href="https://github.com/FrancisMurillo/xlbutil">xlbutil</a> to avoid this missing module definition

Core

- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/Fn.bas">Fn.bas</a> - The core module, this module creates and runs the pseudo functions described above via **Fn.Invoke**. Aside from the normal invokation, this allows the creation of **composite pseudo functions** that allow the functional concept of currying, closures and composition in a certain way.
- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnArrayUtil.bas">FnArrayUtil.bas</a> - Not a true dependency but this module hosts all the core functions this module was built upon specially the promise of functional application.

Optional
- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnIterator.bas">FnIterator.bas</a> - A toy module to replicate iterators or generators from Python, it's not practical to use than a while-loop but hey it's nice to know.
- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnPredicate.bas">FnPredicate.bas</a> - A set of pseudo predicate functions for common conditions
- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnFunction.bas">FnFunction.bas</a> - A set of pseudo functions that are useful for Map operations or whatnot
- <a href="https://raw.githubusercontent.com/FrancisMurillo/xlfn/master/Modules/FnOperator.bas">FnOperator.bas</a> - A set of pseudo operator functions or binary functions that encapsulate the common operators, useful for ZipWith like operations

And include in your project references the following.

1. **Microsoft Visual Basic for Applications Extensibility 5.3** - Any version would do but it has been tested with version 5.3
2. **Microsoft Scripting Runtime** - Just for compatibility sake

So to see if it's working, run in the Intermediate Window or what I call the *terminal*.

```VB.net
  Fn.Hello()
```

You should see in the window output **"Hello Fn: The Pseudo Functional Programming Library for VBA"** in the intermediate window.

### Composite Functions And More Functionality

We can stop with just **Fn.Invoke** and be happy with **FnArrayUtil** for the three major functions but we can go deeper and get a few more pieces of the functional power. Within the bounds of VBA, they are <a href="https://en.wikipedia.org/wiki/Currying">Currying</a>, <a href="https://en.wikipedia.org/wiki/Function_composition_%28computer_science%29">Composition</a>, <a href="https://en.wikipedia.org/wiki/Closure_%28computer_programming%29">Closure</a> and probably more. If you don't know these concepts, I encourage you to check out the links.

A quick sample for each major C concept.

```VB.net
' Let's call this module MyModule again for reference
' Note: The arguments are wrapped in Array() rather than using Varargs since it complicates parameter and argument passing. 

Public Sub Curring()
  ' Let's curry the add operator
  Dim AddTwoFp as Variant
  AddTwoFp = Fn.Curry("MyModule.Add_", Array(2))
  Debug.Print Fn.InvokeOneArg(AddTwoFp, 1) ' Outputs 3
  Debug.Print Fn.InvokeTwoArgs("MyModule.Add_", 2, 1) ' Like above
  
  Dim AddTwoAndThreeFp as Variant
  AddTwoAndThreeFp = Fn.Curry(AddTwoFp, Array(3)) 
  ' OR AddTwoAndThreeFp = Fn.Curry("MyModule.Add_", Array(2, 3)) 
  Debug.Print Fn.InvokeNoArg(AddTwoAndThreeFp) ' Outputs 5
  Debug.Print Fn.InvokeTwoArgs("MyModule.Add_", 2, 3) ' Like above
End Sub
Public Sub Add_(LVal as Variant, RVal as Variant) 
  Fn.Result = (LVal + RVal)
End Sub

Public Sub Composition()
  ' Let's compose strings of functions to make this example better
  Dim PipelineFp as Variant
  PipelineFp = Fn.Compose(Array("MyModule.Format_", "MyModule.Negative_", "MyModule.Add_"))
  Debug.Print Fn.InvokeTwoArgs(PipelineFp, 2, 3) ' Outputs Value: -5
  Debug.Print Fn.InvokeOneArg(MyModule.Format_", Fn.InvokeOneArg("MyModule.Negative_", Fn.InvokeTwoArgs("MyModule.Add_", 2, 3))) ' Same as above
  Debug.Print Composed(2, 3) ' Or simple like this defined normal function
End Sub
Public Sub Negative_(Val as Variant)
  Fn.Result = -1 * Val ' For safety purposes but -Val is okay
End Sub
Public Sub Format_(Val as Variant)
  Fn.Result = "Value: " & Val
End Sub
Public Function Composed(LVal as Variant, RVal as Variant)
  Composed = "Value: " & (-(LVal + RVal))
End Function

Public Sub Closure()
  ' Let's replicate a counter generator for this example
  Dim CountFromTenFp as Variant
  CountFromTenFp = Closure_(10)
  
  Debug.Print Fn.InvokeNoArgs(CountFromTenFp) ' Outputs 10
  Debug.Print Fn.InvokeNoArgs(CountFromTenFp) ' Outputs 11
  Debug.Print Fn.InvokeNoArgs(CountFromTenFp) ' Outputs 12
End Sub
Public Function Closure_(Start_ as Long) as Variant
  Closure_ = Fn.CreateLambda("MyModule.Counter_Fn", Empty, Empty, Start_) ' This creates a function with a closure variable of Start_
End Funciton
Public Sub Counter_Fn(Optional Args as Variant = Empty)
  ' Optional Args is required when defining composite functions like this which uses Closure or PreArgs although we won't be using the arguments
  Fn.Result = Fn.Closure
  Fn.Closure = Fn.Closure + 1
End Sub
```

