Function using an array as a parameter may fail if the array is not passed correctly.  VBScript is weakly typed, so this error might not be caught during development.  Consider this example:

```vbscript
Function processArray(arr)
  For Each item In arr
    Debug.Print item
  Next
End Function

Dim myArray(1 To 2)
myArray(1) = "apple"
myArray(2) = "banana"

processArray myArray
```

This function will work correctly.  However, if the array is not explicitly defined with `Dim` before passing:

```vbscript
Function processArray(arr)
  For Each item In arr
    Debug.Print item
  Next
End Function

Dim myArray
myArray(0) = "apple"
myArray(1) = "banana"

processArray myArray
```

This will fail, as VBScript creates the array with index 0 instead of index 1. This subtle difference may lead to unexpected results or runtime errors, often causing the `For Each` loop to run into trouble.