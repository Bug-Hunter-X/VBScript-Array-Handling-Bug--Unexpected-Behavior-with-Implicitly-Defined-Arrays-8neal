Function processArray(arr)
  For Each item In arr
    Debug.Print item
  Next
End Function

Dim myArray(1 To 2) ' Explicitly dimension the array
myArray(1) = "apple"
myArray(2) = "banana"

processArray myArray

'Another example with a different way to create an array
Dim arr(5)
For i = 1 To 5
arr(i) = i
Next
processArray arr

'Passing an array literal
processArray Array("One", "Two", "Three")