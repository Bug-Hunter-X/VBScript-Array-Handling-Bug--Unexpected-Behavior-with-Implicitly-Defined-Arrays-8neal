# VBScript Array Handling Bug

This repository demonstrates a subtle bug in VBScript related to how arrays are handled when passed as function parameters.  The issue arises when an array is not explicitly dimensioned using the `Dim` statement before being passed to a function.  This can lead to unexpected behavior and runtime errors, particularly when using `For Each` loops.

## Bug Description

VBScript's weak typing system can mask this error.  If an array is created without explicit dimensioning using `Dim`, VBScript defaults to a base index of 0.  This behavior differs from explicitly dimensioned arrays, which start at index 1.  This difference can lead to off-by-one errors or runtime exceptions when processing the array within a function.

## Reproduction Steps

1.  Run `bug.vbs`.
2.  Observe the output (or the lack thereof) and errors, if any.
3.  Run `bugSolution.vbs` to see the corrected code.

## Solution

Always explicitly dimension your arrays using the `Dim` statement to avoid unexpected behavior and ensure consistent indexing.  The `bugSolution.vbs` file provides a corrected version of the code.  Explicitly stating the bounds of the array prevents issues caused by implicit array creation.