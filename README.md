# Ledger Consolidation Algorithm
An Algorithm aimed to consolidate two arrays (size: 10 - 100) containing doubles representing currency

## Background

By the end of the month, accountants would consolidate the company and bank ledgers. The aim is to check for any mismatched items, list these items, and declare compensations so that two sides can consolidate.

This task is made into an algorithmic task for the ideal situation: Given two arrays, $L_b$ and $L_c$, with size varying from 10 - 100, each element containing a double type number ranging from -1,000,000 to 1,000,000, try to create an algorithm that returns items collections from both sides that cannot find a sum of an element or a subset from the other side.

## Problem

The most straight forward way is to calculate the combination of all possibilities, but then the list size would be astronomical. The formula for combination a subset size of r and a total size of n is:
$$
rCn = \frac{n!}{r!\times(n-r)!}
$$
To get all the possible combinations of size n, the formula is
$$
\begin{align}
	R&=\sum_{k=1}^n kCn\\
	R&=\sum_{k=1}^n \frac{n!}{k!\times(n-k)!}

\end{align}
$$
To calculate the relationship for all possible combinations versus size, I wrote

```vb
Sub Cacl()
    Dim rng As Range
    Dim result As Long
    Dim i As Integer
    Dim n As Integer
    
    Set rng = Me.[A2]
    
    Do While rng.Value <> ""
        n = rng.Value
        result = 0
        For i = 1 To n
            result = result + Single_rCn(i, n)
        
        
        Next i
        rng.Offset(0, 1).Value = result
    
    
        Set rng = rng.Offset(1, 0)
    Loop




End Sub

Function Single_rCn(CombiSize As Integer, TotalSize As Integer) As Long
    Single_rCn = WorksheetFunction.Fact(TotalSize) / _
        (WorksheetFunction.Fact(CombiSize) * _
        WorksheetFunction.Fact(TotalSize - CombiSize))

End Function
```

The time to calculate all combination possibilities is:
$$
O=2^n-1
$$
Therefore combination on larger sizes by brute force is not recommended.

