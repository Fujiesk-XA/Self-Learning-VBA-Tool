Public Function targetLotSize(descriptionTar As String) As Double

'Lot Size is the middle number
'Return Split Text Sample: Material Number /Name/ 230 / 75
'Name might include / as well

  Dim LArray() As String
  Dim iiii As Integer
  Dim LText As String

      LText = descriptionTar
      LArray() = Split(LText, "/")
      iiii = UBound(LArray, 1)
      if iiii > 4 then
        i = 0
        Do while i=iiii
          if LArray(i) < 4000000 And LArray(i) > 229  then 
            targetLotSize = LArray(i)
            Exit Loop
          End if
          i = i + 1
        Loop
      else
        targetLotSize = LArray(2)
      end if
      
End Function
