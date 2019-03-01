; ###########################################################################################################################################################################################
; 배열 작은 값부터 차례대로 정렬
; https://www.google.com/search?newwindow=1&source=hp&ei=kIGDWpXYLdLajwOtnLT4DA&q=autohotkey+array+sort&oq=autohotkey+array+sort&gs_l=psy-ab.3..0i22i30k1.3181.7947.0.8239.32.20.0.0.0.0.148.1671.8j8.17.0....0...1.1.64.psy-ab..19.13.1372.6..0j35i39k1j0i20i264k1j0i203k1.104.qKKh-koFig0	

; ###########################################################################################################################################################################################






/*
; 실행 부분
MySelect := new Quickselect

;nth smallest number
MyArray := [9, 8, 7, 6, 0, 2, 3]
MsgBox, %  MySelect.MaxValue(MyArray)   ;9
    . "`n" MySelect.MinValue(MyArray)   ;0
    . "`n" MySelect.Select(MyArray, 5)  ;7

;nth item alphabetically
MyArray := ["quickselect", "ahk", "script", "demo"]
MsgBox, %  MySelect.MinValue(MyArray)   ;ahk
    . "`n" MySelect.Select(MyArray, 3)  ;quickselect
    . "`n" MySelect.Select(MyArray, 2)  ;demo
    . "`n" MySelect.MaxValue(MyArray)   ;script
return
*/






/*  Quickselect
 *  by kon
 *  May 10, 2014
 *
 *		MaxValue(InArray)
 *			Returns the largest value in an array.
 *
 *		MinValue(InArray)
 *			Returns the smallest value in an array.
 *
 *		Select(InArray, n)
 *      	Returns the nth smallest value in an array.
 */
class Quickselect {
   
    MaxValue(InArray) {
        return this.Select(InArray, InArray.MaxIndex())
    }
   
    MinValue(InArray) {
        return this.Select(InArray, 1)
    }
   
    Partition(InArray, Left, Right, PivotIndex) {
        PivotValue := InArray[PivotIndex]
        , this.Swap(InArray, pivotIndex, Right)
        , StoreIndex := Left
        , i := Left - 1
        Loop, % Right - Left
            if (InArray[j := i + A_Index] <= PivotValue)
                this.Swap(InArray, StoreIndex, j)
                , StoreIndex++
        this.Swap(InArray, Right, StoreIndex)
        return StoreIndex
    }

    Select(InArray, n, Left:=1) {
        Right := InArray.MaxIndex()
        if (Left = Right)
            return InArray[Left]
        SortArray := InArray.Clone()
        Loop {
            PivotIndex := (Left + Right) // 2
            , PivotIndex := this.Partition(SortArray, Left, Right, PivotIndex)
            if (n = PivotIndex)
                return SortArray[n]
            else if (n < PivotIndex)
                Right := PivotIndex - 1
            else
                Left := PivotIndex + 1
        }
    }

    Swap(InArray, i1, i2) {
        t := InArray[i1]
        , InArray[i1] := InArray[i2]
        , InArray[i2] := t
    }
}