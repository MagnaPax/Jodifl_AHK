
AutoSort(Arr) {
    t:=Object()
    for k, v in Arr
        t[RegExReplace(v,"\s")]:=v
    for k, v in t
        Arr[A_Index]:=v
    return Arr
}