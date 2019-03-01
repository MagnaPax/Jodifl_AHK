; ########################################################################################
; https://autohotkey.com/boards/viewtopic.php?t=23286
; 배열에서 값을 찾아서 그 위치를 반환
; ########################################################################################


HasVal(haystack, needle) {
	if !(IsObject(haystack)) || (haystack.Length() = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}