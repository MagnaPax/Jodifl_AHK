#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


class OOP{

	; this 사용 시 전역변수 사용 가능
	Array := object()
	
	value :=
	
	AddingValuesInArray(){
		this.Array.Insert("1")
		this.Array.Insert("2")
		this.Array.Insert("3")
	}
	
	ReadingValuesFromArray(){
		
		 ; 배열로부터 읽기 첫 번째 방법
		Loop % this.Array.Maxindex(){
			MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
		}
	}
	
	InsertValueInTheVariable(){
		this.value := "Hello World"
	}
	
	ReadingValueInTheVariable(){
		MsgBox, % This.value
	}
}




class Overriding_n_Overloading{
	
	Overriding_OriginalValueIs123(){
		MsgBox, 123
	}
		
}



class ExtensTest extends Overriding_n_Overloading{

	; Overriding
	Overriding_OriginalValueIs123(){
		MsgBox, 456
	}
}
