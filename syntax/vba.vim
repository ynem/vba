if exists("b:current_syntax")
    finish
endif

let b:current_syntax = "vba"

syntax keyword vbaKeyword  If Then Else ElseIf End For Each Next In With While Loop Case
                         \ And Or Not True False
                         \ Set Dim Call As Private Public Const Enum
                         \ Long String Boolean Long Integer Variant Collection Byte Single Dboule Currency Date Object Type
                         \ ByVal ByRef
                         \ Worksheet Debug Select
                         \ Option Explicit Function Sub Exit
                         \ Formula New Workbook Value Sleep Lib Module Declare PtrSafe LongPtr
                         \ On Error GoTo ErrorHandler
                         \ Nothing DoEvents
syntax keyword vbaFunction Cells Range Print CUInt CStr CBool CByte CCur CDate CDbl CInt CLng CSng CVar MsgBox Delete
syntax match   vbaComment "\v'.*$"
syntax match   vbaOperator "\v\*"
syntax match   vbaOperator "\v/"
syntax match   vbaOperator "\v\+"
syntax match   vbaOperator "\v-"
syntax match   vbaOperator "\v\="
syntax region  vbaString start=/\v"/ skip=/\v\\"/ end=/\v"/
highlight link vbaKeyword  Keyword
highlight link vbaFunction Function
highlight link vbaComment  Comment
highlight link vbaOperator Operator
highlight link vbaString   String

