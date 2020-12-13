if exists("b:current_syntax")
    finish
endif

let b:current_syntax = "vba"

syntax keyword vbaKeyword  If Then Else ElseIf End For Each Next In With While Loop
                         \ And Or Not True False
                         \ Set Dim Call As Private Public Const Enum
                         \ Long String Boolean Long Integer Variant Collection Byte Single Dboule Currency Date Object Type
                         \ ByVal ByRef
                         \ Worksheet
                         \ Option Explicit Function Sub Exit
                         \ Formula New Workbook Value
syntax keyword vbaFunction Cells Range Print CUInt CStr CBool CByte CCur CDate CDbl CInt CLng CSng CVar
syntax match   vbaComment "\v'.*$"
syntax match   vbaOperator "\v\*"
syntax match   vbaOperator "\v/"
syntax match   vbaOperator "\v\+"
syntax match   vbaOperator "\v-"
syntax match   vbaOperator "\v\="
syntax region  vbaString start=/\v"/ end=/\v"/
highlight link vbaKeyword  Keyword
highlight link vbaFunction Function
highlight link vbaComment  Comment
highlight link vbaOperator Operator
highlight link vbaString   String

