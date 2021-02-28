if exists("b:current_syntax")
    finish
endif

let b:current_syntax = "vba"

syntax keyword vbaType        Long String Boolean Long Integer Variant Collection Byte Single Dboule Currency Date Object
                            \ Workbook Worksheet
                            \ Type LongPtr
syntax keyword vbaBoolean     True False
syntax keyword vbaConditional If Then Else ElseIf Select Case
syntax keyword vbaRepeat      For Each Next Loop While Do
syntax keyword vbaKeyword     End In With
                            \ And Or Not
                            \ Set Dim Call As Private Public Const Enum
                            \ ByVal ByRef
                            \ Debug
                            \ Option Explicit Function Sub Exit
                            \ Formula New Value Sleep Lib Module Declare PtrSafe
                            \ On GoTo
                            \ Nothing DoEvents
syntax keyword vbaError       Error
syntax keyword vbaFunction    Cells Range Print CUInt CStr CBool CByte CCur CDate CDbl CInt CLng CSng CVar MsgBox Delete
syntax match vbaComment       "\v'.*$"
syntax match vbaOperator      "\v\*"
syntax match vbaOperator      "\v/"
syntax match vbaOperator      "\v\+"
syntax match vbaOperator      "\v-"
syntax match vbaOperator      "\v\="
syntax region  vbaString      start=/\v"/ skip=/\v\\"/ end=/\v"/
highlight link vbaConditional Conditional
highlight link vbaType        Type
highlight link vbaBoolean     Boolean
highlight link vbaRepeat      Repeat
highlight link vbaKeyword     Keyword
highlight link vbaFunction    Function
highlight link vbaComment     Comment
highlight link vbaOperator    Operator
highlight link vbaString      String
highlight link vbaError       Exception

