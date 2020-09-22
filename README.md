<div align="center">

## Eval \(Evaluate String Expression\) \*REPOST\*


</div>

### Description

This is a recursive function that evaluates strings expressions. It supports multiple levels of parenthesis, algebraic evaluation of expressions (in this example

exponentiation ^ has same level of multiplication and division), function calls, logical operators, string/date/numeric functions and expresion evaluation. This is the base for

the creation of a scripting language.
 
### More Info
 
Logical evaluations requires that expressions be inside parenthesis. Example: ((-1) and (-1)) or (-1)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aldo Vargas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aldo-vargas.md)
**Level**          |Intermediate
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aldo-vargas-eval-evaluate-string-expression-repost__1-21856/archive/master.zip)





### Source Code

```
Public Function Eval(expr As String)
 Dim value As Variant, operand As String
 Dim pos As Integer
 pos = 1
 Do Until pos > Len(expr)
  Select Case Mid(expr, pos, 3)
   Case "not", "or ", "and", "xor", "eqv", "imp"
   operand = Mid(expr, pos, 3)
   pos = pos + 3
  End Select
  Select Case Mid(expr, pos, 1)
   Case " "
    pos = pos + 1
   Case "&", "+", "-", "*", "/", "\", "^"
    operand = Mid(expr, pos, 1)
    pos = pos + 1
   Case ">", "<", "=":
    Select Case Mid(expr, pos + 1, 1)
     Case "<", ">", "="
      operand = Mid(expr, pos, 2)
      pos = pos + 1
     Case Else
      operand = Mid(expr, pos, 1)
    End Select
    pos = pos + 1
   Case Else
    Select Case operand
    Case "": value = Token(expr, pos)
    Case "&": Eval = Eval & value
         value = Token(expr, pos)
    Case "+": Eval = Eval + value
         value = Token(expr, pos)
    Case "-": Eval = Eval + value
         value = -Token(expr, pos)
    Case "*": value = value * Token(expr, pos)
    Case "/": value = value / Token(expr, pos)
    Case "\": value = value \ Token(expr, pos)
    Case "^": value = value ^ Token(expr, pos)
    Case "not": Eval = Eval + value
         value = Not Token(expr, pos)
    Case "and": value = value And Token(expr, pos)
    Case "or ": value = value Or Token(expr, pos)
    Case "xor": value = value Xor Token(expr, pos)
    Case "eqv": value = value Eqv Token(expr, pos)
    Case "imp": value = value Imp Token(expr, pos)
    Case "=", "==": value = value = Token(expr, pos)
    Case ">": value = value > Token(expr, pos)
    Case "<": value = value < Token(expr, pos)
    Case ">=", "=>": value = value >= Token(expr, pos)
    Case "<=", "=<": value = value <= Token(expr, pos)
    Case "<>": value = value <> Token(expr, pos)
    End Select
  End Select
 Loop
 Eval = Eval + value
End Function
Private Function Token(expr, pos)
 Dim char As String, value As String, fn As String
 Dim es As Integer, pl As Integer
 Const QUOTE As String = """"
 Do Until pos > Len(expr)
  char = Mid(expr, pos, 1)
  Select Case char
  Case "&", "+", "-", "/", "\", "*", "^", " ", ">", "<", "=": Exit Do
  Case "("
   pl = 1
   pos = pos + 1
   es = pos
   Do Until pl = 0 Or pos > Len(expr)
    char = Mid(expr, pos, 1)
    Select Case char
     Case "(": pl = pl + 1
     Case ")": pl = pl - 1
    End Select
    pos = pos + 1
   Loop
   value = Mid(expr, es, pos - es - 1)
   fn = LCase(Token)
   Select Case fn
    Case "sin": Token = Sin(Eval(value))
    Case "cos": Token = Cos(Eval(value))
    Case "tan": Token = Tan(Eval(value))
    Case "exp": Token = Exp(Eval(value))
    Case "log": Token = Log(Eval(value))
    Case "atn": Token = Atn(Eval(value))
    Case "abs": Token = Abs(Eval(value))
    Case "sgn": Token = Sgn(Eval(value))
    Case "sqr": Token = Sqr(Eval(value))
    Case "rnd": Token = Rnd(Eval(value))
    Case "int": Token = Int(Eval(value))
    Case "day": Token = Day(Eval(value))
    Case "month": Token = Month(Eval(value))
    Case "year": Token = Year(Eval(value))
    Case "weekday": Token = WeekDay(Eval(value))
    Case "hour": Token = Hour(Eval(value))
    Case "minute": Token = Minute(Eval(value))
    Case "second": Token = Second(Eval(value))
    Case "date": Token = Date
    Case "date$": Token = Date$
    Case "time": Token = Time
    Case "time$": Token = Time$
    Case "timer": Token = Timer
    Case "now": Token = Now()
    Case "len": Token = Len(Eval(value))
    Case "trim": Token = Trim(Eval(value))
    Case "ltrim": Token = LTrim(Eval(value))
    Case "rtrim": Token = RTrim(Eval(value))
    Case "ucase": Token = UCase(Eval(value))
    Case "lcase": Token = LCase(Eval(value))
    Case "val": Token = Val(Eval(value))
    Case "chr": Token = Chr(Eval(value))
    Case "asc": Token = Asc(Eval(value))
    Case "space": Token = Space(Eval(value))
    Case "hex": Token = Hex(Eval(value))
    Case "oct": Token = Oct(Eval(value))
    Case "environ": Token = Environ$(Eval(value))
    Case "curdir": Token = CurDir$
    Case "dir": If Len(value) Then Token = Dir(Eval(value)) Else Token = Dir
    Case Else: Token = Eval(value)
   End Select
   Exit Do
  Case QUOTE
   pl = 1
   pos = pos + 1
   es = pos
   Do Until pl = 0 Or pos > Len(expr)
    char = Mid(expr, pos, 1)
    pos = pos + 1
    If char = QUOTE Then
     If Mid(expr, pos, 1) = QUOTE Then
      value = value & QUOTE
      pos = pos + 1
     Else
      Exit Do
     End If
    Else
     value = value & char
    End If
   Loop
   Token = value
   Exit Do
  Case Else
   Token = Token & char
   pos = pos + 1
  End Select
 Loop
 If IsNumeric(Token) Then
  Token = Val(Token)
 ElseIf IsDate(Token) Then
  Token = CDate(Token)
 End If
End Function
```

