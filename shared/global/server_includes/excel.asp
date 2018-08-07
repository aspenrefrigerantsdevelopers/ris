<%
Dim MyFont
set MyFont = xlw.CreateFont
If xlw.Version = 0 Then MyFont.Name = "Tahoma"
MyFont.Size = 10

'---- Number styles ----
Dim NumStyle
Set NumStyle = xlw.CreateStyle
NumStyle.Font = MyFont
NumStyle.HorizontalAlignment = sahaLeft
NumStyle.Number = "#,##0"

Dim NumStyle2
Set NumStyle2 = xlw.CreateStyle
NumStyle2.Font = MyFont
NumStyle2.HorizontalAlignment = sahaRight
NumStyle2.Number = "#,##0"

Dim NumStyle3
Set NumStyle3 = xlw.CreateStyle
NumStyle3.Font = MyFont
NumStyle3.HorizontalAlignment = sahaRight
NumStyle3.Font.Bold = True
NumStyle3.Number = "#,##0"

Dim NumStyle4
Set NumStyle4 = xlw.CreateStyle
NumStyle4.Font = MyFont
NumStyle4.HorizontalAlignment = sahaRight
NumStyle4.Number = "#,##0.00"

Dim NumStyle5
Set NumStyle5 = xlw.CreateStyle
NumStyle5.Font = MyFont
NumStyle5.HorizontalAlignment = sahaRight
NumStyle5.Font.Bold = True
NumStyle5.Number = "#,##0.00"

Dim NumStyle6
Set NumStyle6 = xlw.CreateStyle
NumStyle6.Font = MyFont
NumStyle6.HorizontalAlignment = sahaRight
NumStyle6.Number = "#,##0.0##"

Dim NumStyle7
Set NumStyle7 = xlw.CreateStyle
NumStyle7.Font = MyFont
NumStyle7.HorizontalAlignment = sahaRight
NumStyle7.Font.Bold = True
NumStyle7.Number = "#,##0.0##"

Dim NumStyle8
Set NumStyle8 = xlw.CreateStyle
NumStyle8.Font = MyFont
NumStyle8.HorizontalAlignment = sahaLeft
NumStyle8.Number = "#0"

'---- Date styles ----
Dim DateStyle
Set DateStyle = xlw.CreateStyle
DateStyle.Font = MyFont
DateStyle.Number = 14					' m/d/y
DateStyle.HorizontalAlignment = sahaLeft

Dim DateTimeStyle
Set DateTimeStyle = xlw.CreateStyle
DateTimeStyle.Font = MyFont
DateTimeStyle.Number = "[$-409]m/d/yyyy h:mm AM/PM;@"
DateTimeStyle.HorizontalAlignment = sahaLeft

'---- Text styles ----
Dim TxtStyle
Set TxtStyle = xlw.CreateStyle
TxtStyle.Font = MyFont
TxtStyle.HorizontalAlignment = sahaLeft

Dim TxtStyle2
Set TxtStyle2 = xlw.CreateStyle
TxtStyle2.Font = MyFont
TxtStyle2.HorizontalAlignment = sahaRight

'---- Currency styles ----
Dim CurrencyStyle
Set CurrencyStyle = xlw.CreateStyle
CurrencyStyle.Font = MyFont
CurrencyStyle.Number = "$#,##0.00"
CurrencyStyle.HorizontalAlignment = sahaRight

Dim CurrencyStyle2
Set CurrencyStyle2 = xlw.CreateStyle
CurrencyStyle2.Font = MyFont
CurrencyStyle2.Number = "$#,##0.00"
CurrencyStyle2.HorizontalAlignment = sahaRight
CurrencyStyle2.Font.Bold = True

'---- Heading styles ----
Dim HeadingStyle
Set HeadingStyle = xlw.CreateStyle
HeadingStyle.Font = MyFont
HeadingStyle.Font.Bold = true
HeadingStyle.Font.Size = 10

Dim HeadingStyle2
Set HeadingStyle2 = xlw.CreateStyle
HeadingStyle2.Font = MyFont
HeadingStyle2.Font.Bold = true
HeadingStyle2.Font.Size = 10
HeadingStyle2.HorizontalAlignment = sahaRight

'---- Title styles ----
Dim TitleStyle
Set TitleStyle = xlw.CreateStyle
TitleStyle.Font = MyFont
TitleStyle.Font.Bold = true
TitleStyle.Font.Size = 14

Dim TitleStyle2
Set TitleStyle2 = xlw.CreateStyle
TitleStyle2.Font = MyFont
TitleStyle2.Font.Bold = true
TitleStyle2.Font.Size = 14
TitleStyle2.HorizontalAlignment = sahaRight
%>