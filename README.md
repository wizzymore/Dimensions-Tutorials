# Dimensions-Tutorials
## Splitting a Loop on diffrent pages

```vbscript
Dim startpos 
Dim numrows 
startpos = 0 
numrows = 5 
While (startpos < LargeGrid.Categories.Count)
    LargeGrid.QuestionFilter = _
LargeGrid.Categories.Mid(startpos, numrows)
    LargeGrid.Ask()
    startpos = startpos + numrows
    LargeGrid.QuestionFilter = NULL
End While
```

## Checking categorical responses

### Operators

>'`=`'

*Tests whether the response exactly matches a specified answer*

> '`<>`'

*Tests whether the response does not exactly match a specified answer*

> '`<`'

*Tests whether the response contains a subset of specified answers but not all of them*

>'`<=`'

*Tests whether the response contains one or more specified answers*

> '`>`'

*Tests whether the response contains all the specified answers and at least one other answer*

> '`>=`'

*Tests whether the response contains all specified answers, with or without additional answers*

> '`*`'

*Tests whether the specified categories intersect with any specified category list and returns the intersection or {} when none intersect*

> `AnswerCount`

*Returns the number of responses chosen*

> `ContainsAll`

*Tests whether the response contains all the specified answers*

> `ContainsAny`

*Tests whether the response contains at least one of the specified answers*

> `ContainsSome`

*Tests whether the response contains a given number of the specified answers*

## All the specified responses chosen
### All specified responses and no others

USING:
```vbscript
Tried "Which of the test products did you try?"
    categorical [1..4]
{
  ProductA "Product A",
  ProductB "Product B",
  ProductC "Product C",
  ProductD "Product D"
};
```

To test whether the respondent chose all the specified responses and no others, type either:

`Qname = {Resp1, Resp2, ... Respn}`

or:

`Qname.ContainsAll({Resp1, Resp2, ... Respn}, true)`

For example:

`Tried = {ProductA, ProductB}`

`Tried.ContainsAll({ProductA, ProductB}, true)`

### All specified responses with or without others
If you want to test whether the respondent chose all the specified responses with or without others, type either:

`Qname >= {Resp1, Resp2, ... Respn}`

or:

`Qname.ContainsAll({Resp1, Resp2, ... Respn})`

For example:

`Tried >= {ProductA, ProductB}`

`Tried.ContainsAll({ProductA, ProductB})`

### All specified responses and at least one other

The final possibility to is test whether the respondent chose all the specified responses with at least one other response. To do this, type either:

`Qname > {Resp1, Resp2, ... Respn}`

For example:

`Tried > {ProductA, ProductB}`

This expression is **True** if the respondent chooses products A and B with at least one of products C and D — that is, A, B, and C, or A, B, and D, or all four products.

## At least one of the specified responses chosen
### At least one specified response and no others
To test whether the respondent chose at least one of the specified responses and no others, type either:

`Qname <= {Resp1, Resp2, ... Respn}`

or:

`Qname.ContainsAny({Resp1, Resp2, ... Respn}, true)`

**Where**:

*Qname is the name of the question whose response you want to check.
*Resp1 to Respn are the names of the responses you want to check for.
(With ContainsAny, the true parameter is passed to the function when it is executed, and is the setting to be applied to the internal boolean variable, exactly. If the variable is True, the function tests for exactly the listed values and no others; if it is False, the function just tests for the presence of all the listed values. See ContainsAny for more information.

In both cases, the expression is **True** if the respondent chooses any of the specified answers at the named question and no others. For example, if the question is defined as:

```vbscript
DaysVisit "On which days do you normally visit the gym?"
 categorical [1..7]
{
   Monday, Tuesday, Wednesday, Thursday,
   Friday, Saturday, Sunday
};
```

**The expressions**:

`DaysVisit <= {Saturday, Sunday}`

`DaysVisit.ContainsAny({Saturday, Sunday}, true)`

Are **True** for all respondents who go to the gym at the weekend only; that is, on Saturday only, on Sunday only, or on both Saturday and Sunday.

The expressions are **False** for people who go to the gym on any weekday even if they also go at the weekend.

### Some but not all specified responses with no others
You can check whether some but not all of the specified responses are chosen, without any other responses:

`Qname < {Resp1, Resp2, ... Respn}`

**The expression**:

`DaysVisit < {Saturday, Sunday, Monday}`

Is **True** if the respondent goes to the gym on one or two of the specified days, but not on all three, and also not on any other days of the week.

In other words, it is **True** for people who go on Saturday, Sunday, or Monday only, or who go on Saturday and Sunday, or on Saturday and Monday, or on Sunday and Monday only.

The expression is **False** for anyone who goes to the gym on Saturday, Sunday, and Monday, or on any other day or days of the week.

### Not just all the specified responses
You can also specify a test that returns True if none, or some but not all, of the listed responses are chosen, with or without other answers. In other words, you want to reject answers that contain the listed responses and nothing else. To do this, type:

`Qname<> {Resp1, Resp2, ... Respn}`

**For example**:

`DaysVisit <> {Saturday, Sunday}`

Is **False** for anyone who goes to the gym on both Saturday and Sunday and not on any other day.

It is **True** for everyone else; that is, for people who go only during the week, and for people who go on Friday and Saturday, and for people who go on Friday, Saturday, and Sunday.

### A given number of specified responses chosen
To test whether the respondent chose a given number of answers from a set of responses, as in, did the respondent choose three track and field sports from a selection of sports, use the ContainsSome function.

**Syntax**

`ContainsSome({Resp1, Resp2, ... Respn}[, Minimum][, Maximum][, No_Others])`

**Parameters**

**Qname**

The name of the question whose response you want to check.

**Resp1 .. Respn**

The names of the responses you want to check for.

**Minimum**

The minimum number of responses that must be chosen from the set in order for ContainsSome to be True. If you omit this parameter, the default is zero meaning that none of the specified responses need be chosen.

**Maximum**

The maximum number of responses that must be chosen from the set in order for ContainsSome to be True. If you omit this parameter, the default is the number of responses specified in the expression.

**No_Others**

No_Others is True if the response may contain only answers from the specified set, or False if other answers may be chosen too. The default is False.See ContainsSome for more information.

**Example**

If the Sports question is defined in the Metadata section as:
```vbscript
Sports "Which sports do you take part in?" categorical [1..]
{
  Running, Cycling, Football, Javelin, Discus,
  HighJump "High jump", LongJump "Long jump",
  Swimming, Basketball, Tennis, Hockey, Rugby,
  PoleVault "Pole vault", Netball, Gymnastics
};
```
and you want to test whether the respondent’s answer contains between three and five track and field events, you would write the following expression:

```vbscript
Sports.ContainsSome({Running,Javelin,Discus,HighJump,LongJump,
    PoleVault}, 3, 5)
```

This expression ignores any other sports that may be present in the response to the Sports question, and has the same effect as:

```vbscript
Sports.ContainsSome({Running,Javelin,Discus,HighJump,LongJump,
    PoleVault}, 3, 5, false)
```

If you want to take other sports into account and test whether the respondent mentioned only three, four, or five track and field events and no other sports, you would write:

```vbscript
Sports.ContainsSome({Running,Javelin,Discus,HighJump,LongJump,
    PoleVault}, 3, 5, true)
```
You can leave the minimum and/or maximum number of required responses undefined, and the interviewing program will use the function’s default settings instead. The default minimum is one, so the expression:

```vbscript
Sports.ContainsSome({Running,Javelin,Discus,HighJump,LongJump,
    PoleVault}, 5)
```
Is **True** if the respondent’s answer contains between one and five of the specified sports.

It is **False** if all six sports or none are mentioned. The results of the expression are unaffected by any other sports mentioned.

Similarly, if the expression is:

```vbscript
Sports.ContainsSome({Running,Javelin,Discus,HighJump,LongJump,
    PoleVault}, 3)
```
The result is **True** if the respondent mentions at least three of the specified sports, regardless of any other sports mentioned.

## Combining Expressions
| Expression |                  Description                    |
|------------|-------------------------------------------------|
|    AND     |Tests whether both expressions are true          |
|    OR      |Tests whether atleast one expression is true     |
|    XOR     |Tests whether one expression or the other is true|

### Both expressions true
When your test uses two expressions that must both be True, combine them using the And operator:

`Expression1 And Expression2`

### At least one expression true
To test whether one expression or the other, or both expressions are True, use the Or operator:

`Expression1 Or Expression2`

### Only one expression true
The Xor operator combines expressions such that only one of the expressions must be True. If both expressions are True then the expression as a whole is False:

`Expression1 Xor Expression2`
