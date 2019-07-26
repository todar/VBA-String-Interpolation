### "This is {0} cool!", "freaking"

I've always wanted an easy and intuitive way to inject variables into a string. So after about 10 variations, I finally came up with this function.

---

### How it works

The concept is that I can find every pattern such as `{key}` or `{0}` or whatever `{taco}` and get a unique list of these. If there are more keys than there are variables it will raise a custom error. (*Keys are case sensitive.*)

If it matches then it uses the index of each `ParamArray` variable and matches that to the index of the pattern list. For example `"{bacon} {burrito}"` bacon: 0, burrito: 1.

With that match, it simply replaces every instance of the match with the value of the variable.

---

### A few extra notes

I originally had the pattern start with a dollar sign `${0}` to copy JavaScripts syntax but decided to keep it shorter for simplicity.

It does use the escape character `\`. Example: `\{test}` would be print `{test}`.

It also includes shortcuts for vbNewLine `\n` and vbTab `\t`.

---

### The formula

Make sure to first set references to **Microsoft Scripting Runtime** and **Microsoft VBScript Regular Expressions 5.5**.

I thought about doing this late binding but figured performance is probably better with these libraries referenced and they are common enough that it should not matter.

```vb
' Returns a new cloned string that replaced special {keys} with its associated pair value.
' Keys can be anything since it goes off of the index, so variables must be in proper order!
' Can't have whitespace in the key.
' Also Replaces "\t" with VbTab and "\n" with VbNewLine
'
' @author: Robert Todar <https://github.com/todar>
' @reference: Microsoft Scripting Runtime - [Dictionary]
' @reference: Microsoft VBScript Regular Expressions 5.5 - [RegExp, Match]
' @example: Inject("Hello, {name}!\nJS Object = {name: {name}, age: {age}}\n", "Robert", 31)
Public Function Inject(ByVal source As String, ParamArray values() As Variant) As String
    
    ' Want to get a copy and not mutate original
    Inject = source
    
    Dim regEx As RegExp
    Set regEx = New RegExp ' Late Binding would be: CreateObject("vbscript.regexp")
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        
        ' This section is only when user passes in variables
        If Not IsMissing(values) Then
            
            ' Looking for pattern like: {key}
            ' First capture group is the full pattern: {key}
            ' Second capture group is just the name:    key
            .Pattern = "(?:^|[^\\])(\{([\w\d\s]*)\})"
            
            ' Used to make sure there are even number of uniqueKeys and values.
            Dim keys As New Scripting.Dictionary
            
            Dim keyMatch As match
            For Each keyMatch In .Execute(Inject)
                
                Debug.Print
                
                ' Extract key name
                Dim key As Variant
                key = keyMatch.submatches(1)
                
                ' Only want to increment on unique keys.
                If Not keys.Exists(key) Then
                    
                    If (keys.Count) > UBound(values) Then
                        Err.Raise 9, "Inject", "Inject expects an equal amount of keys to values. Keys found: " & Join(keys.keys, ", ") & ", " & key
                    End If
                    
                    ' Replace {key} with the pairing value.
                    Inject = Replace(Inject, keyMatch.submatches(0), values(keys.Count))
                    
                    ' Add key to make sure it isn't looped again.
                    keys.Add key, vbNullString
                    
               End If
            Next
        End If
        
        ' Replace extra special characters. Must allow code above to run first!
        .Pattern = "(^|[^\\])\{"
        Inject = .Replace(Inject, "$1" & "{")
    
        .Pattern = "(^|[^\\])\\t"
        Inject = .Replace(Inject, "$1" & vbTab)
        
        .Pattern = "(^|[^\\])\\n"
        Inject = .Replace(Inject, "$1" & vbNewLine)
        
        .Pattern = "(^|[^\\])\\"
        Inject = .Replace(Inject, "$1" & "")
    End With
    
End Function
```


---

### The tests for it

My first test is using [RegExr.com](https://regexr.com/4i4r7) to see if my pattern would match. Just to note, I use the first capture group as the actual replacement so the characters before will not be replaced.

[![RegExr Tests][1]][1]

---

The next step was to try it in VBA. I just copied the same lines and printed them to the immediate window.

```vba
Private Sub testingInject()
    Debug.Print Inject("{it} works with with words.", "It")
    Debug.Print Inject("{0} works with digits.", "It")
    Debug.Print Inject("{it } works with whitespace.", "It")
    Debug.Print Inject("{ {it} } doesn't effect outer nestings.", "It")
    Debug.Print Inject("\{it} should be escaped.", "It did not but")
    Debug.Print Inject("Hello, {name}! {name}, \{(escaped) you} are {age} years old!.", "Robert", 31)
    Debug.Print Inject("Hello, {name}!\n{\n\tname: {name},\n\t age: {age}\n}", "Robert", 31)
    
    On Error Resume Next 'Expect this to fail
    Debug.Print Inject("Hello, {name}! How are you {Name}", "Robert")
    Debug.Print Err.Description
End Sub
```

Here are the results. They printed how I expected them to.

[![Immediate Window Results][2]][2]

---

### What others could contribute to.

- **Performance** I want to make sure I'm not missing anything that might be a big trade-off of using this.
- **RegEx Check** I am not the best at this and would love to get better at writing these. This is a good example I feel for learning.
- **Improvements** is there anything I'm missing? Could this become even cooler?
- **Possible bugs** really are there any tests I should be running more than what I have.
- **Anything really** I want to continue to learn and grow as a programmer. =)


  [1]: https://i.stack.imgur.com/XTSX7.png
  [2]: https://i.stack.imgur.com/nhN4s.png
