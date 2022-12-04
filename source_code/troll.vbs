Option Explicit

Function one() 'step 1
    Dim a
    Do 'check if "a" is "i am stupid"
        If a="i am stupid" Then
            Exit Do
        Else
            a=inputbox("Step 1:input 'i am stupid' or you will never finish the first step!","yep you are stupid")
        End If
    Loop
End Function

Function two() 'step 2
    Dim c
    c=150
    Do until c=0 'print user's IQ from 150 to 0
        MsgBox "your IQ is "&c,4096,"yep you are stupid"
        c=c-1
    Loop
End Function

Function three() 'step 3
    Dim yee,wsh
    yee=0
    Set wsh=wscript.createobject("wscript.shell")
    Do Until yee=50 'open 50 calculators and something more
        wsh.run "calc"
        wsh.run "msedge https://www.bilibili.com/video/BV1GJ411x7h7/?spm_id_from=333.337.search-card.all.click"
        yee=yee+1
    Loop
End Function

Function four() 'step 4
    Dim WSHshell
    Set WSHshell=wscript.createobject("wscript.shell")
    WSHshell.run "shutdown -s -t 10",0,true 'shutdown the computer after 10 seconds
End Function

Do 'the way to ridicule
    MsgBox "well well well",4096,"yep you are stupid"
    MsgBox "i'm sorry for you for opening the program",4096,"yep you are stupid"
    MsgBox "are you ready to be trolled?",4096,"yep you are stupid"
    Call one()

    MsgBox "fine",4096,"yep you are stupid"
    MsgBox "but i won't let you go easily",4096,"yep you are stupid"
    Call two()

    MsgBox "i'm surprised that you have came over that",4096,"yep you are stupid"
    MsgBox "but you know,i won't stop until you give up",4096,"yep you are stupid"
    Call three()

    MsgBox "now the final one hehhehhehhehh",4096,"yep you are stupid"
    MsgBox "wait for 10 seconds......",4096,"yep you are stupid"
    Call four()
Loop