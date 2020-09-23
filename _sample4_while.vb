'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================
'Write, while and Rnd

sub write(n)
    while (n>0)
        print (rnd()*n)
        n--
    wend
    print ("-----------")
end sub

dim a=0
while (a<=10)
    write(a)
    a++
wend

'Infinite Loop with Exit
a=50
while(true)
    if (a<25) then exit()
    print (a)
    a--
wend

