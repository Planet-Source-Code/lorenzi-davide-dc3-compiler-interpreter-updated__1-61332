'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================

dim a=10,b=20
a=10+2*5+b 'operator precedence

function pluto (par1)
    print ("Hello!")
    pluto=par1
end function

msgbox (pluto(a))




