'================================================================================
' Part of DC3 Compiler - Interpreter
' Author: Lorenzi Davide (http://www.hexagora.com)
' See the file 'license.txt' for informations
'================================================================================

sub writenum(n)
	if n>0 then
		print (n)
		writenum(n-1)
		print(n)
	end if
end sub

writenum(10)



