
&& UnitTestCase
&& @link http://www.simpletest.org/en/unit_test_documentation.html

DEFINE CLASS UnitTestCase As Custom

	FUNCTION start
	
		LOCAL i
		LOCAL iUserMethodsCount, cMethodName
		LOCAL cClassName
		LOCAL cCallback
		
		&& get class name
		ACLASS(aClassNames, THIS)
		cClassName = aClassNames[1]
		
		&& get class method name which match TEST*
		iUserMethodsCount = AMEMBERS(aUserMethods, THIS, 1, 'U')
		FOR i = 1 TO iUserMethodsCount 
			cMethodName = aUserMethods(i,1)
			IF LEFT(cMethodName, 4) = 'TEST'
				cCallback = 'THIS.' + cMethodName
				? 'TEST: ', cClassName+'.'+cMethodName+'()'
				&cCallback
			ENDIF
		ENDFOR
		? 'COMPLETED'
	ENDFUNC
	
	&& Fail if $x == $y is false
	FUNCTION assertEqual
		LPARAMETERS x, y
		IF x == y 
			&& DO NOTHING
		ELSE
			? 'X: ', x
			? 'Y: ', y
			THROW 'Fail.'
		ENDIF
	ENDFUNC
	
	FUNCTION getCRC32
		LPARAMETERS cText
		RETURN SUBSTR(TRANSFORM(VAL(SYS(2007, cText, 0, 1)), '@0x'), 3)
	ENDFUNC
ENDDEFINE
