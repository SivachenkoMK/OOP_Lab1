grammar LabCalculator;

/*
* Parser Rules
*/
compileUnit : expression EOF;
expression :
LPAREN expression RPAREN #ParenthesizedExpr
|SUBTRACT LPAREN expression RPAREN #UnaryMinusExpr
|operatorToken=(INCREMENT|DECREMENT) LPAREN expression RPAREN #IncDecExpr
|MAX LPAREN exp1=expression COMMA exp2=expression RPAREN #MaxExpr
|MIN LPAREN exp1=expression COMMA exp2=expression RPAREN #MinExpr
|expression EXPONENT expression #ExponentialExpr
|expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
| expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
|NUMBER #NumberExpr
|IDENTIFIER #IdentifierExpr;

/*
 * Lexer Rules
 */

NUMBER : INT (',' INT)?; 
IDENTIFIER : [a-zA-Z]+[0-9]+;

INT : ('0'..'9')+;

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
INCREMENT : 'inc';
DECREMENT : 'dec';
MAX : 'max';
MIN : 'min';
LPAREN : '(';
RPAREN : ')';
COMMA : ',';

WS : [ \t\r\n] -> channel(HIDDEN);