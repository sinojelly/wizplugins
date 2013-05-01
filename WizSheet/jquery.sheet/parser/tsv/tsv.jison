/* description: Parses a tab separated value to an array, value being parsed is expected to have ':::::' at beginning */

/* lexical grammar */
%lex

/* lexical states */
%s PRE_QUOTE QUOTE STRING

/*begin lexing */
%%

/* quote handling */
<QUOTE>(\n|"\n")                    {return 'CHAR';}
<QUOTE>([\'\"])(?=<<EOF>>) {
    this.popState();
    return 'QUOTE_OFF';
}
<QUOTE>([\'\"]) {
    if (yytext == this.quoteChar) {
        this.popState();
        this.begin('STRING');
        return 'QUOTE_OFF';
	} else {
	    return 'CHAR';
	}
}
<QUOTE>(?=(\t)) {
	this.popState();
	this.begin('STRING');
	return 'CHAR';
}
(":::::")([\'\"]) {
    this.quoteChar = yytext.substring(5);
	this.begin('QUOTE');
	return 'QUOTE_ON';
}
<PRE_QUOTE>([\'\"]) {
    this.quoteChar = yytext;
    this.popState();
	this.begin('QUOTE');
	return 'QUOTE_ON';
}
(\t|"\t")(?=[\'\"]) {
	this.begin('PRE_QUOTE');
	return 'COLUMN_STRING';
}
(\n|"\n")(?=[\'\"]) {
	this.begin('PRE_QUOTE');
	return 'END_OF_LINE';
}
<QUOTE>(.)	                        {return 'CHAR';}
/* end quote handling */


/*spreadsheet control characters*/
<STRING>(\n\n|"\n\n") {
	this.popState();
	return 'END_OF_LINE_WITH_NO_COLUMNS';
}
<STRING>(\n|"\n") {
	this.popState();
	return 'END_OF_LINE';
}
<STRING>(\t|"\t") {
	this.popState();
	return 'COLUMN_STRING';
}
<STRING>(.)                         {return 'CHAR';}
(':::::')                           {return 'BOF';}
(\n\n|"\n\n")                       {return 'END_OF_LINE_WITH_NO_COLUMNS';}
(\t\n|"\t\n")                       {return 'END_OF_LINE_WITH_EMPTY_COLUMN';}
(\t|"\t")                           {return 'COLUMN_EMPTY';}
(\n|"\n")                           {return 'END_OF_LINE';}
(.) {
	this.begin('STRING');
	return 'CHAR';
}
<<EOF>> {
    //lexer.yy.conditionStack = [];
    return 'EOF';
}

/* end lexing */
/lex


%start grid

%% /* language grammar */

grid :
	rows EOF {
		return $1;
	}
	| EOF {
		return [['']];
	}
;

rows :
	row {
		$$ = [$1];
	}
	| END_OF_LINE {
        $$ = [];
	}
	| END_OF_LINE_WITH_NO_COLUMNS {
	    $$ = [''];
	}
    | END_OF_LINE_WITH_EMPTY_COLUMN {
        $$ = [''];
    }
    | rows END_OF_LINE {
        $$ = $1;
    }
    | rows END_OF_LINE_WITH_NO_COLUMNS {
        $1.push(['']);
        $$ = $1;
    }
    | rows END_OF_LINE_WITH_EMPTY_COLUMN {
        $1[$1.length - 1].push('');
        $$ = $1;
    }
    | rows END_OF_LINE row {
        $1.push($3);
        $$ = $1;
    }
    | rows END_OF_LINE_WITH_NO_COLUMNS row {
        $1.push(['']);
        $1.push($3);
        $$ = $1;
    }
    | rows END_OF_LINE_WITH_EMPTY_COLUMN row {
        $1[$1.length - 1].push('');
        $1.push($3);
        $$ = $1;
    }
;

row :
	string {
		$$ = [$1.join('')];
	}
	| COLUMN_EMPTY {
		$$ = [''];
	}
	| COLUMN_STRING {
	}
    | row COLUMN_EMPTY {
        $1.push('');
        $$ = $1;
    }
    | row COLUMN_STRING {
        $$ = $1;
    }
    | row COLUMN_EMPTY string {
        $1.push('');
        $1.push($3.join(''));
        $$ = $1;
    }
    | row COLUMN_STRING string {
        $1.push($3.join(''));
        $$ = $1;
    }
;

string :
	BOF {
		$$ = [];
	}
	| CHAR {
		$$ = [$1];
	}
	| string CHAR {
		$1.push($2);
		$$ = $1;
	}
	| QUOTE_ON string QUOTE_OFF {
         $$ = $2;
     }
;