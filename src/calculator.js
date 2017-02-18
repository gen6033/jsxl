
Utils = require ("./utils")
function isNumber(x){
  if(typeof(x) != 'number' && typeof(x) != 'string'){
    return false;
  }else{
    return (x == parseFloat(x) && isFinite(x));
  }
}
function checkNumber(x){
  if(!isNumber(x)){
    throw new Error("NaN");
  }
  return x;
}
function isString(obj) {
  return (typeof (obj) === "string" || obj instanceof String);
}
function checkString(x){
  if(!isString(x)){
    throw new Error("NaN");
  }
  return x;
}

function yyerror(msg) {
  console.log(msg);
}


module.exports = function(worksheet, formula){

  var ans;

/* Prototype file of JavaScript parser.
 * Written by MORI Koichiro
 * This file is PUBLIC DOMAIN.
 */

var buffer;
var token;
var toktype;

var YYERRTOK = 256;
var CELL = 257;
var STRING = 258;
var NUMBER = 259;
var IDENT = 260;
var TRUE = 261;
var FALSE = 262;
var LE = 263;
var GE = 264;
var NEQ = 265;
var SUM = 266;

  
/*
  #define yyclearin (yychar = -1)
  #define yyerrok (yyerrflag = 0)
  #define YYRECOVERING (yyerrflag != 0)
  #define YYERROR  goto yyerrlab
*/


/** Debug mode flag **/
var yydebug = false;

/** lexical element object **/
var yylval = null;

/** Dialog window **/
var yywin = null;
var yydoc = null;

function yydocopen() {
  if (yywin == null) {
    yywin = window.open("", "yaccdiag", "resizable,status,width=600,height=400");
    yydoc = null;
  }
  if (yydoc == null)
    yydoc = yywin.document;
  yydoc.open();
}

function yyprintln(msg)
{
  if (yydoc == null)
    yydocopen();
  yydoc.write(msg + "<br>");
}

function yyflush()
{
  if (yydoc != null) {
    yydoc.close();
    yydoc = null;
    yywin = null;
  }
}



var yytranslate = [
      0,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   16,   18,
     18,   18,   14,   12,   18,   13,   18,   15,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
      9,   11,   10,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   17,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,   18,   18,   18,   18,
     18,   18,   18,   18,   18,   18,    1,   18,    2,    3,
     18,    4,    5,    6,    7,    8,   18
  ];

var YYBADCH = 18;
var YYMAXLEX = 267;
var YYTERMS = 18;
var YYNONTERMS = 5;

var yyaction = [
  -32767,-32767,-32767,-32767,-32767,-32767,   14,   15,    8,    9,
     10,   11,   12,   13,   14,   15,   33,   32,   31,   30,
     16,   17,   18,   19,    0,   19
  ];

var YYLAST = 26;

var yycheck = [
      6,    7,    8,    9,   10,   11,   12,   13,    6,    7,
      8,    9,   10,   11,   12,   13,    2,    3,    4,    5,
     14,   15,   16,   17,    0,   17
  ];

var yybase = [
     14,    2,   -6,   -6,   -6,   -6,   -6,   -6,   14,   14,
     14,   14,   14,   14,   14,   14,   14,   14,   14,   14,
      6,    6,   24,    8,    8,    8,    0,    6,    6,    6,
      6,    6,    6,    6
  ];

var YY2TBLSTATE = 8;

var yydefault = [
  32767,    1,   15,   17,   19,   14,   16,   18,32767,32767,
  32767,32767,32767,32767,32767,32767,32767,32767,32767,32767,
      8,    9,32767,   10,   11,   13
  ];



var yygoto = [
      2,    3,    4,    5,    6,    7,   20,   21,   23,   24,
     25,   38
  ];

var YYGLAST = 12;

var yygcheck = [
      2,    2,    2,    2,    2,    2,    2,    2,    2,    2,
      2,    2
  ];

var yygbase = [
      0,    0,   -8,    0,    0
  ];

var yygdefault = [
  -32768,   22,    1,   28,   29
  ];

var yylhs = [
      0,    1,    2,    2,    3,    3,    3,    3,    4,    4,
      4,    4,    4,    4,    4,    4,    4,    4,    4,    4
  ];

var yylen = [
      1,    1,    1,    1,    1,    1,    1,    1,    3,    3,
      3,    3,    3,    3,    3,    3,    3,    3,    3,    3
  ];

var YYSTATES = 33;
var YYNLSTATES = 26;
var YYINTERRTOK = 1;
var YYUNEXPECTED = 32767;
var YYDEFAULT = -32766;

/*
 * Parser entry point
 */

function yyparse()
{
  var yyastk = new Array();
  var yysstk = new Array();

  yystate = 0;
  yychar = -1;

  yysp = 0;
  yysstk[yysp] = 0;
  yyerrflag = 0;
  for (;;) {
    if (yybase[yystate] == 0)
      yyn = yydefault[yystate];
    else {
      if (yychar < 0) {
        if ((yychar = yylex()) <= 0) yychar = 0;
        yychar = yychar < YYMAXLEX ? yytranslate[yychar] : YYBADCH;
      }

      if (((yyn = yybase[yystate] + yychar) >= 0
	    && yyn < YYLAST && yycheck[yyn] == yychar
           || (yystate < YY2TBLSTATE
               && (yyn = yybase[yystate + YYNLSTATES] + yychar) >= 0
               && yyn < YYLAST && yycheck[yyn] == yychar))
	  && (yyn = yyaction[yyn]) != YYDEFAULT) {
        /*
         * >= YYNLSTATE: shift and reduce
         * > 0: shift
         * = 0: accept
         * < 0: reduce
         * = -YYUNEXPECTED: error
         */
        if (yyn > 0) {
          /* shift */
          yysp++;

          yysstk[yysp] = yystate = yyn;
          yyastk[yysp] = yylval;
          yychar = -1;
          
          if (yyerrflag > 0)
            yyerrflag--;
          if (yyn < YYNLSTATES)
            continue;
            
          /* yyn >= YYNLSTATES means shift-and-reduce */
          yyn -= YYNLSTATES;
        } else
          yyn = -yyn;
      } else
        yyn = yydefault[yystate];
    }
      
    for (;;) {
      /* reduce/error */
      if (yyn == 0) {
        /* accept */
        yyflush();
        return 0;
      }
      else if (yyn != YYUNEXPECTED) {
        /* reduce */
        yyl = yylen[yyn];
        yyval = yyastk[yysp-yyl+1];
        /* Following line will be replaced by reduce actions */
        switch(yyn) {
        case 1:
{ans = yyastk[yysp-(1-1)];} break;
        case 8:
{yyval = checkNumber(yyastk[yysp-(3-1)])+checkNumber(yyastk[yysp-(3-3)]);} break;
        case 9:
{yyval = checkNumber(yyastk[yysp-(3-1)])-checkNumber(yyastk[yysp-(3-3)]);} break;
        case 10:
{yyval = checkNumber(yyastk[yysp-(3-1)])*checkNumber(yyastk[yysp-(3-3)]);} break;
        case 11:
{yyval = checkNumber(yyastk[yysp-(3-1)])/checkNumber(yyastk[yysp-(3-3)]);} break;
        case 12:
{yyval = Math.pow(checkNumber(yyastk[yysp-(3-1)]), checkNumber(yyastk[yysp-(3-3)]));} break;
        case 13:
{yyval = checkString(yyastk[yysp-(3-1)])+checkString(yyastk[yysp-(3-3)]);} break;
        case 14:
{yyval = yyastk[yysp-(3-1)] < yyastk[yysp-(3-3)];} break;
        case 15:
{yyval = yyastk[yysp-(3-1)] <= yyastk[yysp-(3-3)];} break;
        case 16:
{yyval = yyastk[yysp-(3-1)] > yyastk[yysp-(3-3)];} break;
        case 17:
{yyval = yyastk[yysp-(3-1)] >= yyastk[yysp-(3-3)];} break;
        case 18:
{yyval = yyastk[yysp-(3-1)] === yyastk[yysp-(3-3)];} break;
        case 19:
{yyval = yyastk[yysp-(3-1)] !== yyastk[yysp-(3-3)];} break;
        }
        /* Goto - shift nonterminal */
        yysp -= yyl;
        yyn = yylhs[yyn];
        if ((yyp = yygbase[yyn] + yysstk[yysp]) >= 0 && yyp < YYGLAST
            && yygcheck[yyp] == yyn)
          yystate = yygoto[yyp];
        else
          yystate = yygdefault[yyn];
          
        yysp++;

        yysstk[yysp] = yystate;
        yyastk[yysp] = yyval;
      }
      else {
        /* error */
        switch (yyerrflag) {
        case 0:
          yyerror("syntax error");
        case 1:
        case 2:
          yyerrflag = 3;
          /* Pop until error-expecting state uncovered */

          while (!((yyn = yybase[yystate] + YYINTERRTOK) >= 0
                   && yyn < YYLAST && yycheck[yyn] == YYINTERRTOK
                   || (yystate < YY2TBLSTATE
                       && (yyn = yybase[yystate + YYNLSTATES] + YYINTERRTOK) >= 0
                       && yyn < YYLAST && yycheck[yyn] == YYINTERRTOK))) {
            if (yysp <= 0) {
              yyflush();
              return 1;
            }
            yystate = yysstk[--yysp];
          }
          yyn = yyaction[yyn];
          yysstk[++yysp] = yystate = yyn;
          break;

        case 3:
          if (yychar == 0) {
            yyflush();
            return 1;
          }
          yychar = -1;
          break;
        }
      }
        
      if (yystate < YYNLSTATES)
        break;
      /* >= YYNLSTATES means shift-and-reduce */
      yyn = yystate - YYNLSTATES;
    }
  }
}



  /* Lexical analyzer */

  function yylex()
  {
    buffer = buffer.trim();
    if (buffer.length == 0)
      return 0;

    //セル参照
    var m = buffer.match(/^\$?[A-Z]+\$?\d+/);
    if(m){
      var row_col = Utils.toRowCol(m[0]);
      yylval = worksheet.getCell(row_col[0], row_col[1]);
      buffer = buffer.substr(m[0].length);
      return CELL;
    }
    //数字
    m = buffer.match(/^\d+(?:\.\d+)?/);
    if(m){
      yylval = Number(m[0]);
      buffer = buffer.substr(m[0].length);
      return NUMBER;
    }
    //文字列
    m = buffer.match(/^"((?:""|[^"])*)"/);
    if(m){
      yylval = m[1].replace(/""/g, '"');
      buffer = buffer.substr(m[0].length);
      return STRING;
    }
    //IDENT
    m = buffer.match(/^[A-Z_]+/i);
    if(m){
      yylval = m[0].toUpperCase();
      buffer = buffer.substr(m[0].length);
      switch(yylval){
        case "SUM":
          return SUM;
        case "TRUE":
          yylval = true;
          return TRUE;
        case "FALSE":
          yylval = false;
          return FALSE;
        default:
          return IDENT;
      }
    }
    //binary operator
    m = buffer.match(/^(?:<>|<=|>=)/i);
    if(m){
      switch(m[0]){
        case "<>":
          return NEQ;
        case "<=":
          return NEQ;
        case ">=":
          return NEQ;
      }
      yylval = buffer.substr(0, 2);
      buffer = buffer.substr(2);
    }

    yylval = buffer.substr(0, 1);
    buffer = buffer.substr(1);
    return yylval.charCodeAt(0);
  }

  buffer = formula;
  yyparse();
  return ans;
};
