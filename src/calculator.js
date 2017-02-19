
Utils = require ("./utils");
Range = require ("./range").Range;

function yyerror(msg) {
  console.log(msg);
}


module.exports = function(worksheet, formula){

  //グローバル汚染対策
  "use strict";
  var yystate;
  var yychar;
  var yysp;
  var yyerrflag;
  var yyn;
  var yyl;
  var yyval;
  var yyp;

  var ans;
  if(formula === undefined){
    return;
  }

  function getValue(range){
    if(range.size == 1){
      var cell;
      range.each(function(r, c){
        cell = worksheet.getCell(r, c);
      });
      if(cell.formula){
        return cell.calculate();
      }else{
        return cell.value;
      }
    }
    throw new Error("#!VALUE");
  }
  function expectNumber(x){
    if(x instanceof Range){
      x = Number(getValue(x));
    }
    if(!Utils.isNumber(x)){
      throw new Error("NaN");
    }
    return x;
  }
  function expectString(x){
    if(x instanceof Range){
      x = getValue(x);
    }
    if(!Utils.isString(x)){
      throw new Error("NaN");
    }
    return x;
  }
  function expectRange(x){
    if(!(x instanceof Range)){
      throw new Error("#!VALUE");
    }
    return x;
  }
  function expectList(x){
    if(!Array.isArray(x)){
      throw new Error("#!VALUE");
    }
    return x;
  }

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
var SUM = 263;
var IGNORE = 264;
var LE = 265;
var GE = 266;
var NEQ = 267;
var SPACES = 268;
var SIGN = 269;

  
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
      0,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   19,   25,
     23,   24,   17,   15,    8,   16,   25,   18,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   22,   25,
      9,   13,   11,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   20,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,   25,   25,   25,   25,
     25,   25,   25,   25,   25,   25,    1,    2,    3,    4,
     25,    5,    6,    7,   25,   10,   12,   14,   21,   25
  ];

var YYBADCH = 25;
var YYMAXLEX = 270;
var YYTERMS = 25;
var YYNONTERMS = 10;

var yyaction = [
     13,   14,   15,   16,   17,   18,   19,   20,   21,   22,
     23,   24,   27,   26,   69,   54,   53,   52,   51,   37,
  -32767,-32767,-32767,-32767,-32767,-32767,   12,   10,   11,    0,
     21,   22,   23,   24,   24,   34,   27,   26,   73,   25,
  -32766,    0,    0,   35,    0,   48,   68
  ];

var YYLAST = 47;

var yycheck = [
      9,   10,   11,   12,   13,   14,   15,   16,   17,   18,
     19,   20,   21,   22,    2,    3,    4,    5,    6,    7,
      9,   10,   11,   12,   13,   14,    8,   15,   16,    0,
     17,   18,   19,   20,   20,   23,   21,   22,   21,   21,
     21,   -1,   -1,   23,   -1,   24,   24
  ];

var yybase = [
     17,   18,   -9,   11,   11,   11,   11,   11,   11,   12,
     12,   12,   12,   12,   12,   12,   12,   12,   12,   12,
     12,   12,   12,   12,   12,   12,   12,   12,   13,   13,
     14,   14,   14,   15,   17,   17,   29,   20,   21,   22,
     19,    0,   -9,    0,   -9,   -9,   -9,   -9,   -9,   -9,
      0,    0,    0,    0,    0,    0,    0,    0,    0,    0,
      0,    0,    0,    0,    0,    0,    0,    0,    0,   15,
     15,   15,   15,   15,    0,    0,    0,    0,    0,    0,
      0,   15
  ];

var YY2TBLSTATE = 41;

var yydefault = [
     33,   33,   31,   20,   21,   22,   23,   24,   25,32767,
  32767,32767,32767,32767,32767,32767,32767,32767,32767,32767,
  32767,32767,32767,32767,32767,   32,32767,32767,   14,   15,
     16,   17,   19,   18,   33,   33,32767,32767,32767,32767,
     30
  ];



var yygoto = [
      1,   49,   50,    2,    3,    4,    5,    6,    7,    8,
     28,   29,   30,   31,   32,   33,   43,   70,   38,   39
  ];

var YYGLAST = 20;

var yygcheck = [
      4,    4,    4,    4,    4,    4,    4,    4,    4,    4,
      4,    4,    4,    4,    4,    4,    3,    4,    2,    2
  ];

var yygbase = [
      0,    0,  -16,   15,   -9,    0,    0,    0,    0,    0
  ];

var yygdefault = [
  -32768,   36,   42,    9,   40,   44,   45,   46,   47,   67
  ];

var yylhs = [
      0,    1,    2,    4,    4,    4,    4,    4,    4,    4,
      5,    5,    5,    5,    8,    8,    8,    8,    8,    8,
      8,    8,    8,    8,    8,    8,    7,    9,    6,    6,
      6,    6,    3,    3
  ];

var yylen = [
      1,    1,    3,    1,    1,    1,    1,    3,    2,    2,
      1,    1,    1,    1,    3,    3,    3,    3,    3,    3,
      3,    3,    3,    3,    3,    3,    1,    4,    1,    3,
      3,    3,    1,    0
  ];

var YYSTATES = 59;
var YYNLSTATES = 41;
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
        case 2:
{yyval=yyastk[yysp-(3-2)];} break;
        case 7:
{yyval=yyastk[yysp-(3-2)];} break;
        case 8:
{yyval=expectNumber(yyastk[yysp-(2-2)]);} break;
        case 9:
{yyval=-expectNumber(yyastk[yysp-(2-2)]);} break;
        case 14:
{yyval = expectNumber(yyastk[yysp-(3-1)])+expectNumber(yyastk[yysp-(3-3)]);} break;
        case 15:
{yyval = expectNumber(yyastk[yysp-(3-1)])-expectNumber(yyastk[yysp-(3-3)]);} break;
        case 16:
{yyval = expectNumber(yyastk[yysp-(3-1)])*expectNumber(yyastk[yysp-(3-3)]);} break;
        case 17:
{yyval = expectNumber(yyastk[yysp-(3-1)])/expectNumber(yyastk[yysp-(3-3)]);} break;
        case 18:
{yyval = Math.pow(expectNumber(yyastk[yysp-(3-1)]), expectNumber(yyastk[yysp-(3-3)]));} break;
        case 19:
{yyval = expectString(yyastk[yysp-(3-1)])+expectString(yyastk[yysp-(3-3)]);} break;
        case 20:
{yyval = yyastk[yysp-(3-1)] < yyastk[yysp-(3-3)];} break;
        case 21:
{yyval = yyastk[yysp-(3-1)] <= yyastk[yysp-(3-3)];} break;
        case 22:
{yyval = yyastk[yysp-(3-1)] > yyastk[yysp-(3-3)];} break;
        case 23:
{yyval = yyastk[yysp-(3-1)] >= yyastk[yysp-(3-3)];} break;
        case 24:
{yyval = yyastk[yysp-(3-1)] === yyastk[yysp-(3-3)];} break;
        case 25:
{yyval = yyastk[yysp-(3-1)] !== yyastk[yysp-(3-3)];} break;
        case 27:
{
    (function(){
      var sum = 0;
      var list = [].concat(yyastk[yysp-(4-3)]);
      for(var i = 0; i < list.length; i++){
        var expr = list[i];
        if(expr instanceof Range){
          expr.each(function(r, c){
            var cell = worksheet.getCell(r, c);
            var val = cell.value;
            if(cell.formula){
              val = expectNumber(cell.calculate());
            }
            if(Utils.isNumber(val)){
              sum += Number(val);
            }
          })
        }else{
          sum += expectNumber(expr);
        }
      }
      yyval = sum;
    })()
    } break;
        case 29:
{yyval = yyastk[yysp-(3-1)].union(yyastk[yysp-(3-3)]).unify();} break;
        case 30:
{yyval = yyastk[yysp-(3-1)].intersection(yyastk[yysp-(3-3)]);} break;
        case 31:
{
      if(yyastk[yysp-(3-1)] instanceof Range && yyastk[yysp-(3-3)] instanceof Range){
        yyval = yyastk[yysp-(3-1)].union(yyastk[yysp-(3-3)]);
      }else{
        var list = [].concat(yyastk[yysp-(3-1)])
        list.push(yyastk[yysp-(3-3)]);
        yyval = list;
      }
    } break;
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
    if (buffer.length == 0)
      return 0;

    var m;

    //セル参照
    m = buffer.match(/^\$?[A-Z]+\$?\d+/);
    if(m){
      yylval = new Range(m[0]);
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
        case "TRUE":
          yylval = true;
          return TRUE;
        case "FALSE":
          yylval = false;
          return FALSE;
        case "SUM":
          return SUM;
        default:
          return IDENT;
      }
    }
    //binary operator
    m = buffer.match(/^\s*(=|<>|<=|>=|<|>|\+|-|\*|\/|\^|&|,|:)\s*/i);
    if(m){
      yylval = m[1];
      buffer = buffer.substr(m[0].length);
      switch(m[1]){
        case "<>":
          return NEQ;
        case "<=":
          return NEQ;
        case ">=":
          return NEQ;
        default:
          return yylval.charCodeAt(0);
      }
    }

    //スペースの連続
    m = buffer.match(/^\s+/);
    if(m){
      yylval = m[0];
      buffer = buffer.substr(m[0].length);
      return SPACES;
    }

    yylval = buffer.substr(0, 1);
    buffer = buffer.substr(1);
    return yylval.charCodeAt(0);
  }

  buffer = formula;
  yyparse();
  return ans;
};
