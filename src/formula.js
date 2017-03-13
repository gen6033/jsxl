
Utils = require("./utils");
Range = require("./range").Range;
FormulaError = require("./formula/error")
FormulaEvaluator = require("./formula/evaluator")

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

  if(formula === undefined){
    return;
  }
  var ans;
  var evaluator = new FormulaEvaluator(worksheet);

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
var FUNC = 260;
var TRUE = 261;
var FALSE = 262;
var IGNORE = 263;
var LE = 264;
var GE = 265;
var NEQ = 266;
var SIGN = 267;
var SPACES = 268;

  
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
      0,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   20,   15,   29,
     24,   25,   18,   16,    8,   17,   29,   19,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   22,   28,
      9,   13,   11,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   21,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   26,   29,   27,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,   29,   29,   29,   29,
     29,   29,   29,   29,   29,   29,    1,    2,    3,    4,
      5,    6,    7,   29,   10,   12,   14,   29,   23
  ];

var YYBADCH = 29;
var YYMAXLEX = 269;
var YYTERMS = 29;
var YYNONTERMS = 13;

var yyaction = [
     78,   61,   60,   42,   59,   58,   86,   12,   13,   14,
     15,   16,   17,   18,   19,   13,   14,   15,   16,   17,
     18,   19,   40,    0,   33,   20,   21,   22,   23,   57,
     24,   25,   27,-32767,-32767,-32767,-32767,-32767,-32767,   19,
     38,-32766,-32766,-32766,-32766,-32766,-32766,   10,   11,   25,
     27,    0,    0,   81,   39,   75,    0,   52,    0,    0,
     82
  ];

var YYLAST = 61;

var yycheck = [
      2,    3,    4,    5,    6,    7,   23,    8,    9,   10,
     11,   12,   13,   14,   15,    9,   10,   11,   12,   13,
     14,   15,   24,    0,   26,   16,   17,   18,   19,   20,
     21,   22,   23,    9,   10,   11,   12,   13,   14,   15,
     24,   16,   17,   18,   19,   16,   17,   16,   17,   22,
     23,   -1,   -1,   27,   28,   25,   -1,   25,   -1,   -1,
     27
  ];

var yybase = [
    -17,   -1,    6,   24,   24,   24,   24,   24,   24,   31,
     31,   31,   31,   31,   31,   31,   31,   31,   31,   31,
     31,   31,   31,   31,   31,   31,   31,   -2,    9,   29,
     29,   25,   25,   33,   27,   27,   26,   27,  -17,  -17,
    -17,   23,   16,   32,   30,    0,    9,    9,    9,    9,
      9,    9,    9,    9,   -2,   -2,   -2,   -2,   -2,   -2,
     -2,   -2,   -2,   -2,   -2,   -2,   -2,   -2,   -2,   -2,
     -2,   -2,    0,    0,    9,    9,    9,    9,  -17
  ];

var YY2TBLSTATE = 34;

var yydefault = [
     42,    2,   29,   23,   24,   25,   26,   27,   28,32767,
  32767,32767,32767,32767,32767,32767,32767,32767,32767,32767,
  32767,32767,32767,32767,32767,32767,   32,   11,   22,   17,
     18,   19,   20,   42,    8,    9,32767,   21,   42,   42,
     42,32767,32767,32767,32767
  ];



var yygoto = [
     34,   35,    2,    3,    4,    5,    6,    7,    8,   28,
     29,   30,   31,   32,   37,   79,   46,   80,   26,    0,
     83,    0,    0,    0,    0,    0,    0,    0,    0,    0,
      0,    0,    0,    0,    0,    0,    0,    0,    0,    0,
      0,    0,    0,    0,    0,    0,    0,    0,    0,    0,
      0,    0,    0,    0,   76,    0,   43
  ];

var YYGLAST = 57;

var yygcheck = [
      4,    4,    4,    4,    4,    4,    4,    4,    4,    4,
      4,    4,    4,    4,    4,    4,    2,    4,    3,   -1,
     12,   -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,
     -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,
     -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,   -1,
     -1,   -1,   -1,   -1,    2,   -1,    2
  ];

var yygbase = [
      0,    0,   16,  -20,  -10,    0,    0,    0,    0,    0,
      0,    0,  -19
  ];

var yygdefault = [
  -32768,   41,   85,    9,    1,   48,   49,   50,   51,   55,
     44,   36,   84
  ];

var yylhs = [
      0,    1,    2,    4,    4,    4,    4,    4,    4,    4,
      4,    4,    4,    5,    5,    5,    5,    8,    8,    8,
      8,    8,    8,    8,    8,    8,    8,    8,    8,    8,
      7,   10,   10,    6,    6,    6,    9,    9,   11,   11,
     12,    3,    3
  ];

var yylen = [
      1,    1,    2,    1,    1,    1,    1,    3,    2,    2,
      1,    2,    2,    1,    1,    1,    1,    3,    3,    3,
      3,    3,    3,    3,    3,    3,    3,    3,    3,    3,
      4,    1,    1,    1,    3,    3,    3,    2,    3,    1,
      1,    1,    0
  ];

var YYSTATES = 68;
var YYNLSTATES = 45;
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
{ans = evaluator.getValue(yyastk[yysp-(1-1)]);} break;
        case 2:
{yyval=yyastk[yysp-(2-2)];} break;
        case 7:
{
      var expr = yyastk[yysp-(3-2)];
      if(Array.isArray(expr)){
        var r = expr[0];
        for(var i = 1; i < expr.length; i++){
          r.union(expr[i]);
        }
        yyval = r;
      }else{
        yyval = expr;
      }
    } break;
        case 8:
{yyval=evaluator.expectNumber(yyastk[yysp-(2-2)]);} break;
        case 9:
{yyval=-evaluator.expectNumber(yyastk[yysp-(2-2)]);} break;
        case 12:
{
      yyval = evaluator.expectNumber(yyastk[yysp-(2-1)]) / 100
    } break;
        case 17:
{
      var a = yyastk[yysp-(3-1)];
      var b = yyastk[yysp-(3-3)];
      yyval = evaluator.expectNumber(a) + evaluator.expectNumber(b);
      if(Utils.isDate(a) || Utils.isDate(b)){
        yyval = Utils.offsetToDate(yyval);
      }
    } break;
        case 18:
{
      var a = yyastk[yysp-(3-1)];
      var b = yyastk[yysp-(3-3)];
      yyval = evaluator.expectNumber(a) - evaluator.expectNumber(b);
      if(Utils.isDate(a) || Utils.isDate(b)){
        yyval = Utils.offsetToDate(yyval);
      }
    } break;
        case 19:
{
      var a = yyastk[yysp-(3-1)];
      var b = yyastk[yysp-(3-3)];
      yyval = evaluator.expectNumber(a) * evaluator.expectNumber(b);
      if(Utils.isDate(a) || Utils.isDate(b)){
        yyval = Utils.offsetToDate(yyval);
      }
    } break;
        case 20:
{
      var a = yyastk[yysp-(3-1)];
      var b = yyastk[yysp-(3-3)];
      var num_b = evaluator.expectNumber(b);
      if(num_b == 0){
        yyval = FormulaError.DIV0;
      }else{
        yyval = evaluator.expectNumber(a) / num_b;
        if(Utils.isDate(a) || Utils.isDate(b)){
          yyval = Utils.offsetToDate(yyval);
        }
      }
    } break;
        case 21:
{
      var a = yyastk[yysp-(3-1)];
      var b = yyastk[yysp-(3-3)];
      yyval = Math.pow(evaluator.expectNumber(yyastk[yysp-(3-1)]), evaluator.expectNumber(yyastk[yysp-(3-3)]));
      if(Utils.isDate(a) || Utils.isDate(b)){
        yyval = Utils.offsetToDate(yyval);
      }
    } break;
        case 22:
{yyval = evaluator.expectString(yyastk[yysp-(3-1)])+evaluator.expectString(yyastk[yysp-(3-3)]);} break;
        case 23:
{yyval = yyastk[yysp-(3-1)] < yyastk[yysp-(3-3)];} break;
        case 24:
{yyval = yyastk[yysp-(3-1)] <= yyastk[yysp-(3-3)];} break;
        case 25:
{yyval = yyastk[yysp-(3-1)] > yyastk[yysp-(3-3)];} break;
        case 26:
{yyval = yyastk[yysp-(3-1)] >= yyastk[yysp-(3-3)];} break;
        case 27:
{
      var a = yyastk[yysp-(3-1)];
      var b = yyastk[yysp-(3-3)];
      if(Utils.isString(a) && Utils.isString(b)){
        yyval = (String(a).toLowerCase() == String(b).toLowerCase())
      }else{
        yyval = yyastk[yysp-(3-1)] === yyastk[yysp-(3-3)];
      }
    } break;
        case 28:
{yyval = yyastk[yysp-(3-1)] !== yyastk[yysp-(3-3)];} break;
        case 29:
{
      var list = [].concat(yyastk[yysp-(3-1)])
      list.push(yyastk[yysp-(3-3)]);
      yyval = list;
    } break;
        case 30:
{
      var func = evaluator[yyastk[yysp-(4-1)]];
      if(!func){
        return FormulaError.NAME
      }
      try{
        yyval = func.call(evaluator, [].concat(yyastk[yysp-(4-3)]));
      }catch (e){
        if(e instanceof FormulaError){
          yyval = e
        }else{
          throw e
        }
      }
    } break;
        case 32:
{yyval = []} break;
        case 34:
{yyval = yyastk[yysp-(3-1)].union(yyastk[yysp-(3-3)]).unify();} break;
        case 35:
{yyval = yyastk[yysp-(3-1)].intersection(yyastk[yysp-(3-3)]);} break;
        case 36:
{yyval = yyastk[yysp-(3-2)];} break;
        case 37:
{yyval = [];} break;
        case 38:
{yyastk[yysp-(3-1)].push(yyastk[yysp-(3-3)]);yyval = yyastk[yysp-(3-1)];} break;
        case 39:
{yyval = [yyastk[yysp-(1-1)]];} break;
        case 40:
{
    yyval = [].concat(yyastk[yysp-(1-1)]);
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

    //TRUE
    m = buffer.match(/^TRUE(?!\()/i);
    if(m){
      yylval = true
      buffer = buffer.substr(m[0].length);
      return TRUE;
    }

    //FALSE
    m = buffer.match(/^FALSE(?!\()/i);
    if(m){
      yylval = false
      buffer = buffer.substr(m[0].length);
      return FALSE;
    }

    //FUNC
    m = buffer.match(/^(?:_xlfn\.)?([A-Z_.]+[A-Z0-9_.]*)(?=\()/i);
    if(m){
      yylval = m[1].toUpperCase();
      buffer = buffer.substr(m[0].length);
      return FUNC;
    }

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
    //binary operator
    m = buffer.match(/^\s*(=|<>|<=|>=|<|>|\+|-|\*|\/|\^|&|,|:|;)\s*/i);
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
