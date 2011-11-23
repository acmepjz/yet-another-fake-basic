Attribute VB_Name = "mdlTypes"
Option Explicit

Public Enum enumVBChar
 ["\0"] = 0
 ["\t"] = 9
 ["\n"] = 10
 ["\r"] = 13
 [" "] = 32
 ["!"] = 33
 ["""] = 34
 ["#"] = 35
 ["$"] = 36
 ["%"] = 37
 ["&"] = 38
 ["'"] = 39
 ["("] = 40
 [")"] = 41
 ["*"] = 42
 ["+"] = 43
 [","] = 44
 ["-"] = 45
 ["."] = 46
 ["/"] = 47
 ["0"] = 48
 ["1"] = 49
 ["2"] = 50
 ["3"] = 51
 ["4"] = 52
 ["5"] = 53
 ["6"] = 54
 ["7"] = 55
 ["8"] = 56
 ["9"] = 57
 [":"] = 58
 [";"] = 59
 ["<"] = 60
 ["="] = 61
 [">"] = 62
 ["?"] = 63
 ["@"] = 64
 ["a"] = 65
 ["b"] = 66
 ["c"] = 67
 ["d"] = 68
 ["e"] = 69
 ["f"] = 70
 ["g"] = 71
 ["h"] = 72
 ["i"] = 73
 ["j"] = 74
 ["k"] = 75
 ["l"] = 76
 ["m"] = 77
 ["n"] = 78
 ["o"] = 79
 ["p"] = 80
 ["q"] = 81
 ["r"] = 82
 ["s"] = 83
 ["t"] = 84
 ["u"] = 85
 ["v"] = 86
 ["w"] = 87
 ["x"] = 88
 ["y"] = 89
 ["z"] = 90
 ["lll"] = 91
 ["\"] = 92
 ["rrr"] = 93
 ["^"] = 94
 ["_"] = 95
 ["`"] = 96
 ["aa"] = 97
 ["bb"] = 98
 ["cc"] = 99
 ["dd"] = 100
 ["ee"] = 101
 ["ff"] = 102
 ["gg"] = 103
 ["hh"] = 104
 ["ii"] = 105
 ["jj"] = 106
 ["kk"] = 107
 ["ll"] = 108
 ["mm"] = 109
 ["nn"] = 110
 ["oo"] = 111
 ["pp"] = 112
 ["qq"] = 113
 ["rr"] = 114
 ["ss"] = 115
 ["tt"] = 116
 ["uu"] = 117
 ["vv"] = 118
 ["ww"] = 119
 ["xx"] = 120
 ["yy"] = 121
 ["zz"] = 122
 ["{"] = 123
 ["|"] = 124
 ["}"] = 125
 ["~"] = 126
End Enum

Public Enum enumTokenType
 token_eof = -1
 token_err = 0
 '///
 token_id = 1000
 token_decnum = 2
 token_hexnum = 3
 token_octnum = 4
 token_floatnum = 5
 token_string = 6
 token_crlf = 7
 token_colon = 8 '":"
 token_dot = 9 '"."
 token_comma = 10 '","
 token_semicolon = 11 '";"
 token_poundsign = 12 '"#"
 token_lbracket = 13 '"("
 token_rbracket = 14 '")"
 token_plus = 15 '+
 token_minus = 16 '-
 token_asterisk = 17 '*
 token_slash = 18 '/
 token_backslash = 19 '\
 token_equal = 20 '=
 token_power = 21 '^
 token_lt = 22 '<
 token_gt = 23 '>
 token_le = 24 '<=|=<
 token_ge = 25 '>=|=>
 token_ne = 26 '<>|><
 token_and = 27 '&
 token_currencynum = 28
 token_datenum = 29
 token_linenumber = 30
 token_shl = 31 '<<
 token_shr = 32 '>>
 token_rol = 33 '<<<
 token_ror = 34 '>>>
 '///in frm and ctl
 token_guid = 101
 '///preprocessors
 preprocessor_const = 901
 preprocessor_else = 902
 preprocessor_elseif = 903
 preprocessor_end = 904
 preprocessor_if = 905
 '### BEGIN KEYWORD ENUM
 keyword_alias = 1001
 keyword_and = 1002
 keyword_as = 1003
 keyword_attribute = 1004
 keyword_byref = 1005
 keyword_byval = 1006
 keyword_call = 1007
 keyword_case = 1008
 keyword_cdecl = 1009
 keyword_close = 1010
 keyword_const = 1011
 keyword_declare = 1012
 keyword_dim = 1013
 keyword_do = 1014
 keyword_each = 1015
 keyword_else = 1016
 keyword_elseif = 1017
 keyword_end = 1018
 keyword_enum = 1019
 keyword_eqv = 1020
 keyword_erase = 1021
 keyword_exit = 1022
 keyword_false = 1023
 keyword_fastcall = 1024
 keyword_for = 1025
 keyword_friend = 1026
 keyword_function = 1027
 keyword_get = 1028
 keyword_global = 1029
 keyword_goto = 1030
 keyword_if = 1031
 keyword_imp = 1032
 keyword_in = 1033
 keyword_input = 1034
 keyword_is = 1035
 keyword_let = 1036
 keyword_lib = 1037
 keyword_line = 1038
 keyword_loop = 1039
 keyword_lset = 1040
 keyword_mod = 1041
 keyword_new = 1042
 keyword_next = 1043
 keyword_not = 1044
 keyword_on = 1045
 keyword_open = 1046
 keyword_option = 1047
 keyword_optional = 1048
 keyword_or = 1049
 keyword_paramarray = 1050
 keyword_preserve = 1051
 keyword_print = 1052
 keyword_private = 1053
 keyword_property = 1054
 keyword_public = 1055
 keyword_put = 1056
 keyword_raiseevent = 1057
 keyword_redim = 1058
 keyword_rset = 1059
 keyword_select = 1060
 keyword_set = 1061
 keyword_static = 1062
 keyword_stdcall = 1063
 keyword_step = 1064
 keyword_sub = 1065
 keyword_then = 1066
 keyword_to = 1067
 keyword_true = 1068
 keyword_type = 1069
 keyword_until = 1070
 keyword_wend = 1071
 keyword_while = 1072
 keyword_with = 1073
 keyword_withevents = 1074
 keyword_write = 1075
 keyword_xor = 1076
 '### END KEYWORD ENUM
End Enum

Public Type typeToken
 nType As enumTokenType
 nLine As Long
 nColumn As Long
 sValue As String
 '///some stupid attrubites
 nFlags As Integer
 '1=some spaces before it
 'etc.
 nFlags2 As Integer
 '1=begin with "."
 'etc.
 nReserved2 As Long
 nReserved3 As Long
 nReserved4 As Long
End Type

Public Enum enumASTNodeType
 node_id = 1 '<id>:{id}
 node_const = 2 '<const>:{intnum}|{hexnum}|{octnum}|{floatnum}|{strconst}  'etc...
 node_var = 3 '<var>:(<array_or_func>|<membervar>)<membervar>*
' node_membervar = 4 '<membervar>:{point}<array_or_func>
 node_array_or_func = 5 '<array_or_func>:<id>({(}<arglist>{)})+
 node_arglist = 6 '<arglist>:(({byval}?<exp>)?{,})*{byval}?<exp> 'ByVal??? TODO:
 '///expression process
' node_term = 7 '<term>:<var>|<const>|{(}<exp>{)}
' node_pwr_term = 8 '<pwr_term>:<term>|<pwr_term>{^}<term>
' node_neg_term = 9 '<neg_term>:<pwr_term>|{-}<neg_term>
' node_mul_term = 10 '<mul_term>:<neg_term>|<mul_term>{*}<neg_term>|<mul_term>{/}<neg_term>
' node_div_term = 11 '<div_term>:<mul_term>|<div_term>{\}<mul_term>
' node_mod_term = 12 '<mod_term>:<div_term>|<mod_term>{mod}<div_term>
' node_add_term = 13 '<add_term>:<mod_term>|<add_term>{+}<mod_term>|<add_term>{-}<mod_term>
' node_sadd_term = 14 '<sadd_term>:<add_term>|<sadd_term>{&}<add_term>
' node_rel_term = 15 '<rel_term>:<sadd_term>|<rel_term>{>}<sadd_term>|<rel_term>{<}<sadd_term>|
' '<rel_term>{>=}<sadd_term>|<rel_term>{<=}<sadd_term>|<rel_term>{=}<sadd_term>|<rel_term>{<>}<sadd_term>|<rel_term>{is}<sadd_term>
' node_not_term = 16 '<not_term>:<rel_term>|{not}<not_term>
' node_and_term = 17 '<and_term>:<not_term>|<and_term>{and}<not_term>
' node_or_term = 18 '<or_term>:<and_term>|<or_term>{or}<and_term>
' node_xor_term = 19 '<xor_term>:<or_term>|<xor_term>{xor}<or_term>
 node_exp = 20 '<exp>:<xor_term>
 '///statments
 node_makestat = 100 '<makestat>:({Let}|{Set}|{LSet}|{RSet})?<var>{=}<exp>
 node_callstat = 101 '<callstat>:{call}<var>|<var><arglist>   'and the dirty workaround <var>{,}<arglist> but it will be buggy
 node_ifstat = 102 '<ifstat>:{if}<exp>{then}<ifstatlist>@1@|{if}<exp>{then}<ifstatlist>{else}<ifstatlist>@1@|{if}<exp>{then}{else}<ifstatlist>@1@
 node_ifblock = 103 '<ifblock>:{if}<exp>{then}<br><statlist><elseifblock>*<elseblock>?{end}{if}
 node_forstat = 104 '{for}(<var>{=}<exp>{to}<exp>|{each}<var>{in}<exp>)<br><statlist>{next}<var> 'TODO:Next k,j,i = ?
 node_whilestat = 105 '{while}<exp><br><statlist>{wend}
 node_dostat = 106 '{do}(({while}|{until})<exp>)?<br><statlist>{loop}(({while}|{until})<exp>)?
 node_selectstat = 107 '{select}{case}<exp><br><selectblock>*{end}{select}
 node_exitstat = 108 '{exit}({sub}|{function}|{property}|{for}|{do}...)
 node_withstat = 109 '{with}<exp><br><statlist>{end}{with}
 node_linenumberstat = 110 '{intnum}|<id>{:} 'TODO:<br>???
 node_errorstat = 111 '{on}{local}?{error}({goto}{intnum}|{goto}<id>|{resume}{next})
 node_gotostat = 112 '{goto}{intnum}|{goto}<id>
 node_dimstat = 113 '({dim}|{static}|{private}|{public}|{global})(<dimitem>{,})*<dimitem>
 node_redimstat = 114 '{redim}{preserve}?(<redimitem>{,})*<redimitem>
 node_openstat = 115 '{open}<exp>{for}({input}|{output}|{append}|{random}|{binary})({access}{read}{write}?|{access}{write})?
 '({shared}|{lock}{read}{write}?|{lock}{write})?{as}{#}?<exp>
 node_closestat = 116 '{close}(({#}?<exp>{,})*{#}?<exp>)?
 node_lineinputstat = 117 '{line}{input}{#}?<exp>{,}<var> 'TODO:
 node_inputstat = 118 '{input}{#}?<exp>({,}<var>)+  'TODO:
 node_printstat = 119 '{print}{#}?<exp>{,}<printlist> 'TODO:
 node_writestat = 120 '{write}{#}?<exp>{,}<printlist> 'TODO:
 node_getstat = 121 '{get}{#}?<exp>{,}<exp>?{,}<var>
 node_putstat = 122 '{put}{#}?<exp>{,}<exp>?{,}<exp>
 node_namestat = 123 '{name}<exp>{as}<exp>
 node_raiseeventstat = 124 '{raiseevent}<array_or_func>
 node_debugassert = 125 '{debug}{point}{assert}<exp>
 node_debugprint = 126 '{debug}{point}{print}<printlist>
 node_erasestat = 127 '{erase}<var>({,}<var>)* '?? TODO:
 '///
 '... and lock #1,xx to xx unlock #1,xx to xx (unsupported)
 'hidden
 node_attributestat = 190 '{Attribute}<id>{=}<exp>
 '///declare statments - include node_dimstat
 node_funcstat = 200 '({private}|{public}|{friend})?{static}?({sub}|{function}|{property}({get}|{let}|{set}))<id>{(}<argumentlist>{)}<dimtype>?
 '<br><statlist>{end}({sub}|{function}|{property})
' node_apistat = 201 '({private}|{public})?{declare}({sub}|{function})<id>{lib}{strconst}({alias}{strconst})?{(}<argumentlist>{)}<dimtype>?
 node_enumstat = 202 '({private}|{public})?{enum}<id><br>(<id>{=}<exp><br>)+{end}{enum}
 node_typestat = 203 '({private}|{public})?{type}<id><br>(<dimitem><br>)+{end}{type}
 node_optionstat = 204 '{option}({explicit}|{base}{0}|{base}{1}|{compare}{binary}|{compare}{text})
 node_eventstat = 205 '{public}?{event}<id>{(}<argumentlist>{)}
 '///not really statments
 node_elseifblock = 900 '<elseifblock>:({elseif}|{else}{if})<exp>{then}<br><statlist>
 node_elseblock = 901 '<elseblock>:|{else}<br><statlist>
 node_selectblock = 902 '{case}<selectconditionlist><br><statlist>
 node_selectconditionlist = 903 '{else}|(<selectcondition>{,})*<selectcondition>
 node_selectcondition = 904 '<exp>|<exp>{to}<exp>|{is}({=}|{<>}|{>}|{<}|{>=}|{<=})<exp>
 node_dimitem = 905 '{withevents}?(<id>|<id>{(}{)}|<id>{(}<arraydim>{)})<dimtype>?
 node_arraydim = 906 '(<arraydimitem>{,})*<arraydimitem>
 node_arraydimitem = 907 '<exp>|<exp>{to}<exp>
 node_dimtype = 908 '{as}{new}?(<id>{point})*<id>({*}<exp>)?
 node_redimitem = 909 '(<id>|@#^$...){(}<arraydim>{)}<dimtype>? '??? TODO:
 node_argumentlist = 910 '(<argumentitem>{,})*<argumentitem>? 'TODO:optional
 node_argumentitem = 911 '{optional}?{paramarray}?({byval}|{byref})?<id>({(}{)})?<dimtype>?({=}<exp>)?
 node_printlist = 912 '<printitem>*<exp>?
 node_printitem = 913 '<exp>?({,}|{;})
 '///
 node_statlist = 1000 'end with <br> ,can be empty(<br> only)
 node_ifstatlist = 1001 '{:} only, end without <br> ,can't be empty
 '///<declstatlist>
 node_module_root = 10000
 node_class_root = 10001
 node_form_root = 10002
 node_control_root = 10003
End Enum

Public Enum enumASTNodeVerifyStep
 verify_const = 1
 verify_type = 3
 verify_dim = 4
 verify_all = 99
End Enum

Public Enum enumASTNodeProperty
 prop_endblockhandle = 1
 '/////
 action_const_codegen = 10001
End Enum
