Attribute VB_Name = "RegularExpressionUtility"
'<License>------------------------------------------------------------
'
' Copyright (c) 2019 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
' 指定文字列を正規表現検索して最初に見つかった文字位置(1based)を返す
' 見つからなかった場合は #VALUE! を返す
'
' ## Parameters
'
'  - keyWord
'      検索キーワード
'  - fromThisString
'      検索対象範囲
'  - ignoreCase default True
'      大文字/小文字を区別する
'
' ## Returns
'
'  最初に見つかった文字列の文字位置(1based。Excel数式 `FIND` が返す文字位置に同じ。)
'  文字列が見つからなかった場合は `#VALUE!` を返す( Excel数式 `FIND` の 見つからなかった時の返却値に同じ)
'
Public Function regexFirstIndexOfMatched( _
    ByVal keyword As String, _
    ByVal fromThisString As String, _
    Optional ByVal ignoreCase As Boolean = True _
    ) As Variant
    
    Dim ret As Variant
    
    Set collection_matched = func_regex_execute(keyword, fromThisString, ignoreCase)
    
    If collection_matched.Count() = 0 Then ' 1つも match しなかった時
        
        '#VALUE!を返す( Excel数式 `FIND` の 見つからなかった時の返却値に同じ)
        ret = CVErr(xlErrValue)
    
    Else ' 1つ以上 match した時
    
        '1文字目を 1 と考えた場合の 文字位置を返す( Excel数式 `FIND` が返す文字位置に同じ)
        ret = (collection_matched.Item(0).FirstIndex + 1)
    
    End If
    
    '返却値を格納して終了
    regexFirstIndexOfMatched = ret
    
End Function

'
' 指定文字列を正規表現検索して最初に見つかった文字列の長さを返す
' 見つからなかった場合は #VALUE! を返す
'
' ## Parameters
'
'  - keyWord
'      検索キーワード
'  - fromThisString
'      検索対象範囲
'  - ignoreCase default True
'      大文字/小文字を区別する
'
' ## Returns
'
'  最初に見つかった文字列の文字列長
'
Public Function regexLengthOfMatched( _
    ByVal keyword As String, _
    ByVal fromThisString As String, _
    Optional ByVal ignoreCase As Boolean = True _
    ) As Variant
    
    Dim ret As Variant
    
    Set collection_matched = func_regex_execute(keyword, fromThisString, ignoreCase)
    
    If collection_matched.Count() = 0 Then ' 1つも match しなかった時
        
        '#VALUE!を返す( Excel数式 `FIND` の 見つからなかった時の返却値に同じ)
        ret = CVErr(xlErrValue)
    
    Else ' 1つ以上 match した時
    
        'マッチした文字列の長さを返す
        ret = (collection_matched.Item(0).Length)
    
    End If
    
    '返却値を格納して終了
    regexLengthOfMatched = ret
    
End Function

'
' 指定文字列を正規表現検索して置換する
'
' ## Parameters
'
'  - fromThisString
'      検索対象範囲
'  - keyWord
'      検索キーワード
'  - replacer
'      置換キーワード
'  - ignoreCase default True
'      大文字/小文字を区別する
'
' ## Returns
'
'  置換後の文字列。
'  置換対象文字列が見つからなかった場合は、置換しないまま返す。
'
Public Function regexSubstitute( _
    ByVal fromThisString As String, _
    ByVal keyword As String, _
    ByVal replacer As String, _
    Optional ByVal ignoreCase As Boolean = True _
    ) As Variant
    
    Dim ret As Variant
    
    ret = func_regex_replace(keyword, replacer, fromThisString, ignoreCase)
    
    regexSubstitute = ret
    
End Function

'<Common>---------------------------------------------------------------------------

'
' 指定文字列を正規表現検索して MatchCollection を返す
'
Private Function func_regex_execute( _
    ByVal str_keyword As String, _
    ByVal str_fromThisString As String, _
    Optional ByVal bool_ignoreCase As Boolean = True, _
    Optional ByVal bool_global As Boolean = True)
    
    Set obj_regex = CreateObject("VBScript.RegExp")

    obj_regex.Pattern = str_keyword
    obj_regex.ignoreCase = bool_ignoreCase
    obj_regex.Global = bool_global

    Set collection_matched = obj_regex.Execute(str_fromThisString)
    
    Set func_regex_execute = collection_matched
    
End Function

'
' 指定文字列を正規表現検索して置換する
'
Private Function func_regex_replace( _
    ByVal str_keyword As String, _
    ByVal str_replacer As String, _
    ByVal str_fromThisString As String, _
    Optional ByVal bool_ignoreCase As Boolean = True, _
    Optional ByVal bool_global As Boolean = True)
    
    Set obj_regex = CreateObject("VBScript.RegExp")

    obj_regex.Pattern = str_keyword
    obj_regex.ignoreCase = bool_ignoreCase
    obj_regex.Global = bool_global

    func_regex_replace = obj_regex.Replace(str_fromThisString, str_replacer)
    
End Function

'--------------------------------------------------------------------------</Common>

