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
' �w�蕶����𐳋K�\���������čŏ��Ɍ������������ʒu(1based)��Ԃ�
' ������Ȃ������ꍇ�� #VALUE! ��Ԃ�
'
' ## Parameters
'
'  - keyWord
'      �����L�[���[�h
'  - fromThisString
'      �����Ώ۔͈�
'  - ignoreCase default True
'      �啶��/����������ʂ���
'
' ## Returns
'
'  �ŏ��Ɍ�������������̕����ʒu(1based�BExcel���� `FIND` ���Ԃ������ʒu�ɓ����B)
'  �����񂪌�����Ȃ������ꍇ�� `#VALUE!` ��Ԃ�( Excel���� `FIND` �� ������Ȃ��������̕ԋp�l�ɓ���)
'
Public Function regexFirstIndexOfMatched( _
    ByVal keyword As String, _
    ByVal fromThisString As String, _
    Optional ByVal ignoreCase As Boolean = True _
    ) As Variant
    
    Dim ret As Variant
    
    Set collection_matched = func_regex_execute(keyword, fromThisString, ignoreCase)
    
    If collection_matched.Count() = 0 Then ' 1�� match ���Ȃ�������
        
        '#VALUE!��Ԃ�( Excel���� `FIND` �� ������Ȃ��������̕ԋp�l�ɓ���)
        ret = CVErr(xlErrValue)
    
    Else ' 1�ȏ� match ������
    
        '1�����ڂ� 1 �ƍl�����ꍇ�� �����ʒu��Ԃ�( Excel���� `FIND` ���Ԃ������ʒu�ɓ���)
        ret = (collection_matched.Item(0).FirstIndex + 1)
    
    End If
    
    '�ԋp�l���i�[���ďI��
    regexFirstIndexOfMatched = ret
    
End Function

'
' �w�蕶����𐳋K�\���������čŏ��Ɍ�������������̒�����Ԃ�
' ������Ȃ������ꍇ�� #VALUE! ��Ԃ�
'
' ## Parameters
'
'  - keyWord
'      �����L�[���[�h
'  - fromThisString
'      �����Ώ۔͈�
'  - ignoreCase default True
'      �啶��/����������ʂ���
'
' ## Returns
'
'  �ŏ��Ɍ�������������̕�����
'
Public Function regexLengthOfMatched( _
    ByVal keyword As String, _
    ByVal fromThisString As String, _
    Optional ByVal ignoreCase As Boolean = True _
    ) As Variant
    
    Dim ret As Variant
    
    Set collection_matched = func_regex_execute(keyword, fromThisString, ignoreCase)
    
    If collection_matched.Count() = 0 Then ' 1�� match ���Ȃ�������
        
        '#VALUE!��Ԃ�( Excel���� `FIND` �� ������Ȃ��������̕ԋp�l�ɓ���)
        ret = CVErr(xlErrValue)
    
    Else ' 1�ȏ� match ������
    
        '�}�b�`����������̒�����Ԃ�
        ret = (collection_matched.Item(0).Length)
    
    End If
    
    '�ԋp�l���i�[���ďI��
    regexLengthOfMatched = ret
    
End Function

'
' �w�蕶����𐳋K�\���������Ēu������
'
' ## Parameters
'
'  - fromThisString
'      �����Ώ۔͈�
'  - keyWord
'      �����L�[���[�h
'  - replacer
'      �u���L�[���[�h
'  - ignoreCase default True
'      �啶��/����������ʂ���
'
' ## Returns
'
'  �u����̕�����B
'  �u���Ώە����񂪌�����Ȃ������ꍇ�́A�u�����Ȃ��܂ܕԂ��B
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
' �w�蕶����𐳋K�\���������� MatchCollection ��Ԃ�
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
' �w�蕶����𐳋K�\���������Ēu������
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

