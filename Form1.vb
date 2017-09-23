Imports System.IO
Imports System.Runtime.InteropServices

Public Class Form1

    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure AbnParm
        Dim Abn1Wave As Short
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure Ranks
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=10)> Public Mark As String
        Dim Threshold As Integer
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure ExParm
        <VBFixedString(EXP_PAR_MAX), MarshalAs(UnmanagedType.ByValTStr, SizeConst:=EXP_PAR_MAX)> Public Unit As String
        Dim Par As Short
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure tItemInf
        Dim Ver As Short
        Dim Size As Integer
        Dim OnOff As Short
        <VBFixedString(ITEM_NAME_MAX), MarshalAs(UnmanagedType.ByValTStr, SizeConst:=ITEM_NAME_MAX)> Public Name As String
        Dim MeasWave As Short
        Dim UpDown As Short
        <VBFixedArray(RANK_MAX), MarshalAs(UnmanagedType.ByValArray, SizeConst:=RANK_MAX)> Dim RankTbl() As Ranks
        <VBFixedArray(1), MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)> Dim ExParmDat() As ExParm
        Dim DataPos As Short

        Public Sub Initialize()
            ReDim RankTbl(RANK_MAX)
            ReDim ExParmDat(1)
        End Sub
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure tColorTone
        Dim NameNo As Short
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure tColorInf
        Dim Ver As Short
        <VBFixedArray(COLOR_AREA_MAX), MarshalAs(UnmanagedType.ByValArray, SizeConst:=COLOR_AREA_MAX)> Dim ColTone() As tColorTone
        Dim VioAngMin As Integer

        Public Sub Initialize()
            ReDim ColTone(COLOR_AREA_MAX)
        End Sub
    End Structure


    Public Const VER_SIZE As Integer = 10
    Public Const KDK_ORIGINAL As Integer = 12
    Public Const FILE_COMMENT As Integer = 15
    Public Const PAPER_TYPE_MAX As Integer = 20
    Public Const ITEM_TYPE_MAX As Integer = 30
    Public Const COLOR_AREA_MAX As Integer = 30
    Public Const RANK_MAX As Integer = 30
    Public Const ITEM_NAME_MAX As Integer = 30
    Public Const EXP_PAR_MAX As Integer = 20


    <StructLayoutAttribute(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure tRecord
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=VER_SIZE)> Public rVersion As String

        <VBFixedArray(ITEM_TYPE_MAX), MarshalAs(UnmanagedType.ByValArray, SizeConst:=ITEM_TYPE_MAX)>
        Dim rItemInf() As tItemInf

        Dim rColorInf As tColorInf

        Public Sub Initialize()
            Dim newChan1 As New tItemInf
            ReDim rItemInf(ITEM_TYPE_MAX - 1)
            For i As Integer = 0 To ITEM_TYPE_MAX - 1
                rItemInf(i) = newChan1
            Next
            rColorInf.Initialize()
        End Sub
    End Structure


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OpenFileDialog1.ShowDialog()
        Dim Record As tRecord = New tRecord

        Record.Initialize()

        Record = FileToStruct(OpenFileDialog1.FileName, GetType(tRecord))
    End Sub

    ''' <summary>
    ''' �ṹ��ת��Ϊ�ֽ�����
    ''' </summary>
    ''' <param name="objStruct"></param>
    ''' <param name="bytes"></param>
    Private Sub StructToBytes(ByVal objStruct As Object, ByRef bytes() As Byte)
        '�õ��ṹ�峤��  
        Dim size As Integer = Marshal.SizeOf(objStruct)
        '���¶��������С  
        ReDim bytes(size)
        '����ṹ���С���ڴ�ռ�  
        Dim ptrStruct As IntPtr = Marshal.AllocHGlobal(size)
        '���ṹ������ڴ�ռ�  
        Marshal.StructureToPtr(objStruct, ptrStruct, False)
        '���ڴ濽�����ֽ�����  
        Marshal.Copy(ptrStruct, bytes, 0, size)
        '�ͷ��ڴ�ռ�  
        Marshal.FreeHGlobal(ptrStruct)
    End Sub

    ''' <summary>
    ''' �ֽ�����bytes()ת��Ϊ�ṹ�� 
    ''' </summary>
    ''' <param name="bytes"></param>
    ''' <param name="mytype"></param>
    ''' <returns></returns>
    Private Function BytesToStruct(ByRef bytes() As Byte, ByVal mytype As Type) As Object
        '��ȡ�ṹ���С  
        Dim size As Integer = Marshal.SizeOf(mytype)
        'bytes���鳤��С�ڽṹ���С  
        If (size > bytes.Length) Then
            '���ؿ�  
            Return Nothing
        End If
        '����ṹ���С���ڴ�ռ�  
        Dim ptrStruct As IntPtr = Marshal.AllocHGlobal(size)
        '�����鿽�����ڴ�  
        Marshal.Copy(bytes, 0, ptrStruct, size)
        '���ڴ�ת��ΪĿ��ṹ��  
        Dim obj As Object = Marshal.PtrToStructure(ptrStruct, mytype)
        '�ͷ��ڴ�  
        Marshal.FreeHGlobal(ptrStruct)
        Return obj
    End Function

    ''' <summary>  
    ''' ���ļ��ж�ȡ�ṹ��  
    ''' </summary>  
    ''' <returns>�ṹ��</returns>  
    ''' <remarks></remarks>  
    Private Function FileToStruct(ByVal strPath As String, ByVal mytype As Type) As Object
        '��ýṹ���С  
        Dim size As Integer = Marshal.SizeOf(mytype)
        '���ļ���  
        Dim fs As New FileStream(strPath, FileMode.Open)
        '��ȡsize���ֽ�  
        Dim bytes(size) As Byte
        fs.Read(bytes, 0, size)
        '����ȡ���ֽ�ת��Ϊ�ṹ��  
        Dim obj As Object = BytesToStruct(bytes, mytype)
        fs.Close()

        Return obj
    End Function
End Class