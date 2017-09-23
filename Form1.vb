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
    ''' 结构体转化为字节数组
    ''' </summary>
    ''' <param name="objStruct"></param>
    ''' <param name="bytes"></param>
    Private Sub StructToBytes(ByVal objStruct As Object, ByRef bytes() As Byte)
        '得到结构体长度  
        Dim size As Integer = Marshal.SizeOf(objStruct)
        '重新定义数组大小  
        ReDim bytes(size)
        '分配结构体大小的内存空间  
        Dim ptrStruct As IntPtr = Marshal.AllocHGlobal(size)
        '将结构体放入内存空间  
        Marshal.StructureToPtr(objStruct, ptrStruct, False)
        '将内存拷贝到字节数组  
        Marshal.Copy(ptrStruct, bytes, 0, size)
        '释放内存空间  
        Marshal.FreeHGlobal(ptrStruct)
    End Sub

    ''' <summary>
    ''' 字节数组bytes()转化为结构体 
    ''' </summary>
    ''' <param name="bytes"></param>
    ''' <param name="mytype"></param>
    ''' <returns></returns>
    Private Function BytesToStruct(ByRef bytes() As Byte, ByVal mytype As Type) As Object
        '获取结构体大小  
        Dim size As Integer = Marshal.SizeOf(mytype)
        'bytes数组长度小于结构体大小  
        If (size > bytes.Length) Then
            '返回空  
            Return Nothing
        End If
        '分配结构体大小的内存空间  
        Dim ptrStruct As IntPtr = Marshal.AllocHGlobal(size)
        '将数组拷贝到内存  
        Marshal.Copy(bytes, 0, ptrStruct, size)
        '将内存转换为目标结构体  
        Dim obj As Object = Marshal.PtrToStructure(ptrStruct, mytype)
        '释放内存  
        Marshal.FreeHGlobal(ptrStruct)
        Return obj
    End Function

    ''' <summary>  
    ''' 从文件中读取结构体  
    ''' </summary>  
    ''' <returns>结构体</returns>  
    ''' <remarks></remarks>  
    Private Function FileToStruct(ByVal strPath As String, ByVal mytype As Type) As Object
        '获得结构体大小  
        Dim size As Integer = Marshal.SizeOf(mytype)
        '打开文件流  
        Dim fs As New FileStream(strPath, FileMode.Open)
        '读取size个字节  
        Dim bytes(size) As Byte
        fs.Read(bytes, 0, size)
        '将读取的字节转化为结构体  
        Dim obj As Object = BytesToStruct(bytes, mytype)
        fs.Close()

        Return obj
    End Function
End Class