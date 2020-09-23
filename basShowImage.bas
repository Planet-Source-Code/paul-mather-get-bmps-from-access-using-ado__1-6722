Attribute VB_Name = "basDatabaseGetBmp"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Type PT
  Width As Integer
  Height As Integer
End Type

Type OBJECTHEADER
  Signature As Integer
  HeaderSize As Integer
  ObjectType As Long
  NameLen As Integer
  ClassLen As Integer
  NameOffset As Integer
  ClassOFfset As Integer
  ObjectSize As PT
  OleInfo As String * 256
End Type
 
Function DisplayBitmap(ByVal OleField As Variant, ByVal destinationFile As String) As Boolean
Dim Arr() As Byte
Dim ObjHeader As OBJECTHEADER
Dim Buffer As String
Dim ObjectOffset As Long
Dim BitmapOffset As Long
Dim BitmapHeaderOffset As Integer
Dim ArrBmp() As Byte
Dim i As Long

    On Error GoTo e_Trap
    'Resize the array, then fill it with
    'the entire contents of the field
    ReDim Arr(OleField.ActualSize)
    Arr() = OleField.GetChunk(OleField.ActualSize)

    'Copy the first 19 bytes into a variable
    'of the OBJECTHEADER user defined type.
    CopyMemory ObjHeader, Arr(0), 19

    'Determine where the Access Header ends.
    ObjectOffset = ObjHeader.HeaderSize + 1

    'Grab enough bytes after the OLE header to get the bitmap header.
    Buffer = ""
    For i = ObjectOffset To ObjectOffset + 512
        Buffer = Buffer & Chr(Arr(i))
    Next i

    'Make sure the class of the object is a Paint Brush object
    If Mid(Buffer, 12, 6) = "PBrush" Then
        BitmapHeaderOffset = InStr(Buffer, "BM")
        If BitmapHeaderOffset > 0 Then

            'Calculate the beginning of the bitmap
            BitmapOffset = ObjectOffset + BitmapHeaderOffset - 1

            'Move the bitmap into its own array
            ReDim ArrBmp(UBound(Arr) - BitmapOffset)
            CopyMemory ArrBmp(0), Arr(BitmapOffset), UBound(Arr) - BitmapOffset + 1

            'Return the bitmap
            DisplayBitmap = ArrBmp
        End If
    End If
    
    Open destinationFile For Binary As #1
    Put #1, , ArrBmp
    Close #1
    DisplayBitmap = True
    Exit Function
    
e_Trap:
    DisplayBitmap = False
End Function
 
