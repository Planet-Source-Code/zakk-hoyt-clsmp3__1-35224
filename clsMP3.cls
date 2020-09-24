VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The purpose of this class is to extract information from MP3s. It was written to do the dirty
'work in an ASP MP3 streamer.

'1. ReadMP3() - Reads ID3v1, ID3v2, bitrate, variable bitrate, freq, and duration of a specified
'MP3. If the MP3 has v1 and v2 tags, then v2 will override v1.

'2. SearchDir() - Give this function a dir, and a file type to look for,
'and it will return the subdirs (pipe delimited), and the files (pipe delimited)






Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WIN32 Types, Constants, and Declares
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const INVALID_HANDLE_VALUE = -1
Private Const DDL_DIRECTORY = &H10
Private Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'application level variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GetMP3Info As MP3Info
Private Const ErrBase = 9994
Private Type MP3Info
    Songname As String
    Artist As String
    Album As String
    Year As String
    Comment As String
    Genre As String
    Track As String
    VBR As Boolean
End Type

Public Duration As String
Public Frequency As String
Public Mode As String
Public MpegLayer As String
Public MpegVersion As String
Public BitRate As String
Public Songname As String
Public Artist As String
Public Album As String
Public Year As String
Public Comment As String
Public Genre As String
Public Track As String
Public VBR As Boolean

Public Files As String
Public Subdirs As String



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SearchDir(ByVal RootDir As String, ByVal FileType As String) As Boolean
    On Error GoTo ErrHand
    
    DoEvents
    Dim Dir_Names() As String
    Dim Num_Dirs As Integer
    Dim I As Integer
    Dim fName As String
    Dim Search_Handle As Long
    Dim File_Data As WIN32_FIND_DATA
    Dim TempInt As Integer
    
    'clear vars because they are passed by ref and could contain data
    Subdirs = ""
    Files = ""

    ' Get the matching files in this directory.
    Search_Handle = FindFirstFile(RootDir & FileType, File_Data)
    If Search_Handle <> INVALID_HANDLE_VALUE Then
        Do
            fName = File_Data.cFileName
            fName = RootDir & Left$(fName, InStr(fName, Chr$(0)) - 1)
            
            'add the file to the string
            Files = Files & fName & "|"
            
            'find the next file
            If FindNextFile(Search_Handle, File_Data) = 0 Then Exit Do
        Loop
        FindClose Search_Handle
    End If

    ' Get the list of subdirectories.
    Search_Handle = FindFirstFile(RootDir & "*.*", File_Data)
    If Search_Handle <> INVALID_HANDLE_VALUE Then
        Do
            If File_Data.dwFileAttributes And DDL_DIRECTORY Then
                fName = File_Data.cFileName
                fName = Left$(fName, InStr(fName, Chr$(0)) - 1)
                If fName <> "." And fName <> ".." Then
                    Subdirs = Subdirs & fName & "|"
                End If
            End If
            DoEvents
            If FindNextFile(Search_Handle, File_Data) = 0 Then Exit Do
        Loop
        FindClose Search_Handle
    End If
    
    SearchDir = True
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function

Public Function ReadMP3(ByVal FileName As Variant) As Boolean
    
    On Error GoTo ErrHand
    
    'clear out mp3 info
    BitRate = ""
    Frequency = ""
    Duration = ""
    Songname = ""
    Artist = ""
    Album = ""
    Year = ""
    Comment = ""
    Genre = ""
    Track = ""
    MpegVersion = ""
    MpegLayer = ""
    VBR = False
    Mode = ""

    'reads header (duration, freq, bitrate, version, layer are stored in the header)
    ReadHeader FileName
    
    'reads the ID3 tag (artist, album, year, comment, genre, track are stored in ID3 tag)
    ReadID3 FileName

    ReadMP3 = True
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description
    
    
End Function





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private (Encapsulted) functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ReadID3(ByVal FileName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim I As Integer                'holds temp int bytes
    Dim B As Byte                   'holds temp byte values
    Dim S As String                 'holds temp strings
    Dim R As Long
    Dim X As Long
    Dim Version As Byte
    Dim TagString As String         'holds the string read from the mp3
    Dim FieldSize As Long
    Dim SizeOffset As Long
    Dim FieldOffset As Long
    Dim strStrip As String
    Dim FFile As Long
    Dim TempString As String
    
    'fixed length vars for reading v1ID3v1 tags
    Dim v1ID As String * 3
    Dim v1Title As String * 30
    Dim v1Artist As String * 30
    Dim v1Album As String * 30
    Dim v1Year As String * 4
    Dim v1Comment As String * 30
    Dim v1Genre As Byte
    'Dim Genre As Byte
    
    'vars that hold v1ID3v2 attibutes
    Dim v2Title As String
    Dim v2Artist As String
    Dim v2Album As String
    Dim v2Year As String
    Dim v2Genre As String
    Dim v2Track As String
    
    FFile = FreeFile
    
    'Open the file so we can read it
    Open FileName For Binary As FFile
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Check for an ID3v1 tag
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Get FFile, LOF(FFile) - 127, v1ID
        If v1ID = "TAG" Then 'If "TAG" is present, then we have a valv1ID v1ID3v1 tag and will extract all available v1ID3v1 info from the file
            Get FFile, , v1Title    'Always limited to 30 characters
            Get FFile, , v1Artist   'Always limited to 30 characters
            Get FFile, , v1Album    'Always limited to 30 characters
            Get FFile, , v1Year     'Always limited to 4 characters
            Get FFile, , v1Comment  'Always limited to 30 characters
            Get FFile, , v1Genre    'Always limited to 1 byte (?)
            
            Songname = v1Title
            Artist = v1Artist
            Album = v1Album
            Year = v1Year
            Comment = v1Comment
            Genre = v1Genre
            
            'get rv1ID of " " on the right sv1IDe of the string
            Do Until Right(Songname, 1) <> " "
                Songname = Left(Songname, Len(Songname) - 1)
            Loop
            Do Until Right(Artist, 1) <> " "
                Artist = Left(Artist, Len(Artist) - 1)
            Loop
            Do Until Right(Album, 1) <> " "
                Album = Left(Album, Len(Album) - 1)
            Loop
            Do Until Right(Year, 1) <> " "
                Year = Left(Year, Len(Year) - 1)
            Loop
            Do Until Right(Comment, 1) <> " "
                Comment = Left(Comment, Len(Comment) - 1)
            Loop
            Do Until Right(Genre, 1) <> " "
                Genre = Left(Genre, Len(Genre) - 1)
            Loop
      
            'get rv1ID of null characters
            Songname = Replace(Songname, vbNullChar, "")
            Artist = Replace(Artist, vbNullChar, "")
            Album = Replace(Album, vbNullChar, "")
            Year = Replace(Year, vbNullChar, "")
            Comment = Replace(Comment, vbNullChar, "")
            Genre = Replace(Genre, vbNullChar, "")
            
            'comments tend to have lots of unprintable chars, so remove them
            If Len(Comment) > 0 Then
                For X = 0 To 47
                    If InStr(Comment, Chr(X)) Then
                        Comment = Replace(Comment, Chr(X), "")
                    End If
                Next
            End If
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Check for a Header for ID3v2 tag
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Get FFile, 2, B
        If (B < 250 Or B > 251) Then
            'We have an v1ID3v2 tag
            If B = 68 Then
                Get FFile, 3, B
                If B = 51 Then
                    'get version
                    Get FFile, 4, Version
                    
                    'get how long the v2 tag is
                    Get FFile, 7, B
                    R = Val(B) * 20917152
                    Get FFile, 8, B
                    R = R + Val(B) * 16384
                    Get FFile, 9, B
                    R = R + Val(B) * 128
                    Get FFile, 10, B
                    R = R + Val(B)
                    
                    'invalid length was calculated
                    If R > LOF(FFile) Or R > 2147483647 Then
                        Exit Function
                    End If
                    
                    'get the v2 tag according to the length calculated
                    TagString = Space$(R)
                    Get FFile, 11, TagString
                End If
            End If
        End If

     
        If TagString <> "" Then     'there is an ID3v2 tag in TagString
            ' Determine if the v1ID3v2 tag is v1ID3v2.2 or v1ID3v2.3
            Select Case Version
                Case 2 'v1ID3v2.2
                    'Set the fieldnames for version 2.0
                    v2Title = "TT2"
                    v2Artist = "TOA"
                    v2Album = "TAL"
                    v2Year = "TYE"
                    v2Genre = "TCO"
                    FieldOffset = 7
                    SizeOffset = 5
                    v2Track = "TRCK"
               
                Case 3 'v1ID3v2.3
                    'Set the fieldnames for version 3.0
                    v2Title = "TIT2"
                    v2Artist = "TPE1"
                    v2Album = "TALB"
                    v2Year = "TYER"
                    v2Genre = "TCON"
                    v2Track = "TRCK"
                    FieldOffset = 11
                    SizeOffset = 7
                    
                Case Else
                    'We don't have a valv1ID v1ID3v2 tag, so bail
                    Exit Function
            End Select
                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Extract track title
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
            I = InStr(TagString, v2Title)
            If I > 0 Then
                'read the title
                FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
                If Version = 3 Then
                   'check for compressed or encrypted field
                   B = Asc(Mid$(TagString, I + 9))
                   If (B And 128) = True Or (B And 64) = True Then GoTo ReadAlbum
                End If
                TempString = Mid$(TagString, I + FieldOffset, FieldSize)
                If TempString <> "" Then
                    Songname = TempString
                End If
            End If
                   
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Extract album title
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadAlbum:
            I = InStr(TagString, v2Album)
            If I > 0 Then
                FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
                If Version = 3 Then
                   'check for compressed or encrypted field
                   B = Asc(Mid$(TagString, I + 9))
                   If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadArtist
                End If
                TempString = Mid$(TagString, I + FieldOffset, FieldSize)
                If TempString <> "" Then
                    Album = TempString
                End If
            End If
                
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Extract artist name
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadArtist:
            I = InStr(TagString, v2Artist)
            If I > 0 Then
                FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
                If Version = 3 Then
                    'check for compressed or encrypted field
                    B = Asc(Mid$(TagString, I + 9))
                    If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadYear
                End If
                TempString = Mid$(TagString, I + FieldOffset, FieldSize)
                If TempString <> "" Then
                    Artist = TempString
                End If
            End If
               
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Extract year title
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadYear:
            I = InStr(TagString, v2Year)
            If I > 0 Then
                FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
                If Version = 3 Then
                    'check for compressed or encrypted field
                    B = Asc(Mid$(TagString, I + 9))
                    If (B And 128) = 128 Or (B And 64) = 64 Then GoTo ReadGenre
                End If
                TempString = Mid$(TagString, I + FieldOffset, FieldSize)
                If TempString <> "" Then
                    Year = TempString
                End If
            End If
               
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Extract genre
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ReadGenre:
            I = InStr(TagString, v2Genre)
            If I > 0 Then
                FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
                
                If Version = 3 Then
                   'check for compressed or encrypted field
                   B = Asc(Mid$(TagString, I + 9))
                   If (B And 128) = 128 Or (B And 64) = 64 Then GoTo Readv2Track
                End If
                
                S = Mid$(TagString, I + FieldOffset, FieldSize)
                
                If Left$(S, 1) = "(" Then
                    If InStr(S, ")") Then
                        'the tag is in the form of "(10)metal"
                        TempString = Right(S, Len(S) - InStr(S, ")"))
                    End If
                Else
                    TempString = S
                End If
                Genre = TempString
            End If
               
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Extract track number
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Readv2Track:
            I = InStr(TagString, v2Track)
            If I > 0 Then
                FieldSize = Asc(Mid$(TagString, I + SizeOffset)) - 1
                If Version = 3 Then
                   'check for compressed or encrypted field
                   B = Asc(Mid$(TagString, I + 9))
                   If (B And 128) = 128 Or (B And 64) = 64 Then GoTo Done
                End If
                TempString = Mid$(TagString, I + FieldOffset, FieldSize)
                If TempString <> "" Then
                    Track = TempString
                End If
            End If
Done:
        End If
        
    Close FFile
    
    'if the genre tag has a # instead of a string, then decode it
    If IsNumeric(Genre) Then
        Genre = GetGenreFromID(Genre)
    End If
    
    Exit Function
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3 (" & FileName & ")"

    Close FFile
    
End Function

Private Function ReadHeader(ByVal FileName As String)
    On Error GoTo ErrHand
    
    Dim ByteArray(4) As Byte
    Dim XingH As String * 4
    Dim FIO As Long
    Dim I As Long
    Dim Z As Long
    Dim X As Byte
    Dim HeadStart As Long
    Dim Frames As Long
    Dim Bin As String
    Dim Temp As Variant
    Dim Brate As Variant
    Dim Freq As Variant
         
    'tables
    Dim VersionLayer(3) As String
    VersionLayer(0) = 0
    VersionLayer(1) = 3
    VersionLayer(2) = 2
    VersionLayer(3) = 1
    
    Dim SMode(3) As String
    SMode(0) = "stereo"
    SMode(1) = "joint stereo"
    SMode(2) = "dual channel"
    SMode(3) = "single channel"
    
    
    FIO = FreeFile
    
    'read the header
    Open FileName For Binary Access Read As FIO
        If LOF(FIO) < 256 Then
            Close FIO
            Exit Function
        End If
        
        '''''start check startposition for header''''''''''''
        '''''if start position <>1 then id3v2 tag exists'''''
        For I = 1 To LOF(FIO)           'check the whole file for the header
            Get #FIO, I, X
            If X = 255 Then             'header always start with 255 followed by 250 or 251
                Get #FIO, I + 1, X
                If X > 249 And X < 252 Then
                    HeadStart = I       'set header start position
                    Exit For
                End If
            End If
        Next I
        
        'no header start position was found
        If HeadStart = 0 Then
            Exit Function
        End If
        '''end check start position for header'''''''''''''
    
        ''start check for XingHeader'''
        Get #FIO, HeadStart + 36, XingH
        If XingH = "Xing" Then
            VBR = True
            For Z = 1 To 4 '
                Get #1, HeadStart + 43 + Z, ByteArray(Z)  'get framelength to array
            Next Z
            Frames = BinToDec(ByteToBit(ByteArray))   'calculate # of frames
        Else
            VBR = False
        End If
        '''end check for XingHeader
    
        '''start extract the first 4 bytes (32 bits) to an array
        For Z = 1 To 4 '
            Get #FIO, HeadStart + Z - 1, ByteArray(Z)
        Next Z
        '''stop extract the first 4 bytes (32 bits) to an array
    Close FIO
    
    'header string
    Bin = ByteToBit(ByteArray)
    
    
    'get mpegversion from table
    MpegVersion = VersionLayer(BinToDec(Mid(Bin, 12, 2)))
    'get layer from table
    MpegLayer = VersionLayer(BinToDec(Mid(Bin, 14, 2)))
    'get mode from table
    Mode = SMode(BinToDec(Mid(Bin, 25, 2)))
    
    'look for version to create right table
    Select Case MpegVersion
        Case 1
            'for version 1
            Freq = Array(44100, 48000, 32000)
        Case 2 Or 25
            'for version 2 or 2.5
            Freq = Array(22050, 24000, 16000)
        Case Else
            Frequency = 0
            Exit Function
    End Select
    
    'look for frequency in table
    Frequency = Freq(BinToDec(Mid(Bin, 21, 2)))
    
    If VBR = True Then
        'define to calculate correct bitrate
        Temp = Array(, 12, 144, 144)
        BitRate = (FileLen(FileName) * Frequency) / (Int(Frames)) / 1000 / Temp(MpegLayer)
    Else
        'look for the right bitrate table
        Select Case Val(MpegVersion & MpegLayer)
            Case 11
                'Version 1, Layer 1
                Brate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
            Case 12
                'V1 L1
                Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
            Case 13
                'V1 L3
                Brate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
            Case 21 Or 251
                'V2 L1 and 'V2.5 L1
                Brate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 144, 160, 176, 192, 224, 256)
            Case 22 Or 252 Or 23 Or 253
                'V2 L2 and 'V2.5 L2 etc...
                Brate = Array(0, 8, 16, 24, 32, 40, 48, 56, 64, 80, 96, 112, 128, 144, 160)
            Case Else
                'if variable bitrate
                BitRate = 1
            Exit Function
        End Select
        
        BitRate = Brate(BinToDec(Mid(Bin, 17, 4)))
    End If
    
    'if there is a decimal place, then parse it off
    If InStr(1, BitRate, ".") Then
        BitRate = Left(BitRate, InStr(1, BitRate, ".") - 1)
    End If
    
    'calculate duration
    Duration = Int((FileLen(FileName) * 8) / BitRate / 1000)
    Duration = Duration \ 60 & "m " & Duration - (Duration \ 60) * 60 & "s"
    
    
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3 (" & FileName & ")"
    Close FIO
    

End Function

Private Function BinToDec(BinValue As String) As Long
    On Error GoTo ErrHand
    
    Dim I As Long
    BinToDec = 0
    For I = 1 To Len(BinValue)
        If Mid(BinValue, I, 1) = 1 Then
            BinToDec = BinToDec + 2 ^ (Len(BinValue) - I)
        End If
    Next I
    
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function

Private Function ByteToBit(ByteArray) As String
    On Error GoTo ErrHand

    Dim Z As Integer
    Dim I As Integer
    'convert 4*1 byte array to 4*8 bits'''''
    ByteToBit = ""
    For Z = 1 To 4
        For I = 7 To 0 Step -1
            If Int(ByteArray(Z) / (2 ^ I)) = 1 Then
                ByteToBit = ByteToBit & "1"
                ByteArray(Z) = ByteArray(Z) - (2 ^ I)
            ElseIf ByteToBit <> "" Then
                ByteToBit = ByteToBit & "0"
            End If
        Next
    Next Z
    
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"

End Function
Private Function GetGenreFromID(ByVal strListGenre As Integer) As String
    On Error GoTo ErrHand
    
    If strListGenre = 0 Then
        GetGenreFromID = "Blues"
    ElseIf strListGenre = 1 Then
        GetGenreFromID = "Classic Rock"
    ElseIf strListGenre = 2 Then
        GetGenreFromID = "Country"
    ElseIf strListGenre = 3 Then
        GetGenreFromID = "Dance"
    ElseIf strListGenre = 4 Then
        GetGenreFromID = "Disco"
    ElseIf strListGenre = 5 Then
        GetGenreFromID = "Funk"
    ElseIf strListGenre = 6 Then
        GetGenreFromID = "Grunge"
    ElseIf strListGenre = 7 Then
        GetGenreFromID = "Hip-Hop"
    ElseIf strListGenre = 8 Then
        GetGenreFromID = "Jazz"
    ElseIf strListGenre = 9 Then
        GetGenreFromID = "Metal"
    ElseIf strListGenre = 10 Then
        GetGenreFromID = "New Age"
    ElseIf strListGenre = 11 Then
        GetGenreFromID = "Oldies"
    ElseIf strListGenre = 12 Then
        GetGenreFromID = "Other"
    ElseIf strListGenre = 13 Then
        GetGenreFromID = "Pop"
    ElseIf strListGenre = 14 Then
        GetGenreFromID = "R&B"
    ElseIf strListGenre = 15 Then
        GetGenreFromID = "Rap"
    ElseIf strListGenre = 16 Then
        GetGenreFromID = "Reggae"
    ElseIf strListGenre = 17 Then
        GetGenreFromID = "Rock"
    ElseIf strListGenre = 18 Then
        GetGenreFromID = "Techno"
    ElseIf strListGenre = 19 Then
        GetGenreFromID = "Industrial"
    ElseIf strListGenre = 20 Then
        GetGenreFromID = "Alternative"
    ElseIf strListGenre = 21 Then
        GetGenreFromID = "Ska"
    ElseIf strListGenre = 22 Then
        GetGenreFromID = "Death Metal"
    ElseIf strListGenre = 23 Then
        GetGenreFromID = "Pranks"
    ElseIf strListGenre = 24 Then
        GetGenreFromID = "Soundtrack"
    ElseIf strListGenre = 25 Then
        GetGenreFromID = "Euro-Techno"
    ElseIf strListGenre = 26 Then
        GetGenreFromID = "Ambient"
    ElseIf strListGenre = 27 Then
        GetGenreFromID = "Trip-Hop"
    ElseIf strListGenre = 28 Then
        GetGenreFromID = "Vocal"
    ElseIf strListGenre = 29 Then
        GetGenreFromID = "Jazz+Funk"
    ElseIf strListGenre = 30 Then
        GetGenreFromID = "Fusion"
    ElseIf strListGenre = 31 Then
        GetGenreFromID = "Trance"
    ElseIf strListGenre = 32 Then
        GetGenreFromID = "Classical"
    ElseIf strListGenre = 33 Then
        GetGenreFromID = "Instrumental"
    ElseIf strListGenre = 34 Then
        GetGenreFromID = "Acid"
    ElseIf strListGenre = 35 Then
        GetGenreFromID = "House"
    ElseIf strListGenre = 36 Then
        GetGenreFromID = "Game"
    ElseIf strListGenre = 37 Then
        GetGenreFromID = "Sound Clip"
    ElseIf strListGenre = 38 Then
        GetGenreFromID = "Gospel"
    ElseIf strListGenre = 39 Then
        GetGenreFromID = "Noise"
    ElseIf strListGenre = 40 Then
        GetGenreFromID = "AlternRock"
    ElseIf strListGenre = 41 Then
        GetGenreFromID = "Bass"
    ElseIf strListGenre = 42 Then
        GetGenreFromID = "Soul"
    ElseIf strListGenre = 43 Then
        GetGenreFromID = "Punk"
    ElseIf strListGenre = 44 Then
        GetGenreFromID = "Space"
    ElseIf strListGenre = 45 Then
        GetGenreFromID = "Meditative"
    ElseIf strListGenre = 46 Then
        GetGenreFromID = "Instrumental Pop"
    ElseIf strListGenre = 47 Then
        GetGenreFromID = "Instrumental Rock"
    ElseIf strListGenre = 48 Then
        GetGenreFromID = "Ethnic"
    ElseIf strListGenre = 49 Then
        GetGenreFromID = "Gothic"
    ElseIf strListGenre = 50 Then
        GetGenreFromID = "Darkwave"
    ElseIf strListGenre = 51 Then
        GetGenreFromID = "Techno-Indistrial"
    ElseIf strListGenre = 52 Then
        GetGenreFromID = "Electronic"
    ElseIf strListGenre = 53 Then
        GetGenreFromID = "Pop-Folk"
    ElseIf strListGenre = 54 Then
        GetGenreFromID = "Eurodance"
    ElseIf strListGenre = 55 Then
        GetGenreFromID = "Dream"
    ElseIf strListGenre = 56 Then
        GetGenreFromID = "Southern Rock"
    ElseIf strListGenre = 57 Then
        GetGenreFromID = "Comedy"
    ElseIf strListGenre = 58 Then
        GetGenreFromID = "Cult"
    ElseIf strListGenre = 59 Then
        GetGenreFromID = "Gansta"
    ElseIf strListGenre = 60 Then
        GetGenreFromID = "Top 40"
    ElseIf strListGenre = 61 Then
        GetGenreFromID = "Christian Rap"
    ElseIf strListGenre = 62 Then
        GetGenreFromID = "Pop/Funk"
    ElseIf strListGenre = 63 Then
        GetGenreFromID = "Jungle"
    ElseIf strListGenre = 64 Then
        GetGenreFromID = "Native American"
    ElseIf strListGenre = 65 Then
        GetGenreFromID = "Cabaret"
    ElseIf strListGenre = 66 Then
        GetGenreFromID = "New Wave"
    ElseIf strListGenre = 67 Then
        GetGenreFromID = "Psychaledic"
    ElseIf strListGenre = 68 Then
        GetGenreFromID = "Rave"
    ElseIf strListGenre = 69 Then
        GetGenreFromID = "Showtunes"
    ElseIf strListGenre = 70 Then
        GetGenreFromID = "Trailer"
    ElseIf strListGenre = 71 Then
        GetGenreFromID = "Lo-Fi"
    ElseIf strListGenre = 72 Then
        GetGenreFromID = "Tribal"
    ElseIf strListGenre = 73 Then
        GetGenreFromID = "Acid Punk"
    ElseIf strListGenre = 74 Then
        GetGenreFromID = "Acid Jazz"
    ElseIf strListGenre = 75 Then
        GetGenreFromID = "Polka"
    ElseIf strListGenre = 76 Then
        GetGenreFromID = "Retro"
    ElseIf strListGenre = 77 Then
        GetGenreFromID = "Musical"
    ElseIf strListGenre = 78 Then
        GetGenreFromID = "Rock & Roll"
    ElseIf strListGenre = 79 Then
        GetGenreFromID = "Hard Rock"
    ElseIf strListGenre = 80 Then
        GetGenreFromID = "Folk"
    ElseIf strListGenre = 81 Then
        GetGenreFromID = "Folk-Rock"
    ElseIf strListGenre = 82 Then
        GetGenreFromID = "National Folk"
    ElseIf strListGenre = 83 Then
        GetGenreFromID = "Swing"
    ElseIf strListGenre = 84 Then
        GetGenreFromID = "Fast Fusion"
    ElseIf strListGenre = 85 Then
        GetGenreFromID = "Bebop"
    ElseIf strListGenre = 86 Then
        GetGenreFromID = "Latin"
    ElseIf strListGenre = 87 Then
        GetGenreFromID = "Revival"
    ElseIf strListGenre = 88 Then
        GetGenreFromID = "Celtic"
    ElseIf strListGenre = 89 Then
        GetGenreFromID = "Bluegrass"
    ElseIf strListGenre = 90 Then
        GetGenreFromID = "Avantgarde"
    ElseIf strListGenre = 91 Then
        GetGenreFromID = "Gothic Rock"
    ElseIf strListGenre = 92 Then
        GetGenreFromID = "Progressive Rock"
    ElseIf strListGenre = 93 Then
        GetGenreFromID = "Psychadelic Rock"
    ElseIf strListGenre = 94 Then
        GetGenreFromID = "Symphonic Rock"
    ElseIf strListGenre = 95 Then
        GetGenreFromID = "Slow Rock"
    ElseIf strListGenre = 96 Then
        GetGenreFromID = "Big Band"
    ElseIf strListGenre = 97 Then
        GetGenreFromID = "Chorus"
    ElseIf strListGenre = 98 Then
        GetGenreFromID = "Easy Listening"
    ElseIf strListGenre = 99 Then
        GetGenreFromID = "Acoustic"
    ElseIf strListGenre = 100 Then
        GetGenreFromID = "Humour"
    ElseIf strListGenre = 101 Then
        GetGenreFromID = "Speech"
    ElseIf strListGenre = 102 Then
        GetGenreFromID = "Chanson"
    ElseIf strListGenre = 103 Then
        GetGenreFromID = "Opera"
    ElseIf strListGenre = 104 Then
        GetGenreFromID = "Chamber Music"
    ElseIf strListGenre = 105 Then
        GetGenreFromID = "Sonata"
    ElseIf strListGenre = 106 Then
        GetGenreFromID = "Symphony"
    ElseIf strListGenre = 107 Then
        GetGenreFromID = "Booty Bass"
    ElseIf strListGenre = 108 Then
        GetGenreFromID = "Primus"
    ElseIf strListGenre = 109 Then
        GetGenreFromID = "Porn Groove"
    ElseIf strListGenre = 110 Then
        GetGenreFromID = "Satire"
    ElseIf strListGenre = 111 Then
        GetGenreFromID = "Slow Jam"
    ElseIf strListGenre = 112 Then
        GetGenreFromID = "Club"
    ElseIf strListGenre = 113 Then
        GetGenreFromID = "Tango"
    ElseIf strListGenre = 114 Then
        GetGenreFromID = "Samba"
    ElseIf strListGenre = 115 Then
        GetGenreFromID = "Folklore"
    ElseIf strListGenre = 116 Then
        GetGenreFromID = "Ballad"
    ElseIf strListGenre = 117 Then
        GetGenreFromID = "Power Ballad"
    ElseIf strListGenre = 118 Then
        GetGenreFromID = "Rhythmic Soul"
    ElseIf strListGenre = 119 Then
        GetGenreFromID = "Freestyle"
    ElseIf strListGenre = 120 Then
        GetGenreFromID = "Duet"
    ElseIf strListGenre = 121 Then
        GetGenreFromID = "Punk Rock"
    ElseIf strListGenre = 122 Then
        GetGenreFromID = "Drum Solo"
    ElseIf strListGenre = 123 Then
        GetGenreFromID = "A Capella"
    ElseIf strListGenre = 124 Then
        GetGenreFromID = "Euro-House"
    ElseIf strListGenre = 125 Then
        GetGenreFromID = "Dance Hall"
    ElseIf strListGenre > 125 Then
        GetGenreFromID = ""
    End If
    
    Exit Function
    
ErrHand:
    Err.Raise ErrBase + Err.Number, "clsMP3", Err.Description & " in clsMP3"


End Function



