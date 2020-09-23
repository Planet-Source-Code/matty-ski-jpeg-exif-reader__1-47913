VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "JPeg info"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFile 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   9810
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Picture 004.jpg"
      Top             =   120
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open a JPeg Picture"
      Filter          =   "All Files (*.*)"
      FilterIndex     =   2
      Flags           =   4100
   End
   Begin VB.Label Label1 
      Caption         =   "Read File:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim ByteFormat As Byte

Private Sub cmdFile_Click()
    On Error Resume Next
    CommonDialog.ShowOpen
    If Err Then Exit Sub
    
    txtFile = CommonDialog.FileName
End Sub

Private Sub Command1_Click()
    If Dir(txtFile) = "" Then Beep: Exit Sub
    Dim Cmds As String
    List1.Clear
    Open txtFile For Binary Access Read As #1
    
    
    
    Cmds = Input(2, 1)
    If Cmds <> H2B("FFD8") Then List1.AddItem "Not a JPEG"
    
    Cmds = Input(2, 1) 'FF E0 to FF EF = 'Application Marker'
    If Cmds <> H2B("FFE1") Then List1.AddItem "Contains no 'Application Marker'"
    
    AppDataLen = B2D(Input(2, 1)) ' Motorola byte
    List1.AddItem "Application Data Length = " & AppDataLen ' (-6 from below)
    
    Cmds = Input(6, 1)
    If Cmds <> "Exif" & H2B("0000") Then List1.AddItem "Not Exif data"
    
    
    
    
    '49492A00 08000000 (Common TIFF Header)
    
    'MainOffset = CLng("&H" & "0C")
    ExifDataChunk = Input(AppDataLen, 1)
    
    Select Case Mid$(ExifDataChunk, 1, 2)
    Case H2B("4949"): List1.AddItem "Intel Header Format": ByteFormat = 0 ' Reverse bytes
    Case H2B("4D4D"): List1.AddItem "Motarola Header Format - Might have probs": ByteFormat = 1
    Case Else: List1.AddItem "Unknown/Error Header Format"
    End Select
    
    'If Mid$(ExifDataChunk, 3, 2) <> H2B("2A00") Then List1.AddItem "Header Problem"
    
    FID = B2D(Rev(Mid$(ExifDataChunk, 5, 4))) 'Image File Directory Offset = 8
    List1.AddItem "Image File Dir. Offset = " & FID ' (-8)
    
    NoOfDirEntries = B2D(Rev(Mid$(ExifDataChunk, 9, 2)))
    List1.AddItem "No Of Dir Entries = " & NoOfDirEntries
    
    Dim DataFormat As Long
    Dim tmpStr As String
    Dim NxtExifChunk As Long
    For I = 0 To NoOfDirEntries - 1
        DirEntryInfo = Mid$(ExifDataChunk, (I * 12) + 11, 12)
        ' Dir Entry Order
        'List1.AddItem "Dir Entry Data = " & I & " = " & DirEntryInfo ' Dump Data
        'List1.AddItem "Tagger = " & Hex(B2D(Rev(Mid$(DirEntryInfo, 1, 2))))
        'List1.AddItem "Format = " & B2D(Rev(Mid$(DirEntryInfo, 3, 2)))
        'List1.AddItem "Number of Components (1,2,4,X Bytes) = " & B2D(Rev(Mid$(DirEntryInfo, 5, 4)))
        'List1.AddItem "Value(<=4B) / Offset = " & B2D(Rev(Mid$(DirEntryInfo, 9, 4)))
        
        TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)))
        
        DataFormat = B2D(Rev(Mid$(DirEntryInfo, 3, 2))) ' Byte, Single, Long...
        SizeMultiplier = B2D(Rev(Mid$(DirEntryInfo, 5, 4)))
        LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier
        
        If TagName = "ExifOffset" Then NxtExifChunk = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
        If LenOfTagData <= 4 Then ' No Offset
            List1.AddItem TagName & " = " & ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
        Else ' Offset required
            tmpStr = Mid$(ExifDataChunk, B2D(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData)
            List1.AddItem TagName & " = " & ConvertData2Format(DataFormat, tmpStr)
        End If
    Next I
    
    
    
    NxtIFDO = B2D(Rev(Mid$(ExifDataChunk, (I * 12) + 11, 4)))
    List1.AddItem "Next IFD Offset? = " & NxtIFDO ' Seems incorrect, so its saved above
    If NxtIFDO = 0 Then List1.AddItem "No more IFD entires?"
    
    
    
    List1.AddItem ""
    
    
    
    'FID = B2D(Rev(Mid$(ExifDataChunk, NxtExifChunk + 11, 4))) 'Image File Directory
    'List1.AddItem "Image File Dir. Offset = " & FID
    
    NoOfDirEntries = B2D(Rev(Mid$(ExifDataChunk, NxtExifChunk + 1, 2)))
    List1.AddItem "No Of Dir Entries = " & NoOfDirEntries
    
    For I = 0 To NoOfDirEntries - 1
        DirEntryInfo = Mid$(ExifDataChunk, (I * 12) + NxtExifChunk + 11 + 4, 12)
        
        TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)))
        
        DataFormat = B2D(Rev(Mid$(DirEntryInfo, 3, 2))) ' Byte, Single, Long...
        SizeMultiplier = B2D(Rev(Mid$(DirEntryInfo, 5, 4)))
        LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier
        
        If LenOfTagData <= 4 Then ' No Offset
            List1.AddItem TagName & " = " & ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
        Else ' Offset required
            tmpStr = Mid$(ExifDataChunk, B2D(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData)
            List1.AddItem TagName & " = " & ConvertData2Format(DataFormat, tmpStr)
        End If
    Next I
    
    NxtIFDO = B2D(Rev(Mid$(ExifDataChunk, NxtExifChunk + (I * 12), 4)))
    List1.AddItem "Next IFD Offset? = " & NxtIFDO ' Seems incorrect, so its saved above
    If NxtIFDO = 0 Then List1.AddItem "No more IFD entires?"
    
    
    
    
    
    
    List1.AddItem ""
    Close #1
End Sub

Private Function H2B(InHex As String) As String ' Conv Hex to Bytes
    Dim I As Long
    
    For I = 1 To Len(InHex) Step 2
        H2B = H2B & Chr$(CLng("&H" & Mid$(InHex, I, 2)))
    Next I
End Function

Private Function B2D(InBytes As String) As Double ' Conv. Bytes to Decimal - Could be > 4 Billion
    Dim I As Long
    Dim tmp As String
    
    For I = 1 To Len(InBytes)
        tmp = tmp & Hex(Format$(Asc(Mid$(InBytes, I, 1)), "00"))
    Next I
    B2D = "&H" & tmp
End Function

Private Function Rev(InBytes As String) As String ' Reverse bytes
    If ByteFormat = 1 Then Exit Function ' Not needed for Motorola format
    
    Dim I As Long
    Dim tmp As String
    
    For I = Len(InBytes) To 1 Step -1
        tmp = tmp & Mid$(InBytes, I, 1)
    Next I
    Rev = tmp
End Function

Private Function GetTagName(TagNum As String) As String
    Select Case TagNum
    Case H2B("010E"): GetTagName = "ImageDescription"
    Case H2B("010F"): GetTagName = "Make"
    Case H2B("0110"): GetTagName = "Model"
    Case H2B("0112"): GetTagName = "Orientation"
    Case H2B("011A"): GetTagName = "XResolution"
    Case H2B("011B"): GetTagName = "YResolution"
    Case H2B("0128"): GetTagName = "ResolutionUnit"
    Case H2B("0131"): GetTagName = "Software"
    Case H2B("0132"): GetTagName = "DateTime"
    Case H2B("013E"): GetTagName = "WhitePoint"
    Case H2B("013F"): GetTagName = "PrimaryChromaticities"
    Case H2B("0211"): GetTagName = "YCbCrCoefficients"
    Case H2B("0213"): GetTagName = "YCbCrPositioning"
    Case H2B("0214"): GetTagName = "ReferenceBlackWhite"
    Case H2B("8298"): GetTagName = "Copyright"
    Case H2B("8769"): GetTagName = "ExifOffset"
    
    Case H2B("829A"): GetTagName = "ExposureTime"
    Case H2B("829D"): GetTagName = "FNumber"
    Case H2B("8822"): GetTagName = "ExposureProgram"
    Case H2B("8827"): GetTagName = "ISOSpeedRatings"
    Case H2B("9000"): GetTagName = "ExifVersion"
    Case H2B("9003"): GetTagName = "DateTimeOriginal"
    Case H2B("9004"): GetTagName = "DateTimeDigitized"
    Case H2B("9101"): GetTagName = "ComponentConfiguration"
    Case H2B("9102"): GetTagName = "CompressedBitsPerPixel"
    Case H2B("9201"): GetTagName = "ShutterSpeedValue"
    Case H2B("9202"): GetTagName = "ApertureValue"
    Case H2B("9203"): GetTagName = "BrightnessValue"
    Case H2B("9204"): GetTagName = "ExposureBiasValue"
    Case H2B("9205"): GetTagName = "MaxApertureValue"
    Case H2B("9206"): GetTagName = "SubjectDistance"
    Case H2B("9207"): GetTagName = "MeteringMode"
    Case H2B("9208"): GetTagName = "LightSource"
    Case H2B("9209"): GetTagName = "Flash"
    Case H2B("920A"): GetTagName = "FocalLength"
    Case H2B("927C"): GetTagName = "MakerNote" ': Stop
    Case H2B("9286"): GetTagName = "UserComment"
    Case H2B("A000"): GetTagName = "FlashPixVersion"
    Case H2B("A001"): GetTagName = "ColorSpace"
    Case H2B("A002"): GetTagName = "ExifImageWidth"
    Case H2B("A003"): GetTagName = "ExifImageHeight"
    Case H2B("A004"): GetTagName = "RelatedSoundFile"
    Case H2B("A005"): GetTagName = "ExifInteroperabilityOffset"
    Case H2B("A20E"): GetTagName = "FocalPlaneXResolution"
    Case H2B("A20F"): GetTagName = "FocalPlaneYResolution"
    Case H2B("A210"): GetTagName = "FocalPlaneResolutionUnit"
    Case H2B("A217"): GetTagName = "SensingMethod"
    Case H2B("A300"): GetTagName = "FileSource"
    Case H2B("A301"): GetTagName = "SceneType"
    
    Case H2B("0100"): GetTagName = "ImageWidth"
    Case H2B("0101"): GetTagName = "ImageLength"
    Case H2B("0102"): GetTagName = "BitsPerSample"
    Case H2B("0103"): GetTagName = "Compression"
    Case H2B("0106"): GetTagName = "PhotometricInterpretation"
    Case H2B("0111"): GetTagName = "StripOffsets"
    Case H2B("0115"): GetTagName = "SamplesPerPixel"
    Case H2B("0116"): GetTagName = "RowsPerStrip"
    Case H2B("0117"): GetTagName = "StripByteConunts"
    Case H2B("011A"): GetTagName = "XResolution"
    Case H2B("011B"): GetTagName = "YResolution"
    Case H2B("011C"): GetTagName = "PlanarConfiguration"
    Case H2B("0128"): GetTagName = "ResolutionUnit"
    Case H2B("0201"): GetTagName = "JpegIFOffset"
    Case H2B("0202"): GetTagName = "JpegIFByteCount"
    Case H2B("0211"): GetTagName = "YCbCrCoefficients"
    Case H2B("0212"): GetTagName = "YCbCrSubSampling"
    Case H2B("0213"): GetTagName = "YCbCrPositioning"
    Case H2B("0214"): GetTagName = "ReferenceBlackWhite"
    
    Case H2B("00FE"): GetTagName = "NewSubfileType"
    Case H2B("00FF"): GetTagName = "SubfileType"
    Case H2B("012D"): GetTagName = "TransferFunction"
    Case H2B("013B"): GetTagName = "Artist"
    Case H2B("013D"): GetTagName = "Predictor"
    Case H2B("0142"): GetTagName = "TileWidth"
    Case H2B("0143"): GetTagName = "TileLength"
    Case H2B("0144"): GetTagName = "TileOffsets"
    Case H2B("0145"): GetTagName = "TileByteCounts"
    Case H2B("014A"): GetTagName = "SubIFDs"
    Case H2B("015B"): GetTagName = "JPEGTables"
    Case H2B("828D"): GetTagName = "CFARepeatPatternDim"
    Case H2B("828E"): GetTagName = "CFAPattern"
    Case H2B("828F"): GetTagName = "BatteryLevel"
    Case H2B("83BB"): GetTagName = "IPTC/NAA"
    Case H2B("8773"): GetTagName = "InterColorProfile"
    Case H2B("8824"): GetTagName = "SpectralSensitivity"
    Case H2B("8825"): GetTagName = "GPSInfo"
    Case H2B("8828"): GetTagName = "OECF"
    Case H2B("8829"): GetTagName = "Interlace"
    Case H2B("882A"): GetTagName = "TimeZoneOffset"
    Case H2B("882B"): GetTagName = "SelfTimerMode"
    Case H2B("920B"): GetTagName = "FlashEnergy"
    Case H2B("920C"): GetTagName = "SpatialFrequencyResponse"
    Case H2B("920D"): GetTagName = "Noise"
    Case H2B("9211"): GetTagName = "ImageNumber"
    Case H2B("9212"): GetTagName = "SecurityClassification"
    Case H2B("9213"): GetTagName = "ImageHistory"
    Case H2B("9214"): GetTagName = "SubjectLocation"
    Case H2B("9215"): GetTagName = "ExposureIndex"
    Case H2B("9216"): GetTagName = "TIFF/EPStandardID"
    Case H2B("9290"): GetTagName = "SubSecTime"
    Case H2B("9291"): GetTagName = "SubSecTimeOriginal"
    Case H2B("9292"): GetTagName = "SubSecTimeDigitized"
    Case H2B("A20B"): GetTagName = "FlashEnergy"
    Case H2B("A20C"): GetTagName = "SpatialFrequencyResponse"
    Case H2B("A214"): GetTagName = "SubjectLocation"
    Case H2B("A215"): GetTagName = "ExposureIndex"
    Case H2B("A302"): GetTagName = "CFAPattern"
    
    Case H2B("0200"): GetTagName = "SpecialMode"
    Case H2B("0201"): GetTagName = "JpegQual"
    Case H2B("0202"): GetTagName = "Macro"
    Case H2B("0203"): GetTagName = "Unknown"
    Case H2B("0204"): GetTagName = "DigiZoom"
    Case H2B("0205"): GetTagName = "Unknown"
    Case H2B("0206"): GetTagName = "Unknown"
    Case H2B("0207"): GetTagName = "SoftwareRelease"
    Case H2B("0208"): GetTagName = "PictInfo"
    Case H2B("0209"): GetTagName = "CameraID"
    Case H2B("0F00"): GetTagName = "DataDump"
    'Case H2B(""): GetTagName = ""
    Case Else: GetTagName = "Unknown"
    End Select
End Function

Private Function TypeOfTag(InDec As Long) As Byte
    'Format Info
    'Value              1               2               3               4               5                   6
    'Format             unsigned byte   ascii Strings   unsigned Short  unsigned long   unsigned rational   signed byte
    'Bytes/component    1               1               2               4               8                   1
    
    'Value              7               8               9               10              11                  12
    'Format             undefined       signed Short    signed long     signed rational single float        double float
    'Bytes/component    1               2               4               8               4                   8
    
    Select Case InDec
    Case 1:  TypeOfTag = 1
    Case 2:  TypeOfTag = 1
    Case 3:  TypeOfTag = 2
    Case 4:  TypeOfTag = 4
    Case 5:  TypeOfTag = 8
    Case 6:  TypeOfTag = 1
    Case 7:  TypeOfTag = 1
    Case 8:  TypeOfTag = 2
    Case 9:  TypeOfTag = 4
    Case 10: TypeOfTag = 8
    Case 11: TypeOfTag = 4
    Case 12: TypeOfTag = 8
    End Select
End Function

Private Function ConvertData2Format(DataFormat As Long, InBytes As String) As String
    ' Read function aboves details
    ' Double check for Motorola format esp. CopyMemory
    Dim tmpInt As Integer
    Dim tmpLng As Long
    Dim tmpSng As Single
    Dim tmpDbl As Double
    
    Select Case DataFormat
    Case 1, 3, 4: ConvertData2Format = B2D(InBytes)
    Case 2, 7: ConvertData2Format = InBytes
    Case 5 ' Kinda Unsigned Fraction
        ConvertData2Format = CDbl(B2D(Mid$(InBytes, 1, 4))) / CDbl(B2D(Mid$(InBytes, 5, 4)))
    Case 6
        tmpVal = B2D(InBytes)
        If tmpVal > 127 Then ConvertData2Format = -(tmpVal - 127) Else Convert = tmpVal
    Case 8
        'tmpVal = B2D(InBytes)
        'If tmpVal > 32767 Then ConvertData2Format = -(tmpVal - 32767) Else ConvertData2Format = tmpVal
        CopyMemory tmpInt, InBytes, 2
        ConvertData2Format = tmpInt
    Case 9
        CopyMemory tmpLng, InBytes, 4
        ConvertData2Format = tmpLng
    Case 10 ' Kinda Signed Fraction (Lens Apeture?)
        CopyMemory tmpLng, Mid$(InBytes, 1, 4), 4
        ConvertData2Format = tmpLng
        CopyMemory tmpLng, Mid$(InBytes, 5, 4), 4
        ConvertData2Format = ConvertData2Format / tmpLng
    Case 11
        CopyMemory tmpSng, InBytes, 4
        ConvertData2Format = tmpSng
    Case 12
        CopyMemory tmpDbl, InBytes, 8
        Convert = tmpDbl
    End Select
End Function
