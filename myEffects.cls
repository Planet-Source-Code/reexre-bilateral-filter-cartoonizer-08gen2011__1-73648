VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Cartoonizer - Convert Photos into Cartoon Like Images
'    Copyright (c) 2011 - Roberto Mior
'
'    This file is part of "Bilateral Filter - Cartoonizer".
'
'    "Bilateral Filter - Cartoonizer" is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    "Bilateral Filter - Cartoonizer" is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with "Bilateral Filter - Cartoonizer".  If not, see <http://www.gnu.org/licenses/>.


Option Explicit

Enum eSource
    source
End Enum




Private Type tHSP
    H              As Single
    S              As Single
    P              As Single
End Type

Private Type tVector
    X              As Single
    Y              As Single
    L              As Single
    A              As Single
    TA             As Single
    CosA           As Single
    SinA           As Single

End Type

Private Type tLAB
    L              As Single
    A              As Single
    B              As Single
End Type


Private Type Bitmap
    bmType         As Long
    bmWidth        As Long
    bmHeight       As Long
    bmWidthBytes   As Long
    bmPlanes       As Integer
    bmBitsPixel    As Integer
    bmBits         As Long
End Type


Private Type tREGION
    cx             As Long
    cy             As Long

    NP             As Long
    X()            As Long
    Y()            As Long
    mR             As Single
    mG             As Single
    mB             As Single

End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private Sbyte()    As Byte
Private Sbyte2()   As Byte

Private BlurByte() As Byte

Private SepaByte() As Byte

'Private PGMByte() As Byte
'Private PGM_C_Byte() As Byte

Private BILAByte() As Byte
Private ContByte() As Byte
Private ContByte2() As Byte
Private COMPUTEsingle() As Single
Private COMPUTEsingle2() As Single
Private BILASingle() As Single


Private HSP()      As tHSP
Private Vec()      As tVector
Private LAB()      As tLAB


Private hBmp       As Bitmap

Private pW         As Long
Private pH         As Long
Private pB         As Long

Private Radius     As Long


Private Fast_ExpIntensity() As Single


Private Fast_ExpSpatial() As Single
Private Fast_SpatialDomain() As Single
Private Fast_SpatialDomain2() As Single

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, _
                                                                     source As Any, ByVal bytes As Long)

Public Event PercDONE(Filter As String, PercValue As Single, CurrIteration As Long)



'*************************************************************************
'BrightNess Contrast Saturation

Private Type THistSingle
    A(0 To 255)    As Single
End Type

Private histR(0 To 255) As Long
Private histG(0 To 255) As Long
Private histB(0 To 255) As Long
'*************************************************************************

'*************************************************************************
'Luminance segmentation video frames OLD MODE
'Private HistoCache(0 To 15) As THistSingle
'Private hcIDX As Long
'Private Const hcSIZE As Long = 16
'*************************************************************************

'*************************************************************************
'Luminance segmentation video frames NEW MODE
Private Type tSEG
    Start          As Single
    End            As Single
    Value          As Single
End Type
Private Type tHistoCache
    Seg()          As tSEG
End Type

Dim HistoCache(0 To 5) As tHistoCache
Private hcIDX      As Long
Private Const hcSIZE As Long = 6
'*************************************************************************

Public Sub SetUpHistoChache(Nsegm)

    Dim I          As Long
    For I = 0 To hcSIZE - 1
        ReDim HistoCache(I).Seg(1 To Nsegm)
    Next
End Sub



Private Function Fast_IntensityDomain(AbsDeltaIntensity As Single) As Single    '

    Fast_IntensityDomain = Fast_ExpIntensity(AbsDeltaIntensity)


End Function


Public Sub zInit_IntensityDomain(ByVal SigmaI, ByVal Mode)

    Dim V          As Single
    Dim V2         As Single
    Dim lV         As Long
    Dim lV2        As Long

    Dim Cos0       As Boolean

    If SigmaI = 0 Then SigmaI = 0.00001


    Select Case Mode
        Case 0
            SigmaI = 2 * SigmaI * SigmaI
        Case 1
            SigmaI = 2 * SigmaI
        Case 2
            SigmaI = SigmaI
        Case 3
            SigmaI = 3 * SigmaI
        Case 4
            SigmaI = Atn(1) * 2 * SigmaI
    End Select

    'Stop
    ReDim Fast_ExpIntensity(-25500 To 25500)

    Cos0 = False
    For V = -25500 To 25500


        V2 = Abs(V / 25500)
        Select Case Mode
            Case 0                'Gaussian
                V2 = V2 * V2
                Fast_ExpIntensity(V) = Exp(-(V2 / SigmaI))
            Case 1                'Gaussian2
                Fast_ExpIntensity(V) = (Exp(-(V2 ^ 3 / SigmaI ^ 3)))
            Case 2                'NotGaussian
                Fast_ExpIntensity(V) = Exp(-(V2 / SigmaI))
            Case 3
                Fast_ExpIntensity(V) = 1 - V2 / SigmaI
                If Fast_ExpIntensity(V) < 0 Then Fast_ExpIntensity(V) = 0

            Case 4

                If Not (Cos0) Then
                    Fast_ExpIntensity(V) = Cos(Atn(1) * V2 / SigmaI)
                    If Fast_ExpIntensity(V) < 0 Then Fast_ExpIntensity(V) = 0: Cos0 = True

                Else
                    Fast_ExpIntensity(V) = 0
                End If



        End Select


        'Fast_ExpIntensity(V) = Fast_ExpIntensity(V) * 10
    Next
    'Stop

    'ReDim Fast_IntensityDomain(0 To 2550, 0 To 2550)
    'For lV = 0 To 2550
    '    For lV2 = 0 To 2550
    '        Fast_IntensityDomain(lV, lV2) = Fast_ExpIntensity(Abs(lV - lV2)) * 100
    '    Next
    'Next

End Sub
Public Sub zInit_SpatialDomain(ByVal SigmaS)
    Dim V          As Single
    Dim V2         As Single
    Dim X          As Long
    Dim Y          As Long
    Dim D          As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * 2

    ReDim Fast_ExpSpatial(30 * 30 * 2)
    For V = 0 To 30 * 30 * 2
        V2 = V
        Fast_ExpSpatial(V) = Exp(-(V2 / (SigmaS)))
    Next


    ReDim Fast_SpatialDomain(-30 To 30, -30 To 30)
    For X = -30 To 30
        For Y = -30 To 30
            D = X * X + Y * Y
            Fast_SpatialDomain(X, Y) = Fast_ExpSpatial(D)
        Next
    Next



End Sub

Public Sub zInit_SpatialDomain2(ByVal SigmaS)
    Dim V          As Single
    Dim V2         As Single
    Dim X          As Long
    Dim Y          As Long
    Dim D          As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * 2

    ReDim Fast_ExpSpatial(30 * 30 * 2)
    For V = 0 To 30 * 30 * 2
        V2 = V
        Fast_ExpSpatial(V) = Exp(-(V2 / (SigmaS)))
    Next

    ReDim Fast_SpatialDomain2(-30 To 30, -30 To 30)
    For X = -30 To 30
        For Y = -30 To 30
            D = X * X + Y * Y
            Fast_SpatialDomain2(X, Y) = Fast_ExpSpatial(D)
        Next
    Next


End Sub
Public Sub zInit_SpatialDomain2LoG(ByVal SigmaS)
    Dim V          As Single
    Dim V2         As Single
    Dim X          As Long
    Dim Y          As Long
    Dim D          As Long

    If SigmaS = 0 Then Exit Sub
    SigmaS = SigmaS * SigmaS


    ReDim Fast_SpatialDomain2(-30 To 30, -30 To 30)
    For X = -30 To 30
        For Y = -30 To 30
            D = X * X + Y * Y
            Fast_SpatialDomain2(X, Y) = (D - 2 * SigmaS) / (SigmaS * SigmaS) * Exp(-(D / (2 * SigmaS)))
        Next
    Next


End Sub
Public Sub zSet_Source(pboxImageHandle As Long)
'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet       As Long
    'Get the bitmap header
    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)

    'Resize to hold image data
    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into Sbyte array
    'iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))


    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1



    'ReDim PGMByte(0 To PB, 0 To pW, 0 To PH)


End Sub



Private Sub Class_Terminate()
'Stop

    Erase Sbyte
    Erase Sbyte2
    Erase ContByte
    Erase BlurByte
    Erase SepaByte

    'Erase PGMByte
    'Erase PGM_C_Byte

End Sub






Public Function zLimitMin0(V) As Byte
    If V < 0 Then V = 0 Else If V > 255 Then V = 255

    zLimitMin0 = V

End Function
Public Function zLimitMax255(V As Single) As Byte
    If V > 255 Then zLimitMax255 = 255 Else: zLimitMax255 = CByte(V)
End Function

Public Sub zEFF_Contour_Apply()
    Dim X          As Long
    Dim Y          As Long

    For X = 0 + 1 To pW - 1
        For Y = 0 + 1 To pH - 1

            'If ContByte(0, X, Y) > 0 Then

            BILAByte(0, X, Y) = zLimitMin0(BILAByte(0, X, Y) \ 1 - ContByte(0, X, Y) \ 1)
            BILAByte(1, X, Y) = zLimitMin0(BILAByte(1, X, Y) \ 1 - ContByte(0, X, Y) \ 1)
            BILAByte(2, X, Y) = zLimitMin0(BILAByte(2, X, Y) \ 1 - ContByte(0, X, Y) \ 1)

            ' BILAByte(0, X, Y) = zLimitMin0(BILAByte(0, X, Y) * (1 - CSng(ContByte(0, X, Y)) / 255))
            ' BILAByte(1, X, Y) = zLimitMin0(BILAByte(1, X, Y) * (1 - CSng(ContByte(1, X, Y)) / 255))
            ' BILAByte(2, X, Y) = zLimitMin0(BILAByte(2, X, Y) * (1 - CSng(ContByte(2, X, Y)) / 255))

            'End If
        Next
    Next


End Sub
Public Sub zEFF_BilateralFilter(N As Long, Iterations As Long, RGBmode As Boolean, Optional Directional = False)
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
    Radius = N

    Const d100     As Single = 1 / 100

    Dim I          As Long

    Dim X          As Long
    Dim Y          As Long
    Dim ProgX      As Long        'For Progress Bar
    Dim ProgXStep  As Long        'For Progress Bar

    Dim XP         As Long
    Dim YP         As Long
    Dim XmN        As Long
    Dim XpN        As Long
    Dim YmN        As Long
    Dim YpN        As Long

    Dim dR         As Single
    Dim dG         As Single
    Dim dB         As Single
    Dim TR         As Long
    Dim TG         As Long
    Dim TB         As Long

    Dim RDiv       As Single
    Dim GDiv       As Single
    Dim BDiv       As Single

    Dim SpatialD   As Single

    Dim LL         As Single
    Dim AA         As Single
    Dim BB         As Single
    Dim RRR        As Single
    Dim GGG        As Single
    Dim BBB        As Single

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim XPmX       As Long

    Dim NDX        As Long
    Dim NDY        As Long

    ReDim LAB(0 To pW, 0 To pH)

    ProgXStep = Round(3 * pW / 100)
    ProgX = 0

    If RGBmode Then
        ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
        ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
        ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
        For X = 0 To pW
            For Y = 0 To pH
                COMPUTEsingle(2, X, Y) = CSng(Sbyte(2, X, Y)) * 100
                COMPUTEsingle(1, X, Y) = CSng(Sbyte(1, X, Y)) * 100
                COMPUTEsingle(0, X, Y) = CSng(Sbyte(0, X, Y)) * 100
            Next
            If X > ProgX Then
                RaiseEvent PercDONE("Bilateral INIT", X / pW, 0)
                ProgX = ProgX + ProgXStep
                DoEvents
            End If
        Next
    Else
        ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
        ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
        ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

        For X = 0 To pW
            For Y = 0 To pH
                RGB_CieLAB Sbyte(2, X, Y), Sbyte(1, X, Y), Sbyte(0, X, Y), LL, AA, BB
                COMPUTEsingle(2, X, Y) = LL * 100
                COMPUTEsingle(1, X, Y) = AA * 100
                COMPUTEsingle(0, X, Y) = BB * 100

                'for black border (LAB)
                BILASingle(2, X, Y) = 12750    '0
                BILASingle(1, X, Y) = AA * 100
                BILASingle(0, X, Y) = BB * 100

                BILAByte(1, X, Y) = AA
                BILAByte(0, X, Y) = BB
            Next
            If X > ProgX Then
                RaiseEvent PercDONE("Bilateral INIT", X / pW, 0)
                ProgX = ProgX + ProgXStep
                DoEvents
            End If
        Next

    End If



    ProgXStep = Round(2 * pW / (100 / Iterations))
    Xfrom = 0 + N
    Xto = pW - N
    Yfrom = 0 + N
    Yto = pH - N

    '-------------------------------------------------------------------------------------
    'Classic MODE
    If RGBmode Then
        For I = 1 To Iterations
            ProgX = 0
            'If Directional Then ComputeSlopes
            For X = Xfrom To Xto
                'For X = 0 To pW
                XmN = X - N
                XpN = X + N
                For Y = Yfrom To Yto
                    'For Y = 0 To PH
                    TR = 0
                    TG = 0
                    TB = 0
                    RDiv = 0
                    GDiv = 0
                    BDiv = 0
                    YmN = Y - N
                    YpN = Y + N
                    For XP = XmN To XpN
                        XPmX = XP - X
                        For YP = YmN To YpN

                            dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                            dG = Fast_ExpIntensity(COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y))
                            dB = Fast_ExpIntensity(COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y))


                            SpatialD = Fast_SpatialDomain(XPmX, YP - Y)

                            dR = dR * SpatialD
                            dG = dG * SpatialD
                            dB = dB * SpatialD

                            TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                            TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
                            TB = TB + (COMPUTEsingle(0, XP, YP)) * dB

                            RDiv = RDiv + dR
                            GDiv = GDiv + dG
                            BDiv = BDiv + dB

                        Next YP
                    Next XP

                    TR = TR / RDiv
                    TG = TG / GDiv
                    TB = TB / BDiv

                    BILASingle(2, X, Y) = TR
                    BILASingle(1, X, Y) = TG
                    BILASingle(0, X, Y) = TB

                Next Y

                ' for the progress bar
                If X > ProgX Then

                    RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
                    ProgX = ProgX + ProgXStep
                    DoEvents
                End If

            Next X

            DoEvents

            'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
            COMPUTEsingle = BILASingle

        Next I

    Else
        '------------------------------------------------------------------------------------
        'CIE LAB faster mode
        ' Compute only byte "2" that rappresent the LL Luminance of LL AA BB

        For I = 1 To Iterations
            ProgX = 0
            For X = Xfrom To Xto
                XmN = X - N
                XpN = X + N
                For Y = Yfrom To Yto
                    TR = 0
                    'TG = 0
                    'TB = 0
                    RDiv = 0
                    'GDiv = 0
                    'BDiv = 0
                    YmN = Y - N
                    YpN = Y + N
                    For XP = XmN To XpN
                        XPmX = XP - X
                        For YP = YmN To YpN

                            dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                            'dG = Fast_ExpIntensity(COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y))
                            'dB = Fast_ExpIntensity(COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y))
                            SpatialD = Fast_SpatialDomain(XPmX, YP - Y)
                            dR = dR * SpatialD
                            'dG = dG * SpatialD
                            'dB = dB * SpatialD
                            TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                            'TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
                            'TB = TB + (COMPUTEsingle(0, XP, YP)) * dB
                            RDiv = RDiv + dR
                            'GDiv = GDiv + dG
                            'BDiv = BDiv + dB
                        Next YP
                    Next XP
                    TR = TR / RDiv
                    'TG = TG / GDiv
                    'TB = TB / BDiv
                    BILASingle(2, X, Y) = TR
                    'BILASingle(1, X, Y) = TG
                    'BILASingle(0, X, Y) = TB
                Next Y

                ' for the progress bar
                If X > ProgX Then
                    RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
                    ProgX = ProgX + ProgXStep
                    DoEvents
                End If

            Next X

            DoEvents

            'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
            COMPUTEsingle = BILASingle

        Next I
        '------------------------------------------------------------------------------------
    End If


    If RGBmode Then
        For X = 0 To pW
            For Y = 0 To pH
                dR = COMPUTEsingle(2, X, Y) * d100
                dG = COMPUTEsingle(1, X, Y) * d100
                dB = COMPUTEsingle(0, X, Y) * d100
                RGB_CieLAB dR, dG, dB, LAB(X, Y).L, LAB(X, Y).A, LAB(X, Y).B

                BILAByte(2, X, Y) = dR
                BILAByte(1, X, Y) = dG
                BILAByte(0, X, Y) = dB

            Next
        Next
    Else

        For X = 0 To pW
            For Y = 0 To pH

                LL = COMPUTEsingle(2, X, Y) * d100
                AA = COMPUTEsingle(1, X, Y) * d100
                BB = COMPUTEsingle(0, X, Y) * d100
                LAB(X, Y).L = LL
                LAB(X, Y).A = AA
                LAB(X, Y).B = BB
                CieLAB_RGB LL, AA, BB, RRR, GGG, BBB

                '-----------------------------------------------------------------

                BILAByte(2, X, Y) = RRR
                BILAByte(1, X, Y) = GGG
                BILAByte(0, X, Y) = BBB
            Next
        Next

    End If

    RaiseEvent PercDONE("Bialteral", 1, Iterations)


End Sub
Public Sub zEFF_BilateralFilterDirectional(N As Long, Iterations As Long, RGBmode As Boolean)
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
    Radius = N

    Const d100     As Single = 1 / 100

    Dim I          As Long

    Dim X          As Long
    Dim Y          As Long
    Dim ProgX      As Long        'For Progress Bar
    Dim ProgXStep  As Long        'For Progress Bar

    Dim XP         As Long
    Dim YP         As Long
    Dim XmN        As Long
    Dim XpN        As Long
    Dim YmN        As Long
    Dim YpN        As Long

    Dim dR         As Single
    Dim dG         As Single
    Dim dB         As Single
    Dim TR         As Long
    Dim TG         As Long
    Dim TB         As Long

    Dim RDiv       As Single
    Dim GDiv       As Single
    Dim BDiv       As Single

    Dim SpatialD   As Single

    Dim LL         As Single
    Dim AA         As Single
    Dim BB         As Single

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim XPmX       As Long

    Dim NDX        As Long
    Dim NDY        As Long

    Dim CosA       As Single
    Dim SinA       As Single

    ReDim LAB(0 To pW, 0 To pH)

    If RGBmode Then
        ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
        ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
        ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
        For X = 0 To pW
            For Y = 0 To pH
                COMPUTEsingle(2, X, Y) = CSng(Sbyte(2, X, Y)) * 100
                COMPUTEsingle(1, X, Y) = CSng(Sbyte(1, X, Y)) * 100
                COMPUTEsingle(0, X, Y) = CSng(Sbyte(0, X, Y)) * 100
            Next
        Next
    Else
        ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
        ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
        ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

        For X = 0 To pW
            For Y = 0 To pH
                RGB_CieLAB Sbyte(2, X, Y), Sbyte(1, X, Y), Sbyte(0, X, Y), LL, AA, BB
                COMPUTEsingle(2, X, Y) = LL * 100
                COMPUTEsingle(1, X, Y) = AA * 100
                COMPUTEsingle(0, X, Y) = BB * 100

                'for black border (LAB)
                BILASingle(2, X, Y) = 12750    '0

                BILASingle(1, X, Y) = 12750
                BILASingle(0, X, Y) = 12750

                BILAByte(1, X, Y) = AA
                BILAByte(0, X, Y) = BB
            Next
        Next

    End If






    ProgXStep = Round(pW / (100 / Iterations))
    Xfrom = 0 + N * 2
    Xto = pW - N * 2
    Yfrom = 0 + N * 2
    Yto = pH - N * 2

    '-------------------------------------------------------------------------------------
    'Classic MODE
    If RGBmode Then
        For I = 1 To Iterations
            TEST2
            ProgX = 0
            'If Directional Then ComputeSlopes
            For X = Xfrom To Xto
                'For X = 0 To pW
                XmN = X - N
                XpN = X + N
                For Y = Yfrom To Yto
                    'For Y = 0 To PH
                    TR = 0
                    TG = 0
                    TB = 0
                    RDiv = 0
                    GDiv = 0
                    BDiv = 0
                    YmN = Y - N
                    YpN = Y + N
                    For XP = XmN To XpN
                        XPmX = XP - X
                        For YP = YmN To YpN

                            dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                            dG = Fast_ExpIntensity(COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y))
                            dB = Fast_ExpIntensity(COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y))


                            SpatialD = Fast_SpatialDomain(XPmX, YP - Y)

                            ' If Directional Then SpatialD = SpatialD * Vec(xP, yP).L
                            '  SpatialD = Fast_SpatialDomain((xP - X) * (N - Vec(xP, yP).X) / N, (yP - Y) * (N - Vec(xP, yP).Y) / N)

                            dR = dR * SpatialD
                            dG = dG * SpatialD
                            dB = dB * SpatialD

                            TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                            TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
                            TB = TB + (COMPUTEsingle(0, XP, YP)) * dB

                            RDiv = RDiv + dR
                            GDiv = GDiv + dG
                            BDiv = BDiv + dB

                        Next YP
                    Next XP

                    TR = TR / RDiv
                    TG = TG / GDiv
                    TB = TB / BDiv

                    ''   TR = IIf(TR < 255000, TR, 255000)
                    ''   TG = IIf(TG < 255000, TG, 255000)
                    ''   TB = IIf(TB < 255000, TB, 255000)

                    BILASingle(2, X, Y) = TR
                    BILASingle(1, X, Y) = TG
                    BILASingle(0, X, Y) = TB

                Next Y

                ' for the progress bar
                If X > ProgX Then
                    RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
                    ProgX = ProgX + ProgXStep
                    DoEvents
                End If

            Next X

            DoEvents

            'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
            COMPUTEsingle = BILASingle

        Next I

    Else
        '------------------------------------------------------------------------------------
        'CIE LAB faster mode
        ' Compute only byte "2" that rappresent the LL Luminance of LL AA BB

        For I = 1 To Iterations
            TEST2

            ProgX = 0
            'If Directional Then ComputeSlopes
            For X = Xfrom To Xto
                XmN = X - N
                XpN = X + N
                For Y = Yfrom To Yto
                    TR = 0
                    RDiv = 0
                    YmN = Y - N
                    YpN = Y + N
                    CosA = Vec(X, Y).CosA
                    SinA = Vec(X, Y).SinA
                    For XP = XmN To XpN
                        XPmX = XP - X
                        For YP = YmN To YpN



                            NDX = X + 2 * (XP - X) * CosA + 0.5 * (YP - Y) * SinA
                            NDY = Y - 2 * (XP - X) * SinA + 0.5 * (YP - Y) * CosA


                            dR = Fast_ExpIntensity(COMPUTEsingle(2, NDX, NDY) - COMPUTEsingle(2, X, Y))
                            'SpatialD = Fast_SpatialDomain(NDX - X, NDY - Y)
                            'dR = dR * SpatialD
                            TR = TR + (COMPUTEsingle(2, NDX, NDY)) * dR


                            RDiv = RDiv + dR
                        Next YP
                    Next XP

                    TR = TR / RDiv

                    BILASingle(2, X, Y) = TR
                Next Y

                ' for the progress bar
                If X > ProgX Then
                    RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
                    ProgX = ProgX + ProgXStep
                    DoEvents
                End If

            Next X

            DoEvents

            'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
            COMPUTEsingle = BILASingle

        Next I
        '------------------------------------------------------------------------------------
    End If


    If RGBmode Then
        For X = 0 To pW
            For Y = 0 To pH
                dR = COMPUTEsingle(2, X, Y) * d100
                dG = COMPUTEsingle(1, X, Y) * d100
                dB = COMPUTEsingle(0, X, Y) * d100
                RGB_CieLAB dR, dG, dB, LAB(X, Y).L, LAB(X, Y).A, LAB(X, Y).B

                BILAByte(2, X, Y) = dR
                BILAByte(1, X, Y) = dG
                BILAByte(0, X, Y) = dB

            Next
        Next
    Else

        For X = 0 To pW
            For Y = 0 To pH
                'Old Way
                '                CieLAB_RGB COMPUTEsingle(2, X, Y) * d100, COMPUTEsingle(1, X, Y) * d100, COMPUTEsingle(0, X, Y) * d100, LL, AA, BB

                'But Notice that source AA,BB can be invariant
                'it changes only LL . so this method is fater
                LAB(X, Y).L = COMPUTEsingle(2, X, Y) * d100
                LAB(X, Y).A = BILAByte(1, X, Y)
                LAB(X, Y).B = BILAByte(0, X, Y)


                CieLAB_RGB COMPUTEsingle(2, X, Y) * d100, _
                           BILAByte(1, X, Y), _
                           BILAByte(0, X, Y), LL, AA, BB


                BILAByte(2, X, Y) = LL    '(this is R)
                BILAByte(1, X, Y) = AA    '(this is G)
                BILAByte(0, X, Y) = BB    '(this is B)
            Next
        Next

    End If

    RaiseEvent PercDONE("Bialteral", 1, Iterations)


End Sub
Public Sub zEFF_BilateralFilterNOSPATIAL(N As Long, Iterations As Long, RGBmode As Boolean)
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
    Radius = N

    Const d100     As Single = 1 / 100

    Dim I          As Long

    Dim X          As Long
    Dim Y          As Long
    Dim ProgX      As Long        'For Progress Bar
    Dim ProgXStep  As Long        'For Progress Bar

    Dim XP         As Long
    Dim YP         As Long
    Dim XmN        As Long
    Dim XpN        As Long
    Dim YmN        As Long
    Dim YpN        As Long

    Dim dR         As Single
    Dim dG         As Single
    Dim dB         As Single
    Dim TR         As Long
    Dim TG         As Long
    Dim TB         As Long

    Dim RDiv       As Single
    Dim GDiv       As Single
    Dim BDiv       As Single

    Dim SpatialD   As Single

    Dim LL         As Single
    Dim AA         As Single
    Dim BB         As Single

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim XPmX       As Long

    Dim tUPr       As Single
    Dim tUPg       As Single
    Dim tUPb       As Single
    Dim tUPrDiv    As Single
    Dim tUPgDiv    As Single
    Dim tUPbDiv    As Single

    Dim tDownR     As Single
    Dim tDownG     As Single
    Dim tDownB     As Single
    Dim tDOWNrDiv  As Single
    Dim tDOWNgDiv  As Single
    Dim tDOWNbDiv  As Single

    Dim tTotR      As Single
    Dim tTotG      As Single
    Dim tTotB      As Single
    Dim tTotrDiv   As Single
    Dim tTotgDiv   As Single
    Dim tTotbDiv   As Single


    If RGBmode Then
        ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
        ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
        ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
        For X = 0 To pW
            For Y = 0 To pH
                COMPUTEsingle(2, X, Y) = CSng(Sbyte(2, X, Y)) * 100
                COMPUTEsingle(1, X, Y) = CSng(Sbyte(1, X, Y)) * 100
                COMPUTEsingle(0, X, Y) = CSng(Sbyte(0, X, Y)) * 100
            Next
        Next
    Else
        ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
        ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
        ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

        For X = 0 To pW
            For Y = 0 To pH
                RGB_CieLAB Sbyte(2, X, Y), Sbyte(1, X, Y), Sbyte(0, X, Y), LL, AA, BB
                COMPUTEsingle(2, X, Y) = LL * 100
                COMPUTEsingle(1, X, Y) = AA * 100
                COMPUTEsingle(0, X, Y) = BB * 100

                'for black border (LAB)
                BILASingle(2, X, Y) = 12750    '0

                BILASingle(1, X, Y) = 12750
                BILASingle(0, X, Y) = 12750

                BILAByte(1, X, Y) = AA
                BILAByte(0, X, Y) = BB
            Next
        Next

    End If


    ProgXStep = Round(pW / (100 / Iterations))
    Xfrom = 0 + N
    Xto = pW - N
    Yfrom = 0 + N
    Yto = pH - N

    '-------------------------------------------------------------------------------------
    'Classic MODE
    If RGBmode Then
        For I = 1 To Iterations
            ProgX = 0
            'If Directional Then ComputeSlopes
            For X = Xfrom To Xto
                'For X = 0 To pW
                XmN = X - N
                XpN = X + N
                For Y = Yfrom To Yto
                    'For Y = 0 To PH
                    TR = 0
                    TG = 0
                    TB = 0
                    RDiv = 0
                    GDiv = 0
                    BDiv = 0
                    YmN = Y - N
                    YpN = Y + N
                    For XP = XmN To XpN
                        XPmX = XP - X
                        For YP = YmN To YpN

                            dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                            dG = Fast_ExpIntensity(COMPUTEsingle(1, XP, YP) - COMPUTEsingle(1, X, Y))
                            dB = Fast_ExpIntensity(COMPUTEsingle(0, XP, YP) - COMPUTEsingle(0, X, Y))

                            TR = TR + (COMPUTEsingle(2, XP, YP)) * dR
                            TG = TG + (COMPUTEsingle(1, XP, YP)) * dG
                            TB = TB + (COMPUTEsingle(0, XP, YP)) * dB

                            RDiv = RDiv + dR
                            GDiv = GDiv + dG
                            BDiv = BDiv + dB

                        Next YP
                    Next XP

                    TR = TR / RDiv
                    TG = TG / GDiv
                    TB = TB / BDiv

                    ''   TR = IIf(TR < 255000, TR, 255000)
                    ''   TG = IIf(TG < 255000, TG, 255000)
                    ''   TB = IIf(TB < 255000, TB, 255000)

                    BILASingle(2, X, Y) = TR
                    BILASingle(1, X, Y) = TG
                    BILASingle(0, X, Y) = TB

                Next Y

                ' for the progress bar
                If X > ProgX Then
                    Stop

                    RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
                    ProgX = ProgX + ProgXStep
                    DoEvents
                End If

            Next X

            DoEvents

            'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
            COMPUTEsingle = BILASingle

        Next I

    Else
        '------------------------------------------------------------------------------------
        'CIE LAB faster mode
        ' Compute only byte "2" that rappresent the LL Luminance of LL AA BB

        For I = 1 To Iterations
            ProgX = 0


            For X = Xfrom To Xto
                XmN = X - N
                XpN = X + N

                tTotR = 0
                tTotrDiv = 0
                Y = N
                For XP = XmN To XpN
                    For YP = Y - N + 1 To Y + N - 1
                        dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                        tTotR = tTotR + (COMPUTEsingle(2, XP, YP)) * dR
                        tTotrDiv = tTotrDiv + dR

                    Next
                Next

                tDownR = 0
                tDOWNrDiv = 0
                YP = Y + N
                For XP = XmN To XpN
                    dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                    tDownR = tDownR + (COMPUTEsingle(2, XP, YP)) * dR
                    tDOWNrDiv = tDOWNrDiv + dR
                Next

                tUPr = 0
                tUPrDiv = 0
                YP = Y - N
                For XP = XmN To XpN
                    dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                    tUPr = tUPr + (COMPUTEsingle(2, XP, YP)) * dR
                    tUPrDiv = tUPrDiv + dR
                Next

                'Stop

                'tTotR = (tTotR / tTotrDiv)
                'If tUPrDiv <> 0 Then tUPr = (tUPr / tUPrDiv)
                'If tDOWNrDiv <> 0 Then tDownR = (tDownR / tDOWNrDiv)
                'tTotR = tTotR * (((N * 2 + 1) - 2) / (N * 2 + 1)) + tUPr / (N * 2 + 1) + tDownR / (N * 2 + 1)

                tTotR = (tTotR + tUPr + tDownR) / (tTotrDiv + tUPrDiv + tDOWNrDiv)


                BILASingle(2, X, Y) = tTotR

                For Y = Yfrom + 1 To Yto

                    tDownR = 0
                    tDOWNrDiv = 0
                    YP = Y + N
                    For XP = XmN To XpN
                        dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                        tDownR = tDownR + (COMPUTEsingle(2, XP, YP)) * dR
                        tDOWNrDiv = tDOWNrDiv + dR
                    Next XP

                    'If tDOWNrDiv <> 0 Then tDownR = tDownR / tDOWNrDiv
                    'tTotR = tTotR - tUPr / (N * 2 + 1) + tDownR / (N * 2 + 1)
                    tTotR = (tTotR - tUPr + tDownR) / (tTotrDiv - tUPrDiv + tDOWNrDiv)


                    BILASingle(2, X, Y) = tTotR
                    '                 If tTotR > 25500 Then Stop

                    'Stop
                    tUPr = 0
                    tUPrDiv = 0
                    YP = Y - N
                    For XP = XmN To XpN
                        dR = Fast_ExpIntensity(COMPUTEsingle(2, XP, YP) - COMPUTEsingle(2, X, Y))
                        tUPr = tUPr + (COMPUTEsingle(2, XP, YP)) * dR
                        tUPrDiv = tUPrDiv + dR
                    Next XP

                    'If tUPrDiv <> 0 Then tUPr = tUPr / tUPrDiv

                Next Y

                ' for the progress bar
                If X > ProgX Then
                    RaiseEvent PercDONE("Bilateral", (I - 1) / Iterations + (X / pW) / Iterations, I)
                    ProgX = ProgX + ProgXStep
                    DoEvents
                End If

            Next X

            DoEvents

            'CopyMemory ByVal VarPtr(COMPUTEsingle(0, 0, 0)), ByVal VarPtr(BILASingle(0, 0, 0)), (CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)) * 4    'Single Lenght=4 Bytes
            COMPUTEsingle = BILASingle

        Next I
        '------------------------------------------------------------------------------------
    End If


    If RGBmode Then
        For X = 0 To pW
            For Y = 0 To pH
                BILAByte(2, X, Y) = COMPUTEsingle(2, X, Y) * d100
                BILAByte(1, X, Y) = COMPUTEsingle(1, X, Y) * d100
                BILAByte(0, X, Y) = COMPUTEsingle(0, X, Y) * d100
            Next
        Next
    Else

        For X = 0 To pW
            For Y = 0 To pH
                'Old Way
                '                CieLAB_RGB COMPUTEsingle(2, X, Y) * d100, COMPUTEsingle(1, X, Y) * d100, COMPUTEsingle(0, X, Y) * d100, LL, AA, BB

                'But Notice that source AA,BB can be invariant
                'it changes only LL . so this method is fater
                CieLAB_RGB COMPUTEsingle(2, X, Y) * d100, _
                           BILAByte(1, X, Y), _
                           BILAByte(0, X, Y), LL, AA, BB


                BILAByte(2, X, Y) = LL    '(this is R)
                BILAByte(1, X, Y) = AA    '(this is G)
                BILAByte(0, X, Y) = BB    '(this is B)
            Next
        Next

    End If

    RaiseEvent PercDONE("Bialteral", 1, Iterations)


End Sub

Public Sub zGet_Effect(pboxImageHandle As Long)
    Dim iRet       As Long

    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, BILAByte(0, 0, 0))

    Erase BILAByte
    Erase BILASingle
    Erase COMPUTEsingle


End Sub

Public Sub zGet_Contour(pboxImageHandle As Long)
'USELESS
    Dim iRet       As Long
    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, ContByte(0, 0, 0))

    '    Erase ContByte
End Sub




Sub zPreview_Intensity(ByRef Pic As PictureBox, ByVal cSigma As Single, ByVal Mode)

    Dim X          As Single


    Dim V          As Single
    Dim ky         As Single


    Dim x1         As Single
    Dim y1         As Single
    Dim X2         As Single
    Dim Y2         As Single
    Dim KX         As Single


    zInit_IntensityDomain cSigma, Mode

    Pic.Cls


    ky = 255 / Pic.ScaleHeight

    KX = 32                       '25

    For X = 0 To KX
        V = Fast_IntensityDomain(Abs(X * 100)) * 255    '2.55

        x1 = ((Pic.ScaleWidth / KX) * X)    '* 0.1
        y1 = Pic.ScaleHeight - V / ky
        X2 = ((Pic.ScaleWidth / KX) * (X + 1))    '* 0.1
        Y2 = Pic.ScaleHeight


        Pic.Line (x1, y1)-(X2, Y2), RGB(V, 0, 0), BF
    Next


    Pic.Refresh

End Sub
Sub zPreview_Spatial(ByRef Pic As PictureBox, ByVal nn, ByVal cSigma As Single)
    Dim X          As Long
    Dim Y          As Long
    Dim C          As Integer
    Dim X2
    Dim Y2
    Dim K          As Single

    K = Pic.ScaleWidth / ((nn * 2) + 1)

    zInit_SpatialDomain cSigma
    For X = -nn To nn
        For Y = -nn To nn
            C = Fast_SpatialDomain(X, Y) * 255
            X2 = (nn + X) * K
            Y2 = (nn + Y) * K
            Pic.Line (X2, Y2)-(X2 + K, Y2 + K), RGB(C, 0, 0), BF


        Next
    Next

End Sub

Public Sub zEFF_MedianFilter(N As Long, Iterations As Long)

    Dim I          As Long

    Dim X          As Long
    Dim Y          As Long

    Dim XP         As Long
    Dim YP         As Long
    Dim XmN        As Long
    Dim XpN        As Long
    Dim YmN        As Long
    Dim YpN        As Long


    Dim TR         As Long
    Dim TG         As Long
    Dim TB         As Long



    Dim RR()       As Byte
    Dim GG()       As Byte
    Dim BB()       As Byte
    Dim T          As Byte


    Dim Area       As Long
    Dim MidP       As Long

    Dim C          As Long
    Dim CC         As Long


    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long


    Area = (N * 2 + 1) ^ 2

    ReDim RR(Area)
    ReDim GG(Area)
    ReDim BB(Area)
    MidP = Area \ 2 + 1

    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)

    For I = 1 To Iterations
        Xfrom = 0 + N
        Xto = pW - N
        For X = Xfrom To Xto

            XmN = X - N
            XpN = X + N

            Yfrom = 0 + N
            Yto = pH - N

            For Y = Yfrom To Yto

                TR = 0
                TG = 0
                TB = 0

                YmN = Y - N
                YpN = Y + N
                C = 0
                For XP = XmN To XpN
                    For YP = YmN To YpN

                        C = C + 1


                        RR(C) = Sbyte(2, XP, YP)
                        GG(C) = Sbyte(1, XP, YP)
                        BB(C) = Sbyte(0, XP, YP)

                        CC = C
                        While (CC > 0) And (RR(CC) < RR(CC - 1))
                            T = RR(CC)
                            RR(CC) = RR(CC - 1)
                            RR(CC - 1) = T
                            CC = CC - 1
                        Wend

                        CC = C
                        While (CC > 0) And (GG(CC) < GG(CC - 1))
                            T = GG(CC)
                            GG(CC) = GG(CC - 1)
                            GG(CC - 1) = T
                            CC = CC - 1
                        Wend

                        CC = C
                        While (CC > 0) And (BB(CC) < BB(CC - 1))
                            T = BB(CC)
                            BB(CC) = BB(CC - 1)
                            BB(CC - 1) = T
                            CC = CC - 1
                        Wend

                    Next
                Next


                'Debug.Print "------------------"
                '                For C = 0 To Area
                '                Debug.Print RR(C)
                '                Next
                '                Stop


                BILAByte(2, X, Y) = RR(MidP)
                BILAByte(1, X, Y) = GG(MidP)
                BILAByte(0, X, Y) = BB(MidP)


            Next
            DoEvents
        Next


        'CopyMemory ByVal VarPtr(Sbyte(0, 0, 0)), ByVal VarPtr(BILAByte(0, 0, 0)), CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)

        Sbyte = BILAByte

    Next


End Sub


Public Sub zEFF_BLUR(N As Long, Iterations As Long)

    Dim I          As Long

    Dim X          As Long
    Dim Y          As Long

    Dim XP         As Long
    Dim YP         As Long
    Dim XmN        As Long
    Dim XpN        As Long
    Dim YmN        As Long
    Dim YpN        As Long

    Dim TR         As Long
    Dim TG         As Long
    Dim TB         As Long

    Dim RR         As Long
    Dim GG         As Long
    Dim BB         As Long

    Dim Area       As Long
    Dim C          As Long

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Area = (N * 2 + 1) ^ 2

    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)

    For I = 1 To Iterations
        Xfrom = 0 + N
        Xto = pW - N
        For X = Xfrom To Xto

            XmN = X - N
            XpN = X + N

            Yfrom = 0 + N
            Yto = pH - N

            For Y = Yfrom To Yto

                TR = 0
                TG = 0
                TB = 0

                YmN = Y - N
                YpN = Y + N
                C = 0
                RR = 0
                GG = 0
                BB = 0

                For XP = XmN To XpN
                    For YP = YmN To YpN

                        RR = RR + Sbyte(2, XP, YP) \ 1
                        GG = GG + Sbyte(1, XP, YP) \ 1
                        BB = BB + Sbyte(0, XP, YP) \ 1

                    Next
                Next

                BILAByte(2, X, Y) = RR \ Area
                BILAByte(1, X, Y) = GG \ Area
                BILAByte(0, X, Y) = BB \ Area

            Next
            DoEvents
        Next


        'CopyMemory ByVal VarPtr(Sbyte(0, 0, 0)), ByVal VarPtr(BILAByte(0, 0, 0)), CLng(pB + 1) * CLng(pW + 1) * CLng(pH + 1)
        Sbyte = BILAByte

    Next


End Sub




Public Sub zEFF_ContourBySobel(Contour_0_100 As Single, LumHue01 As Single)
    Dim X          As Long
    Dim Y          As Long

    Dim ContAmount As Single
    Dim PercLUM    As Single
    Dim PercHUE    As Single

    Dim Lx         As Single
    Dim Ly         As Single
    Dim Ax         As Single
    Dim Ay         As Single
    Dim Bx         As Single
    Dim By         As Single
    Dim II         As Single

    Dim dL         As Single
    Dim dA         As Single
    Dim dB         As Single

    Dim ProgXStep  As Long
    Dim ProgX      As Long


    PercHUE = LumHue01
    PercLUM = 1 - PercHUE

    PercHUE = PercHUE * 2

    'Contour_0_100 = 25
    'ContAmount = 0.00004 * Contour_0_100
    ContAmount = 0.000075 * Contour_0_100


    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim ContByte2(0 To pB, 0 To pW, 0 To pH)
    ReDim Vec(0 To pW, 0 To pH)

    'has been put in bilateral filter
    'ReDim LAB(0 To pW, 0 To pH)
    'For X = 0 To pW
    '    For Y = 0 To pH
    '        With LAB(X, Y)
    '            RGB_CieLAB BILAByte(2, X, Y), BILAByte(1, X, Y), BILAByte(0, X, Y), .L, .A, .B
    '        End With
    '    Next
    'Next

    ProgXStep = Round(2 * pW / 100)
    ProgX = 0

    'sobel
    For X = 1 To pW - 1
        For Y = 1 To pH - 1



            Ly = (PercLUM * (-(-LAB(X - 1, Y - 1).L - 2 * LAB(X - 1, Y).L - LAB(X - 1, Y + 1).L + LAB(X + 1, Y - 1).L + 2 * LAB(X + 1, Y).L + LAB(X + 1, Y + 1).L)))
            Lx = (PercLUM * ((-LAB(X - 1, Y - 1).L - 2 * LAB(X, Y - 1).L - LAB(X + 1, Y - 1).L + LAB(X - 1, Y + 1).L + 2 * LAB(X, Y + 1).L + LAB(X + 1, Y + 1).L)))

            'If LAB(X, Y).P < 200 Then
            '    .Y = .Y + PercHUE * HUEDifference(-LAB(X - 1, Y).H, LAB(X + 1, Y).H) * 4
            '    .X = .X + PercHUE * HUEDifference(-LAB(X, Y - 1).H, LAB(X, Y + 1).H) * 4
            'End If
            If PercHUE > 0 Then
                Ay = (PercHUE * (-(-LAB(X - 1, Y - 1).A - 2 * LAB(X - 1, Y).A - LAB(X - 1, Y + 1).A + LAB(X + 1, Y - 1).A + 2 * LAB(X + 1, Y).A + LAB(X + 1, Y + 1).A)))
                Ax = (PercHUE * ((-LAB(X - 1, Y - 1).A - 2 * LAB(X, Y - 1).A - LAB(X + 1, Y - 1).A + LAB(X - 1, Y + 1).A + 2 * LAB(X, Y + 1).A + LAB(X + 1, Y + 1).A)))

                By = (PercHUE * (-(-LAB(X - 1, Y - 1).B - 2 * LAB(X - 1, Y).B - LAB(X - 1, Y + 1).B + LAB(X + 1, Y - 1).B + 2 * LAB(X + 1, Y).B + LAB(X + 1, Y + 1).B)))
                Bx = (PercHUE * ((-LAB(X - 1, Y - 1).B - 2 * LAB(X, Y - 1).B - LAB(X + 1, Y - 1).B + LAB(X - 1, Y + 1).B + 2 * LAB(X, Y + 1).B + LAB(X + 1, Y + 1).B)))
            End If

            dL = Sqr(Lx * Lx + Ly * Ly)
            dA = Sqr(Ax * Ax + Ay * Ay)
            dB = Sqr(Bx * Bx + By * By)

            II = (dL * dL + dA * dA + dB * dB) ^ 0.95
            II = II * ContAmount ^ (1 / 0.95)

            'frmMAIN.PIC2.PSet (X, Y), RGB(.L, .L, .L)
            ContByte(0, X, Y) = zLimitMax255(II)



        Next
        ' for the progress bar
        If X > ProgX Then
            RaiseEvent PercDONE("Contour", X / pW, 0)
            ProgX = ProgX + ProgXStep
            DoEvents
        End If
    Next

    RaiseEvent PercDONE("Contour", 1, 0)

End Sub
Public Sub zEFF_ContourBySobel2(Contour_0_100 As Single, LumHue01 As Single)
    Dim X          As Long
    Dim Y          As Long

    Dim ContAmount As Single
    Dim PercLUM    As Single
    Dim PercHUE    As Single

    Dim Lx         As Single
    Dim Ly         As Single
    Dim Ax         As Single
    Dim Ay         As Single
    Dim Bx         As Single
    Dim By         As Single
    Dim II         As Single

    Dim dL         As Single
    Dim dA         As Single
    Dim dB         As Single

    Dim ProgXStep  As Long
    Dim ProgX      As Long


    PercHUE = LumHue01
    PercLUM = 1 - PercHUE

    PercHUE = PercHUE * 2

    ContAmount = 0.000075 * Contour_0_100
    ContAmount = ContAmount * 4 / 12



    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim ContByte2(0 To pB, 0 To pW, 0 To pH)
    ReDim Vec(0 To pW, 0 To pH)


    ProgXStep = Round(2 * pW / 100)
    ProgX = 0

    'sobel
    For X = 2 To pW - 2
        For Y = 2 To pH - 2

            Ly = (PercLUM * (-(-2 * LAB(X - 1, Y - 1).L - 4 * LAB(X - 1, Y).L - 2 * LAB(X - 1, Y + 1).L + 2 * LAB(X + 1, Y - 1).L + 4 * LAB(X + 1, Y).L + 2 * LAB(X + 1, Y + 1).L)))
            Ly = Ly + (PercLUM * (-(-LAB(X - 2, Y - 1).L - 2 * LAB(X - 2, Y).L - LAB(X - 2, Y + 1).L + LAB(X + 2, Y - 1).L + 2 * LAB(X + 2, Y).L + LAB(X + 2, Y + 1).L)))

            Lx = (PercLUM * ((-2 * LAB(X - 1, Y - 1).L - 4 * LAB(X, Y - 1).L - 2 * LAB(X + 1, Y - 1).L + 2 * LAB(X - 1, Y + 1).L + 4 * LAB(X, Y + 1).L + 2 * LAB(X + 1, Y + 1).L)))
            Lx = Lx + (PercLUM * ((-LAB(X - 1, Y - 2).L - 2 * LAB(X, Y - 2).L - LAB(X + 1, Y - 2).L + LAB(X - 1, Y + 2).L + 2 * LAB(X, Y + 2).L + LAB(X + 1, Y + 2).L)))

            If PercHUE > 0 Then
                Ay = (PercLUM * (-(-2 * LAB(X - 1, Y - 1).A - 4 * LAB(X - 1, Y).A - 2 * LAB(X - 1, Y + 1).A + 2 * LAB(X + 1, Y - 1).A + 4 * LAB(X + 1, Y).A + 2 * LAB(X + 1, Y + 1).A)))
                Ay = Ay + (PercLUM * (-(-LAB(X - 2, Y - 1).A - 2 * LAB(X - 2, Y).A - LAB(X - 2, Y + 1).A + LAB(X + 2, Y - 1).A + 2 * LAB(X + 2, Y).A + LAB(X + 2, Y + 1).A)))
                Ax = (PercLUM * ((-2 * LAB(X - 1, Y - 1).A - 4 * LAB(X, Y - 1).A - 2 * LAB(X + 1, Y - 1).A + 2 * LAB(X - 1, Y + 1).A + 4 * LAB(X, Y + 1).A + 2 * LAB(X + 1, Y + 1).A)))
                Ax = Ax + (PercLUM * ((-LAB(X - 1, Y - 2).A - 2 * LAB(X, Y - 2).A - LAB(X + 1, Y - 2).A + LAB(X - 1, Y + 2).A + 2 * LAB(X, Y + 2).A + LAB(X + 1, Y + 2).A)))

                By = (PercLUM * (-(-2 * LAB(X - 1, Y - 1).B - 4 * LAB(X - 1, Y).B - 2 * LAB(X - 1, Y + 1).B + 2 * LAB(X + 1, Y - 1).B + 4 * LAB(X + 1, Y).B + 2 * LAB(X + 1, Y + 1).B)))
                By = By + (PercLUM * (-(-LAB(X - 2, Y - 1).B - 2 * LAB(X - 2, Y).B - LAB(X - 2, Y + 1).B + LAB(X + 2, Y - 1).B + 2 * LAB(X + 2, Y).B + LAB(X + 2, Y + 1).B)))
                Bx = (PercLUM * ((-2 * LAB(X - 1, Y - 1).B - 4 * LAB(X, Y - 1).B - 2 * LAB(X + 1, Y - 1).B + 2 * LAB(X - 1, Y + 1).B + 4 * LAB(X, Y + 1).B + 2 * LAB(X + 1, Y + 1).B)))
                Bx = Bx + (PercLUM * ((-LAB(X - 1, Y - 2).B - 2 * LAB(X, Y - 2).B - LAB(X + 1, Y - 2).B + LAB(X - 1, Y + 2).B + 2 * LAB(X, Y + 2).B + LAB(X + 1, Y + 2).B)))
            End If

            dL = Sqr(Lx * Lx + Ly * Ly)
            dA = Sqr(Ax * Ax + Ay * Ay)
            dB = Sqr(Bx * Bx + By * By)

            II = (dL * dL + dA * dA + dB * dB) ^ 0.95
            II = II * ContAmount ^ (1 / 0.95)

            'frmMAIN.PIC2.PSet (X, Y), RGB(.L, .L, .L)
            ContByte(0, X, Y) = zLimitMax255(II)

        Next
        ' for the progress bar
        If X > ProgX Then
            RaiseEvent PercDONE("Contour", X / pW, 0)
            ProgX = ProgX + ProgXStep
            DoEvents
        End If
    Next

    RaiseEvent PercDONE("Contour", 1, 0)

End Sub
Public Sub zEFF_ContourByDoG(Contour_0_100 As Single, Thresh As Single)
'by DOG
    Dim X          As Long
    Dim Y          As Long
    Dim XP         As Long
    Dim YP         As Long

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim D          As Single
    Dim Dsum       As Single

    Dim R          As Single
    Dim G          As Single
    Dim B          As Single

    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle2(0 To pB, 0 To pW, 0 To pH)

    Dim XYstep                    '=3

    frmMAIN.PIC2.Cls
    frmMAIN.PIC2.Width = (frmMAIN.PIC1.Width) * 1
    frmMAIN.PIC2.Height = (frmMAIN.PIC1.Height) * 1
    frmMAIN.PIC2.Visible = True
    frmMAIN.PIC2.Top = frmMAIN.PIC1.Top + frmMAIN.PIC1.Height    '/ 2 + 10
    frmMAIN.PIC2.Left = frmMAIN.PIC1.Left




    XYstep = 3

    zInit_SpatialDomain2 6        '8


    Dsum = 0
    For X = -XYstep To XYstep
        For Y = -XYstep To XYstep
            Dsum = Dsum + Fast_SpatialDomain2(X, Y)
        Next
    Next

    For X = XYstep To pW - XYstep
        Xfrom = X - XYstep: Xto = X + XYstep
        For Y = XYstep To pH - XYstep
            Yfrom = Y - XYstep: Yto = Y + XYstep
            R = 0
            G = 0
            B = 0

            For XP = Xfrom To Xto
                For YP = Yfrom To Yto

                    D = Fast_SpatialDomain2(X - XP, Y - YP)

                    R = R + BILAByte(2, XP, YP) * D
                    G = G + BILAByte(1, XP, YP) * D
                    B = B + BILAByte(0, XP, YP) * D
                Next
            Next
            R = R * 0.299 + G * 0.587 + B * 0.114
            R = R / Dsum
            COMPUTEsingle(2, X, Y) = R

        Next
    Next

    'XYstep = 3
    zInit_SpatialDomain2 2


    Dsum = 0
    For X = -XYstep To XYstep
        For Y = -XYstep To XYstep
            Dsum = Dsum + Fast_SpatialDomain2(X, Y)
        Next
    Next


    For X = XYstep To pW - XYstep
        Xfrom = X - XYstep: Xto = X + XYstep
        For Y = XYstep To pH - XYstep
            Yfrom = Y - XYstep: Yto = Y + XYstep
            R = 0
            G = 0
            B = 0

            For XP = Xfrom To Xto
                For YP = Yfrom To Yto

                    D = Fast_SpatialDomain2(X - XP, Y - YP)

                    R = R + BILAByte(2, XP, YP) * D
                    G = G + BILAByte(1, XP, YP) * D
                    B = B + BILAByte(0, XP, YP) * D
                Next
            Next
            R = R * 0.299 + G * 0.587 + B * 0.114
            R = R / Dsum
            COMPUTEsingle2(2, X, Y) = R

        Next
    Next



    For X = 0 To pW
        For Y = 0 To pH
            R = Contour_0_100 * (COMPUTEsingle(2, X, Y) - COMPUTEsingle2(2, X, Y) - Thresh) * 0.5
            If R < 0 Then R = 0

            R = R * 0.4
            If R > 255 Then R = 255

            ContByte(2, X, Y) = R
            ContByte(1, X, Y) = R
            ContByte(0, X, Y) = R
        Next

    Next

    zGet_Contour frmMAIN.PIC2.Image.Handle

End Sub
Public Sub zEFF_ContourByLoG(Contour_0_100 As Single, Thresh As Single)
'by Laplacian of Gaussian
    Dim X          As Long
    Dim Y          As Long
    Dim XP         As Long
    Dim YP         As Long

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim D          As Single
    Dim Dsum       As Single

    Dim R          As Single
    Dim G          As Single
    Dim B          As Single

    ReDim ContByte(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)

    Dim XYstep                    '=3

    frmMAIN.PIC2.Cls
    frmMAIN.PIC2.Width = (frmMAIN.PIC1.Width) * 1
    frmMAIN.PIC2.Height = (frmMAIN.PIC1.Height) * 1
    frmMAIN.PIC2.Visible = True
    frmMAIN.PIC2.Top = frmMAIN.PIC1.Top + frmMAIN.PIC1.Height    '/ 2 + 10
    frmMAIN.PIC2.Left = frmMAIN.PIC1.Left




    XYstep = 2

    zInit_SpatialDomain2LoG 0.75

    Dsum = 0
    For X = -XYstep To XYstep
        For Y = -XYstep To XYstep

            Dsum = Dsum + Fast_SpatialDomain2(X, Y)
            'Fast_SpatialDomain2(X, Y) = -1
        Next
    Next

    'Fast_SpatialDomain2(0, 0) = 24

    Dsum = Abs(Dsum)

    For X = XYstep To pW - XYstep
        Xfrom = X - XYstep: Xto = X + XYstep
        For Y = XYstep To pH - XYstep
            Yfrom = Y - XYstep: Yto = Y + XYstep
            R = 0
            G = 0
            B = 0

            For XP = Xfrom To Xto
                For YP = Yfrom To Yto

                    D = Fast_SpatialDomain2(X - XP, Y - YP)

                    R = R + BILAByte(2, XP, YP) * D
                    G = G + BILAByte(1, XP, YP) * D
                    B = B + BILAByte(0, XP, YP) * D
                Next
            Next
            '            Stop
            'R = (R + G + B) * 0.33333
            R = R * 0.299 + G * 0.587 + B * 0.114
            'R = R / Dsum
            COMPUTEsingle(2, X, Y) = R

        Next
    Next


    For X = 0 To pW
        For Y = 0 To pH
            R = Contour_0_100 * (COMPUTEsingle(2, X, Y))
            If R < 0 Then R = 0
            R = R * 0.0025

            If R > 255 Then R = 255
            If R < Thresh * 255 Then R = 0

            ContByte(2, X, Y) = R
            ContByte(1, X, Y) = R
            ContByte(0, X, Y) = R
        Next

    Next

    zGet_Contour frmMAIN.PIC2.Image.Handle

End Sub

Public Sub zEFF_QuantizeLuminance(Segments As Long, Presence As Single, IsThisVideo As Boolean, Optional Display As Boolean = False)

    Dim X          As Long
    Dim Y          As Long
    Dim R          As Single
    Dim G          As Single
    Dim B          As Single

    Dim Histo(0 To 255) As Double    'Long
    Dim blurHisto(0 To 255) As Double    'Long

    Dim Cache      As tHistoCache

    Dim HistoMax   As Double      'Long

    Dim Area       As Double      'Long
    Dim SegmentLimit As Double    'Long

    Dim I          As Long
    Dim prI        As Long
    Dim J          As Long
    Dim K          As Long
    Dim ws         As Double      'Long
    Dim S2         As Double      'Long

    Dim ACC        As Double      'Long

    Dim WeiSUM     As Long
    Dim StartI     As Long

    Dim NotPres    As Single

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    Dim Curseg     As Long

    ReDim Cache.Seg(1 To Segments)
    Dim S          As Single
    Dim E          As Single
    Dim V          As Single

    If Presence <= 0 Then Exit Sub
    If Presence > 1 Then Presence = 1
    NotPres = 1 - Presence

    If Segments < 2 Then Exit Sub

    RaiseEvent PercDONE("Luminance Segmentation", 0.1, 0)



    'Convert to CieLAB---------------------------------------------------------
    'has been put in bilateral filter
    'ReDim LAB(0 To pW, 0 To pH)
    'For X = 0 To pW
    '    For Y = 0 To pH
    '        With LAB(X, Y)
    '            RGB_CieLAB BILAByte(2, X, Y), BILAByte(1, X, Y), BILAByte(0, X, Y), .L, .A, .B
    '        End With
    '    Next
    'Next

    RaiseEvent PercDONE("Luminance Segmentation", 0.2, 0)

    'Fine quantization---------------------------------------------------------

    'For I = 0 To 255: Histo(I) = 0: Next
    Xfrom = Radius
    Yfrom = Radius
    Xto = pW - Radius
    Yto = pH - Radius

    For X = Xfrom To Xto
        For Y = Yfrom To Yto
            Histo(LAB(X, Y).L \ 1) = Histo(LAB(X, Y).L \ 1) + 1
        Next
    Next

    If Display Then
        For I = 0 To 255
            If Histo(I) > HistoMax Then HistoMax = Histo(I)
        Next
        frmMAIN.PIC2.Visible = True
        frmMAIN.PIC2.Cls
        frmMAIN.PIC2.Width = 255
        frmMAIN.PIC2.Height = 255
        frmMAIN.PIC2.Top = frmMAIN.PIC1.Top + frmMAIN.PIC1.Height + 10
    End If


    RaiseEvent PercDONE("Luminance Segmentation", 0.1, 0)

    Area = (pW + 1) * (pH + 1)
    SegmentLimit = Area / Segments
    'MsgBox "area " & Area & "   SegmentLimit " & SegmentLimit

    ACC = 0
    StartI = 0
    'While Histo(StartI) = 0: StartI = StartI + 1: Wend
    prI = StartI
    For I = StartI To 255
        ACC = ACC + Histo(I)
        If ACC >= SegmentLimit Then

            GoSub doOptimize

            For J = prI To I: Histo(J) = WeiSUM: Next
            prI = I + 1
            ACC = ACC - SegmentLimit
        End If
    Next

    I = 255
    GoSub doOptimize
    For J = prI To I: Histo(J) = WeiSUM: Next

    RaiseEvent PercDONE("Luminance Segmentation", 0.25, 0)

    '---------------------------------------------------------------------------
    'Video Mode OLD
    'If IsThisVideo Then
    '    For I = 0 To 255
    '        HistoCache(hcIDX Mod hcSIZE).A(I) = Histo(I)
    '    Next
    '    For I = 0 To 255
    '    ACC = 0
    '        For K = 0 To hcSIZE - 1
    '        ACC = ACC + HistoCache(K).A(I)
    '        Next
    '        ACC = ACC / hcSIZE
    '        Histo(I) = ACC
    '    Next
    '    hcIDX = hcIDX + 1
    'End If

    If IsThisVideo Then
        HistoCache(hcIDX Mod hcSIZE).Seg = Cache.Seg

        For I = 1 To Segments
            S = 0
            E = 0
            V = 0
            For K = 0 To hcSIZE - 1
                S = S + HistoCache(K).Seg(I).Start
                E = E + HistoCache(K).Seg(I).End
                V = V + HistoCache(K).Seg(I).Value
            Next
            S = S / hcSIZE
            E = E / hcSIZE
            V = V / hcSIZE
            'Stop

            For J = S To E
                Histo(J) = V
            Next
        Next
        hcIDX = hcIDX + 1
    End If


    '---------------------------------------------------------------------------
    'Blur Histo
    For I = 0 To 255: blurHisto(I) = Histo(I): Next
    For I = 1 To 255 - 1
        Histo(I) = 0.25 * (blurHisto(I - 1) + 2 * blurHisto(I) + blurHisto(I + 1))
    Next


    '---------------------------------------------------------------------------
    'Merge
    For X = 0 To pW
        For Y = 0 To pH
            LAB(X, Y).L = Presence * Histo(LAB(X, Y).L \ 1) + NotPres * LAB(X, Y).L
        Next
    Next



    '---------------------------------------------------------------------------
    'reConvert to RGB
    RaiseEvent PercDONE("Luminance Segmentation", 0.5, 0)
    For X = 0 To pW
        For Y = 0 To pH
            With LAB(X, Y)

                CieLAB_RGB .L, .A, .B, R, G, B
                BILAByte(2, X, Y) = R
                BILAByte(1, X, Y) = G
                BILAByte(0, X, Y) = B

            End With
        Next
    Next
    '--------------------------------------------------------------------------



    RaiseEvent PercDONE("Luminance Segmentation", 1, 0)


    Exit Sub


doOptimize:
    S2 = 0
    ws = 0
    For K = prI To I
        ws = ws + Histo(K) * K
        S2 = S2 + Histo(K)
    Next
    ws = ws / (S2 + 1)
    WeiSUM = ws

    If IsThisVideo Then
        Curseg = Curseg + 1
        Cache.Seg(Curseg).Start = prI
        Cache.Seg(Curseg).End = I
        Cache.Seg(Curseg).Value = WeiSUM
        If Curseg = 1 Then
            Cache.Seg(Curseg).Start = 0
            Cache.Seg(Curseg).End = I
        ElseIf Curseg = Segments Then
            Cache.Seg(Curseg).Start = prI
            Cache.Seg(Curseg).End = 255
        End If
    End If

    If Display Then
        For K = prI To I


            frmMAIN.PIC2.Line (K, 255)-(K, 255 - (255 * (Histo(K)) / HistoMax)), RGB(WeiSUM + 25, WeiSUM + 25, WeiSUM + 25)

        Next
        K = prI
        frmMAIN.PIC2.Line (K, 255)-(K, 255 - (255 * (Histo(K)) / HistoMax)), vbBlack
        frmMAIN.PIC2.Line (WeiSUM, 255)-(WeiSUM, 0), vbYellow

    End If



    Return



End Sub


Public Sub ComputeSlopes()
    Dim X          As Long
    Dim Y          As Long

    Dim ContAmount As Single
    Dim PercLUM    As Single
    Dim PercHUE    As Single



    ReDim HSP(0 To pW, 0 To pH)
    ReDim Vec(0 To pW, 0 To pH)

    For X = 0 To pW
        For Y = 0 To pH
            With HSP(X, Y)
                RGBtoHSP Sbyte(2, X, Y), Sbyte(1, X, Y), Sbyte(0, X, Y), .H, .S, .P
            End With
        Next
    Next


    For X = 1 To pW - 1
        For Y = 1 To pH - 1

            With Vec(X, Y)
                .Y = -(-HSP(X - 1, Y - 1).P - 2 * HSP(X - 1, Y).P - HSP(X - 1, Y + 1).P + HSP(X + 1, Y - 1).P + 2 * HSP(X + 1, Y).P + HSP(X + 1, Y + 1).P)
                .X = (-HSP(X - 1, Y - 1).P - 2 * HSP(X, Y - 1).P - HSP(X + 1, Y - 1).P + HSP(X - 1, Y + 1).P + 2 * HSP(X, Y + 1).P + HSP(X + 1, Y + 1).P)

                .X = Abs(.Y) / 255    ' * Val(frmMAIN.tRad)
                .Y = Abs(.X) / 255    ' * Val(frmMAIN.tRad)


                If .X > Val(frmMAIN.scrRAD) Then .X = Val(frmMAIN.scrRAD)
                If .Y > Val(frmMAIN.scrRAD) Then .Y = Val(frmMAIN.scrRAD)
                If .X > 1 Then .X = 1
                If .Y > 1 Then .Y = 1
                .X = 1 - .X
                .Y = 1 - .Y

                .L = (.X * .X + .Y * .Y)

                DoEvents
            End With
        Next
    Next

End Sub





'Author :Roberto Mior






Public Function MagneKleverBCS(ByVal BRIGHT As Long, ByVal CONTRAST As Long, ByVal SATURATION As Long)
'(c)2009 by Roy Magne Klever    www.rmklever.com
' Delphi to VB6 by reexre

    Dim CR         As Long
    Dim Cg         As Long
    Dim cB         As Long
    Dim X          As Long
    Dim Y          As Long

    Dim I          As Long
    Dim K          As Long
    Dim V          As Long
    Dim ci1        As Long
    Dim ci2        As Long
    Dim ci3        As Long

    Dim Alpha      As Long
    Dim A          As Single



    Dim ContrastLut(0 To 255) As Long
    Dim BCLut(0 To 255) As Long
    Dim SATGrays(0 To 767) As Long
    Dim SATAlpha(0 To 255) As Long


    Dim ProgX      As Long
    Dim ProgXStep  As Long


    ReDim output(0 To 2, 0 To pW, 0 To pH) As Byte


    If CONTRAST = 100 Then CONTRAST = 99
    If CONTRAST > 0 Then
        A = 1 / Cos(CONTRAST * (PI / 200))
    Else
        A = 1 * Cos(CONTRAST * (PI / 200))
    End If

    For I = 0 To 255
        V = Round(A * (I - 170) + 170)
        If V > 255 Then V = 255 Else If V < 0 Then V = 0
        ContrastLut(I) = V
    Next


    Alpha = BRIGHT
    For I = 0 To 255
        K = 256 - Alpha
        V = (K + Alpha * I) \ 256
        If V < 0 Then V = 0 Else If V > 255 Then V = 255
        BCLut(I) = ContrastLut(V)
    Next

    For I = 0 To 255
        SATAlpha(I) = (((I + 1) * SATURATION) \ 256)
    Next I

    X = 0
    For I = 0 To 255
        Y = I - SATAlpha(I)
        SATGrays(X) = Y
        X = X + 1
        SATGrays(X) = Y
        X = X + 1
        SATGrays(X) = Y
        X = X + 1
    Next


    For X = 0 To pW

        For Y = 0 To pH

            CR = Sbyte(2, X, Y)
            Cg = Sbyte(1, X, Y)
            cB = Sbyte(0, X, Y)

            V = CR + Cg + cB

            ci1 = SATGrays(V) + SATAlpha(cB)
            ci2 = SATGrays(V) + SATAlpha(Cg)
            ci3 = SATGrays(V) + SATAlpha(CR)
            If ci1 < 0 Then ci1 = 0 Else If ci1 > 255 Then ci1 = 255
            If ci2 < 0 Then ci2 = 0 Else If ci2 > 255 Then ci2 = 255
            If ci3 < 0 Then ci3 = 0 Else If ci3 > 255 Then ci3 = 255
            output(0, X, Y) = BCLut(ci1)
            output(1, X, Y) = BCLut(ci2)
            output(2, X, Y) = BCLut(ci3)

        Next
        If X > ProgX Then
            RaiseEvent PercDONE("BCS", X / pW, 0)
            ProgX = ProgX + ProgXStep
        End If


    Next

    Sbyte = output

    Erase output
    RaiseEvent PercDONE("BCS", 1, 0)
End Function


Public Function MagneKleverExposure(K As Single)
'(c)2009 by Roy Magne Klever    www.rmklever.com
' Delphi to VB6 by reexre

    Dim I          As Long
    Dim X          As Long
    Dim Y          As Long
    Dim LUT(0 To 255) As Long

    Dim ProgX      As Long
    Dim ProgXStep  As Long


    ReDim output(0 To 2, 0 To pW, 0 To pH) As Byte



    For I = 0 To 255
        If K < 0 Then
            LUT(I) = I - ((-Round((1 - Exp((I / -128) * (K / 128))) * 256) * (I Xor 255)) \ 256)
        Else
            LUT(I) = I + ((Round((1 - Exp((I / -128) * (K / 128))) * 256) * (I Xor 255)) \ 256)
        End If

        If LUT(I) < 0 Then LUT(I) = 0 Else If LUT(I) > 255 Then LUT(I) = 255
    Next

    ProgXStep = Round(3 * pW / 100)
    ProgX = 0


    For X = 0 To pW
        For Y = 0 To pH
            output(2, X, Y) = LUT(Sbyte(2, X, Y))
            output(1, X, Y) = LUT(Sbyte(1, X, Y))
            output(0, X, Y) = LUT(Sbyte(0, X, Y))
        Next
        If X > ProgX Then
            RaiseEvent PercDONE("Exposure", X / pW, 0)
            ProgX = ProgX + ProgXStep
        End If
    Next

    Sbyte = output
    Erase output

    RaiseEvent PercDONE("Exposure", 1, 0)

End Function


Public Sub MagneKleverfxHistCalc()
'(c)2009 by Roy Magne Klever    www.rmklever.com
' Delphi to VB6 by reexre

    Dim X          As Long
    Dim Y          As Long
    For X = 0 To 255
        histR(X) = 0
        histG(X) = 0
        histB(X) = 0
    Next
    For Y = 0 To pH
        For X = 0 To pW
            histR(Sbyte(2, X, Y)) = histR(Sbyte(2, X, Y)) + 1
            histG(Sbyte(1, X, Y)) = histG(Sbyte(1, X, Y)) + 1
            histB(Sbyte(0, X, Y)) = histB(Sbyte(0, X, Y)) + 1
        Next

    Next
End Sub


Private Function MagneKleverCumSum(Hist As THistSingle) As THistSingle
'(c)2009 by Roy Magne Klever    www.rmklever.com
' Delphi to VB6 by reexre
    Dim X          As Long
    Dim Temp       As THistSingle

    Temp.A(0) = Hist.A(0)
    For X = 1 To 255
        Temp.A(X) = Temp.A(X - 1) + Hist.A(X)
    Next

    MagneKleverCumSum = Temp
End Function
Public Sub MagneKleverHistogramEQU(Z As Single)
'(c)2009 by Roy Magne Klever    www.rmklever.com
' Delphi to VB6 by reexre

    Dim X          As Long
    Dim Y          As Long

    Dim Q1         As Single
    Dim Q2         As Single
    Dim Q3         As Single

    Dim Hist       As THistSingle
    Dim VCumSumR   As THistSingle
    Dim VCumSumG   As THistSingle
    Dim VCumSumB   As THistSingle

    Dim ProgX      As Long
    Dim ProgXStep  As Long


    ReDim output(0 To 2, 0 To pW, 0 To pH) As Byte

    MagneKleverfxHistCalc

    Q1 = 0                        '// RED Channel
    For X = 0 To 255
        Hist.A(X) = histR(X) ^ Z
        Q1 = Q1 + Hist.A(X)
    Next
    VCumSumR = MagneKleverCumSum(Hist)

    Q2 = 0
    For X = 0 To 255
        Hist.A(X) = histG(X) ^ Z
        Q2 = Q2 + Hist.A(X)
    Next
    VCumSumG = MagneKleverCumSum(Hist)

    Q3 = 0
    For X = 0 To 255
        Hist.A(X) = histB(X) ^ Z
        Q3 = Q3 + Hist.A(X)
    Next
    VCumSumB = MagneKleverCumSum(Hist)


    ProgXStep = Round(3 * pW / 100)
    ProgX = 0



    For X = 0 To pW
        For Y = 0 To pH


            output(2, X, Y) = Fix((255 / Q1) * VCumSumR.A(Sbyte(2, X, Y)))
            output(1, X, Y) = Fix((255 / Q2) * VCumSumG.A(Sbyte(1, X, Y)))
            output(0, X, Y) = Fix((255 / Q3) * VCumSumB.A(Sbyte(0, X, Y)))

            '            RGB[x].B := Trunc((255 / q3) * vcumsumB[RGB[x].B]);
        Next
        If X > ProgX Then
            RaiseEvent PercDONE("Equalize", X / pW, 0)
            ProgX = ProgX + ProgXStep
        End If
    Next
    Sbyte = output
    Erase output
    RaiseEvent PercDONE("Equalize", 1, 0)

End Sub





Public Sub TEST(pboxImageHandle As Long, pboxImageHandle2 As Long)
    Dim X          As Long
    Dim Y          As Long
    Dim xX         As Long
    Dim yY         As Long

    Dim cx         As Long
    Dim cy         As Long

    Dim cXI        As Long
    Dim cYI        As Long

    Dim P00        As Single
    Dim P01        As Single
    Dim P10        As Single
    Dim P11        As Single


    Dim x00        As Single
    Dim x01        As Single
    Dim x10        As Single
    Dim x11        As Single

    Dim y00        As Single
    Dim y01        As Single
    Dim y10        As Single
    Dim y11        As Single

    Dim pDX        As Single
    Dim pDY        As Single

    Dim C20        As Single
    Dim C02        As Single

    Dim C30        As Single
    Dim C03        As Single

    Dim C21        As Single
    Dim C31        As Single
    Dim C12        As Single
    Dim C13        As Single
    Dim C11        As Single

    Dim Csum1      As Single
    Dim Csum2      As Single

    Dim U          As Single
    Dim U2         As Single
    Dim U3         As Single

    Dim V          As Single
    Dim V2         As Single
    Dim V3         As Single

    Dim Pressure
    Dim Density

    Dim iRet       As Long
    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)

    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))


    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1


    ReDim mul(0 To pB, (pW + 1) * 4, (pH + 1) * 4) As Byte
    ReDim out(0 To pB, (pW + 1) * 4, (pH + 1) * 4) As Byte


    For X = 0 To pW - 1
        For Y = 0 To pH - 1

            For xX = X * 4 To (X + 1) * 4
                For yY = Y * 4 To (Y + 1) * 4

                    mul(2, xX, yY) = Sbyte(2, X, Y)
                    mul(1, xX, yY) = Sbyte(1, X, Y)
                    mul(0, xX, yY) = Sbyte(0, X, Y)


                Next
            Next
        Next
    Next

    Stop

    For X = 2 To pW * 4 - 4
        For Y = 2 To pH * 4 - 4


            cx = (X \ 4) * 4
            cy = (Y \ 4) * 4


            cXI = ((X + 4) \ 4) * 4
            cYI = ((Y + 4) \ 4) * 4



            '          Stop

            P00 = mul(2, cx, cy) / 255 / 4
            x00 = cx
            y00 = cy

            P01 = mul(2, cx, cYI) / 255 / 4
            x01 = cx
            y01 = cYI

            P10 = mul(2, cXI, cy) / 255 / 4
            x10 = cXI
            y10 = cy

            P11 = mul(2, cXI, cYI) / 255 / 4
            x11 = cXI
            y11 = cYI

            pDX = P10 - P00
            pDY = P01 - P00

            C20 = 3 * pDX - x10 - 2 * x00
            C02 = 3 * pDY - y01 - 2 * y00
            C30 = -2 * pDX + x10 + x00
            C03 = -2 * pDY + y01 + y00

            Csum1 = P00 + y00 + C02 + C03
            Csum2 = P00 + x00 + C20 + C30

            C21 = 3 * P11 - 2 * x01 - x11 - 3 * Csum1 - C20
            C31 = (-2 * P11 + x01 + x11 + 2 * Csum1) - C30
            C12 = 3 * P11 - 2 * y10 - y11 - 3 * Csum2 - C02
            C13 = (-2 * P11 + y10 + y11 + 2 * Csum2) - C03
            C11 = x01 - C13 - C12 - x00

            U = (X - cx)
            U2 = U * U
            U3 = U * U2
            V = (Y - cy)
            V2 = V * V
            V3 = V * V2

            Density = P00 + x00 * U + y00 * V + _
                      C20 * U2 + C02 * V2 + _
                      C30 * U3 + C03 * V3 + _
                      C21 * U2 * V + C31 * U3 * V + _
                      C12 * U * V2 + C13 * U * V3 + C11 * U * V


            'Pressure = (Density - 1)

            Pressure = Density

            If Pressure < 0 Then Pressure = 0
            If Pressure > 255 Then Pressure = 255
            out(2, X, Y) = Pressure \ 1

        Next
    Next


    iRet = SetBitmapBits(pboxImageHandle2, 4 * CLng(pW * 4) * CLng(pH * 4), out(0, 0, 0))




End Sub

Public Sub IWH(pboxImageHandle As Long, pboxImageHandle2 As Long)

    Dim I          As Long
    Dim K          As Long


    Dim X          As Long
    Dim Y          As Long
    Dim dX         As Long
    Dim dY         As Long

    Dim X2         As Long
    Dim Y2         As Long

    Dim pX         As Long
    Dim pY         As Long

    Dim R          As Single
    Dim G          As Single
    Dim B          As Single

    Dim MinI       As Single
    Dim MaxI       As Single
    Dim Inte       As Single

    Dim CR(-4 To 4) As Long
    Dim Cg(-4 To 4) As Long
    Dim cB(-4 To 4) As Long

    Dim VIx        As Long
    Dim VIy        As Long
    Dim HIx        As Long
    Dim HIy        As Long

    Dim FOUND      As Boolean


    Dim DC(-4 To 4) As Long
    Dim S          As String

    Dim IT         As Long
    Dim C          As Long

    Dim NR         As Long
    Dim CurR       As Long
    Dim Reg()      As tREGION


    Dim iRet       As Long
    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)

    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))


    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1


    ReDim out(0 To pB, pW, pH) As Byte
    ReDim rnds(0 To pB, pW, pH) As Byte


    ReDim blur(0 To pB, 0 To pW, 0 To pH) As Single

    ReDim XdirV(0 To 1, 0 To pW, 0 To pH) As Long
    ReDim YdirV(0 To 1, 0 To pW, 0 To pH) As Long
    ReDim XdirH(0 To 1, 0 To pW, 0 To pH) As Long
    ReDim YdirH(0 To 1, 0 To pW, 0 To pH) As Long


    ReDim LUM(0 To pW, 0 To pH) As Single

    ReDim regC(0 To pB, 0 To pW, 0 To pH) As Single


    ReDim vallx(0 To pW, 0 To pH) As Long
    ReDim vally(0 To pW, 0 To pH) As Long

    ReDim hillx(0 To pW, 0 To pH) As Long
    ReDim hilly(0 To pW, 0 To pH) As Long


    ReDim BILAByte(0 To pB, 0 To pW, 0 To pH)
    ReDim BILASingle(0 To pB, 0 To pW, 0 To pH)
    ReDim COMPUTEsingle(0 To pB, 0 To pW, 0 To pH)



    For X = 0 To pW
        For Y = 0 To pH
            COMPUTEsingle(2, X, Y) = CSng(Sbyte(2, X, Y)) * 100
            COMPUTEsingle(1, X, Y) = CSng(Sbyte(1, X, Y)) * 100
            COMPUTEsingle(0, X, Y) = CSng(Sbyte(0, X, Y)) * 100
            rnds(2, X, Y) = Rnd * 255
            rnds(1, X, Y) = Rnd * 255
            rnds(0, X, Y) = Rnd * 255

        Next
    Next


    '*******************************************************************
    'PreBlur
    For IT = 1 To 30
        NR = 0
        ReDim Reg(0)


        For X = 1 To pW - 1
            For Y = 1 To pH - 1
                R = 0
                G = 0
                B = 0

                For pX = X - 1 To X + 1
                    For pY = Y - 1 To Y + 1

                        R = R + COMPUTEsingle(2, pX, pY)
                        G = G + COMPUTEsingle(1, pX, pY)
                        B = B + COMPUTEsingle(0, pX, pY)
                        If pX - X = 0 And pY - Y = 0 Then
                            R = R + COMPUTEsingle(2, pX, pY) * 5
                            G = G + COMPUTEsingle(1, pX, pY) * 5
                            B = B + COMPUTEsingle(0, pX, pY) * 5
                        End If

                    Next
                Next
                R = R / (9 + 5)
                G = G / (9 + 5)
                B = B / (9 + 5)

                blur(2, X, Y) = R
                blur(1, X, Y) = G
                blur(0, X, Y) = B

                '            lum(X, Y) = GetLuminance(R * 0.01, _
                             G * 0.01, _
                             B * 0.01)
                LUM(X, Y) = (R * 0.241) ^ 2 + (G * 0.691) ^ 2 + (B * 0.068) ^ 2

            Next
        Next
        COMPUTEsingle = blur

        '*****************************************************************************

        '******************************************************************************
        'directions
        For X = 1 To pW - 1
            For Y = 1 To pH - 1

                Inte = LUM(X, Y)
                MinI = Inte
                MaxI = Inte

                XdirV(0, X, Y) = 0
                YdirV(0, X, Y) = 0

                XdirH(0, X, Y) = 0
                YdirH(0, X, Y) = 0

                For pX = X - 1 To X + 1
                    For pY = Y - 1 To Y + 1

                        Inte = LUM(pX, pY)

                        If Inte < MinI Then
                            MinI = Inte
                            XdirV(0, X, Y) = pX - X
                            YdirV(0, X, Y) = pY - Y
                        End If
                        If Inte > MaxI Then
                            MaxI = Inte
                            XdirH(0, X, Y) = pX - X
                            YdirH(0, X, Y) = pY - Y
                        End If

                    Next
                Next



                '            If Abs(XdirV(0, X, Y)) + Abs(YdirV(0, X, Y)) < 2 Then Stop
            Next


        Next
        '********************************************************************************


        For X = 0 To pW


            For Y = 0 To pH


                dX = XdirV(0, X, Y)
                dY = YdirV(0, X, Y)

                X2 = X
                Y2 = Y

                Do

                    X2 = X2 + dX
                    Y2 = Y2 + dY
                    dX = XdirV(0, X2, Y2)
                    dY = YdirV(0, X2, Y2)

                Loop While Not (dX = 0 And dY = 0)

                vallx(X, Y) = X2
                vally(X, Y) = Y2


                'FIND REGIONS



                FOUND = False
                For I = 1 To NR
                    If (Reg(I).cx = X2) Then If (Reg(I).cy = Y2) Then FOUND = True:: CurR = I: Exit For
                Next
                'Stop


                If Not (FOUND) Then
                    NR = NR + 1
                    ReDim Preserve Reg(NR)
                    CurR = NR
                End If

                With Reg(CurR)
                    .cx = X2
                    .cy = Y2
                    .NP = .NP + 1
                    ReDim Preserve .X(.NP)
                    ReDim Preserve .Y(.NP)
                    .X(.NP) = X
                    .Y(.NP) = Y
                End With


                '------------------------------------------

                '---------------------------------------------------

                dX = XdirH(0, X, Y)
                dY = YdirH(0, X, Y)

                X2 = X
                Y2 = Y
                Do
                    X2 = X2 + dX
                    Y2 = Y2 + dY
                    dX = XdirH(0, X2, Y2)
                    dY = YdirH(0, X2, Y2)

                Loop While Not (dX = 0 And dY = 0)

                hillx(X, Y) = X2
                hilly(X, Y) = Y2
            Next
        Next
        'Stop

        '--------------------------------------------------
        'Avarage regions
        For I = 1 To NR
            With Reg(I)
                R = 0
                G = 0
                B = 0
                For K = 1 To .NP
                    R = R + blur(2, .X(K), .Y(K))
                    G = G + blur(1, .X(K), .Y(K))
                    B = B + blur(0, .X(K), .Y(K))
                Next
                R = R / .NP
                G = G / .NP
                B = B / .NP
                R = R * 0.01
                G = G * 0.01
                B = B * 0.01

                For K = 1 To .NP
                    regC(2, .X(K), .Y(K)) = R
                    regC(1, .X(K), .Y(K)) = G
                    regC(0, .X(K), .Y(K)) = B
                Next


            End With
        Next

        '-------------------------------------------------------

        ' GoTo skip:

        '-----------------------------------------------
        For I = -4 To 4
            CR(I) = 100 + Rnd * 155
            Cg(I) = 100 + Rnd * 155
            cB(I) = 100 + Rnd * 155

            If I = 0 Then
                CR(I) = 0
                Cg(I) = 0
                cB(I) = 0
            End If
        Next


        For X = 0 To pW
            For Y = 0 To pH
                'Stop
                'If XdirV(0, X, Y) <> 0 Then Stop
                I = XdirV(0, X, Y) * 3 + YdirV(0, X, Y) * 1
                DC(I) = DC(I) + 1

                out(2, X, Y) = CR(I)
                out(1, X, Y) = Cg(I)
                out(0, X, Y) = cB(I)
                'GoTo nn
                VIx = vallx(X, Y)
                VIy = vally(X, Y)

                HIx = hillx(X, Y)
                HIy = hilly(X, Y)

                'out(2, X, Y) = (COMPUTEsingle(2, VIx, VIy) + COMPUTEsingle(2, HIx, HIy)) * 0.005
                'out(1, X, Y) = (COMPUTEsingle(1, VIx, VIy) + COMPUTEsingle(1, HIx, HIy)) * 0.005
                'out(0, X, Y) = (COMPUTEsingle(0, VIx, VIy) + COMPUTEsingle(0, HIx, HIy)) * 0.005

                'out(2, X, Y) = (COMPUTEsingle(2, VIx, VIy) * 0.01)
                'out(1, X, Y) = (COMPUTEsingle(1, VIx, VIy) * 0.01)
                'out(0, X, Y) = (COMPUTEsingle(0, VIx, VIy) * 0.01)
                GoTo nn
                If I <> 0 Then
                    out(2, X, Y) = (rnds(2, VIx, VIy))
                    out(1, X, Y) = (rnds(1, VIx, VIy))
                    out(0, X, Y) = (rnds(0, VIx, VIy))
                Else
                    out(2, X, Y) = 0
                    out(1, X, Y) = 0
                    out(0, X, Y) = 0
                End If

nn:

                out(2, X, Y) = regC(2, X, Y)
                out(1, X, Y) = regC(1, X, Y)
                out(0, X, Y) = regC(0, X, Y)

            Next

        Next





skip:
        '------------------------------------------------------------------------------------
        iRet = SetBitmapBits(pboxImageHandle2, 4 * CLng(pW) * CLng(pH), out(0, 0, 0))
        frmMAIN.PIC2.Refresh


        S = "Iter " & IT & vbCrLf & vbCrLf
        S = S + "Regions " & NR & vbCrLf & vbCrLf

        For I = -4 To 4
            S = S + "Dir " & I & "   " & DC(I) & vbCrLf
        Next
        MsgBox S

    Next IT
End Sub

Public Function Atan2(ByVal dX As Single, ByVal dY As Single) As Single
'This Should return Angle

    Dim Theta      As Single

    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            Theta = 0#
        ElseIf (dY > 0#) Then
            Theta = 1.5707963267949
            'theta = PI / 2
        Else
            Theta = -1.5707963267949
            'theta = -PI / 2
        End If
    Else
        Theta = Atn(dY / dX)

        If (dX < 0) Then
            If (dY >= 0#) Then
                Theta = PI + Theta
            Else
                Theta = Theta - PI
            End If
        End If
    End If


    Atan2 = Theta

    While Atan2 < 0: Atan2 = Atan2 + PI: Wend
    While Atan2 > PI: Atan2 = Atan2 - PI: Wend


End Function
Public Sub TEST2()
    Dim iRet       As Long
    Dim X          As Long
    Dim Y          As Long
    Dim XP         As Long
    Dim YP         As Long
    Dim I          As Long


    Dim Lx         As Single
    Dim Ly         As Single
    Dim Lx2        As Single
    Dim Ly2        As Single
    Dim Lx3        As Single
    Dim Ly3        As Single



    Dim Lsum       As Single

    Dim A          As Single

    Dim SL         As Single

    '    iRet = GetObject(pboxImageHandle, Len(hBmp), hBmp)
    '    '   iRet = GetObject(pBoxPicHand, Len(hBmp), hBmp)
    '
    '    'Resize to hold image data
    '    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    '    'Get the image data and store into Sbyte array
    '    'iRet = GetBitmapBits(pBox.Picture.Handle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
    '    iRet = GetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))
    '    pW = hBmp.bmWidth - 1
    '    pH = hBmp.bmHeight - 1
    '    pB = (hBmp.bmBitsPixel \ 8) - 1


    ReDim Vec(0 To pW, 0 To pH)
    ReDim vec2(0 To pW, 0 To pH) As tVector


    ReDim LAB(0 To pW, 0 To pH)

    For X = 0 To pW
        For Y = 0 To pH

            'RGB_CieLAB Sbyte(2, X, Y), Sbyte(1, X, Y), Sbyte(0, X, Y), _
             LAB(X, Y).L, LAB(X, Y).A, LAB(X, Y).B

            LAB(X, Y).L = COMPUTEsingle(2, X, Y)

        Next
    Next

    For X = 1 To pW - 1
        '    Stop
        For Y = 1 To pH - 1


            Ly = -(-LAB(X - 1, Y - 1).L - 2 * LAB(X - 1, Y).L - LAB(X - 1, Y + 1).L + LAB(X + 1, Y - 1).L + 2 * LAB(X + 1, Y).L + LAB(X + 1, Y + 1).L)
            Lx = (-LAB(X - 1, Y - 1).L - 2 * LAB(X, Y - 1).L - LAB(X + 1, Y - 1).L + LAB(X - 1, Y + 1).L + 2 * LAB(X, Y + 1).L + LAB(X + 1, Y + 1).L)

            Ly2 = -(-LAB(X - 1, Y - 1).A - 2 * LAB(X - 1, Y).A - LAB(X - 1, Y + 1).A + LAB(X + 1, Y - 1).A + 2 * LAB(X + 1, Y).A + LAB(X + 1, Y + 1).A)
            Lx2 = (-LAB(X - 1, Y - 1).A - 2 * LAB(X, Y - 1).A - LAB(X + 1, Y - 1).A + LAB(X - 1, Y + 1).A + 2 * LAB(X, Y + 1).A + LAB(X + 1, Y + 1).A)

            Ly3 = -(-LAB(X - 1, Y - 1).B - 2 * LAB(X - 1, Y).B - LAB(X - 1, Y + 1).B + LAB(X + 1, Y - 1).B + 2 * LAB(X + 1, Y).B + LAB(X + 1, Y + 1).B)
            Lx3 = (-LAB(X - 1, Y - 1).B - 2 * LAB(X, Y - 1).B - LAB(X + 1, Y - 1).B + LAB(X - 1, Y + 1).B + 2 * LAB(X, Y + 1).B + LAB(X + 1, Y + 1).B)


            Lx = Lx + Lx2 + Lx3
            Ly = Ly + Ly2 + Ly3

            Vec(X, Y).X = (Lx)
            Vec(X, Y).Y = (Ly)

            Vec(X, Y).L = Sqr(Lx * Lx + Ly * Ly)

            Vec(X, Y).A = Atan2(Lx, Ly)

            Vec(X, Y).CosA = Cos(Vec(X, Y).A)
            Vec(X, Y).SinA = Sin(Vec(X, Y).A)

        Next
    Next

    'GoTo skip

    For I = 1 To 1

        vec2 = Vec

        'GoTo skip
        For X = 2 To pW - 2
            For Y = 2 To pH - 2
                Vec(X, Y).X = 0
                Vec(X, Y).Y = 0
                Lx = 0
                Ly = 0
                Lsum = 0
                A = 0
                For XP = X - 2 To X + 2
                    For YP = Y - 2 To Y + 2

                        Lx = Lx + Vec(XP, YP).X
                        Ly = Ly + Vec(XP, YP).Y
                        Lsum = Lsum + 1


                    Next
                Next

                If Lsum <> 0 Then
                    Vec(X, Y).X = Lx / Lsum
                    Vec(X, Y).Y = Ly / Lsum
                End If
                Vec(X, Y).L = Sqr(Vec(X, Y).X * Vec(X, Y).X + Vec(X, Y).Y * Vec(X, Y).Y)
                Vec(X, Y).A = Atan2(Vec(X, Y).X, Vec(X, Y).Y)

                Vec(X, Y).CosA = Cos(A)
                Vec(X, Y).SinA = Sin(A)


            Next
        Next
    Next

skip:
    For X = 1 To pW - 1 Step 1
        For Y = 1 To pH - 1 Step 1
            frmMAIN.PIC2.Line (X * 5, Y * 5)-(X * 5 + Cos(Vec(X, Y).A) * 4, Y * 5 + Sin(Vec(X, Y).A) * 4), RGB(Sqr(Vec(X, Y).L) * 80, 0, 0)
        Next
    Next



End Sub

