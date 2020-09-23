VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSketch"
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

Private Type Bitmap
    bmType         As Long
    bmWidth        As Long
    bmHeight       As Long
    bmWidthBytes   As Long
    bmPlanes       As Integer
    bmBitsPixel    As Integer
    bmBits         As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private hBmp       As Bitmap

Private pW         As Integer
Private pH         As Integer
Private pB         As Integer


Private Const PI   As Single = 3.14159265358979


Private Sbyte()    As Byte
Private ResSingle() As Byte
Private ANGULAR()  As Single

Private GaborFilter(-3 To 3, -3 To 3, 0 To 15) As Single


Public Event PercDONE(PercValue As Single, CurrIteration As Long)

Private Function max(A, B)
    max = IIf(A > B, A, B)

End Function

Private Function ceil(A) As Integer
    Dim B          As Single
    B = Int(A)

    If B <> A Then
        ceil = B + 1
    Else

        ceil = B
    End If


End Function
Public Sub SetSource(PicImageHandle As Long)
'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet       As Long
    'Get the bitmap header
    iRet = GetObject(PicImageHandle, Len(hBmp), hBmp)


    'Resize to hold image data
    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into Sbyte array
    iRet = GetBitmapBits(PicImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))

    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1


    ReDim ResSingle(0 To pB, 0 To pW, 0 To pH)

    ReDim ANGULAR(0 To pB, 0 To pW, 0 To pH, 0 To 15)


End Sub
Public Sub SetSourceToMIX(PicImageHandle As Long)
'Public Sub GetBits(pBoxPicHand As Long)
    Dim iRet       As Long
    'Get the bitmap header
    iRet = GetObject(PicImageHandle, Len(hBmp), hBmp)


    'Resize to hold image data
    ReDim Sbyte(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into Sbyte array
    iRet = GetBitmapBits(PicImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, Sbyte(0, 0, 0))

    pW = hBmp.bmWidth - 1
    pH = hBmp.bmHeight - 1
    pB = (hBmp.bmBitsPixel \ 8) - 1


End Sub

Public Sub MIX(Amount)
    Dim R          As Single
    Dim G          As Single
    Dim B          As Single
    Dim X          As Long
    Dim Y          As Long

    Amount = Amount / 100

    For X = 0 To pW
        For Y = 0 To pH
            R = Sbyte(2, X, Y) - (200 - (ResSingle(2, X, Y))) * Amount
            G = Sbyte(1, X, Y) - (200 - (ResSingle(1, X, Y))) * Amount
            B = Sbyte(0, X, Y) - (200 - (ResSingle(0, X, Y))) * Amount
            If R < 0 Then R = 0 Else If R > 255 Then R = 255
            If G < 0 Then G = 0 Else If G > 255 Then G = 255
            If B < 0 Then B = 0 Else If B > 255 Then B = 255


            ResSingle(2, X, Y) = R
            ResSingle(1, X, Y) = G
            ResSingle(0, X, Y) = B
        Next
    Next

End Sub

Public Sub GetEffect(pboxImageHandle As Long)
    Dim iRet       As Long

    iRet = SetBitmapBits(pboxImageHandle, hBmp.bmWidthBytes * hBmp.bmHeight, ResSingle(0, 0, 0))


End Sub

'FROM WIKI:
'function gb=gabor_fn(sigma,theta,lambda,psi,gamma)
'
'sigma_x = sigma;
'sigma_y = sigma/gamma;
'
'% Bounding box
'nstds = 3;
'xmax = max(abs(nstds*sigma_x*cos(theta)),abs(nstds*sigma_y*sin(theta)));
'xmax = ceil(max(1,xmax));
'ymax = max(abs(nstds*sigma_x*sin(theta)),abs(nstds*sigma_y*cos(theta)));
'ymax = ceil(max(1,ymax));
'xmin = -xmax; ymin = -ymax;
'[x,y] = meshgrid(xmin:xmax,ymin:ymax);
'
'% Rotation
'x_theta=x*cos(theta)+y*sin(theta);
'y_theta=-x*sin(theta)+y*cos(theta);
'
'gb=exp(-.5*(x_theta.^2/sigma_x^2+y_theta.^2/sigma_y^2)).*cos(2*pi/lambda*x_theta+psi);
Private Sub InitGaborFilter(Sigma, Lambda, Psi, Gamma)
    Dim SigmaX     As Single
    Dim SigmaY     As Single
    Dim Xmax       As Single
    Dim Ymax       As Single
    Dim Xmin       As Single
    Dim Ymin       As Single
    Dim Xtheta     As Single
    Dim Ytheta     As Single
    Dim nstds      As Long
    Dim X          As Long
    Dim Y          As Long
    Dim GB         As Single
    Dim Theta      As Single
    Dim CC         As Single
    Dim A          As Long


    SigmaX = Sigma
    SigmaY = Sigma / Gamma
    'Bounding box
    nstds = 3

    A = 0
    '    For Theta = 0 To PI Step (PI / 16)
    For Theta = 0 To PI * 2 Step (2 * PI / 16)


        Xmax = max(Abs(nstds * SigmaX * Cos(Theta)), Abs(nstds * SigmaY * Sin(Theta)))
        Xmax = ceil(max(1, Xmax))
        Ymax = max(Abs(nstds * SigmaX * Sin(Theta)), Abs(nstds * SigmaY * Cos(Theta)))
        Ymax = ceil(max(1, Ymax))
        Xmin = -Xmax
        Ymin = -Ymax

        ' Rotation
        'frmMAIN.Pic.Cls
        For X = -nstds To nstds
            For Y = -nstds To nstds
                Xtheta = X * Cos(Theta) + Y * Sin(Theta)
                Ytheta = -X * Sin(Theta) + Y * Cos(Theta)
                '
                GB = Exp(-0.5 * (Xtheta ^ 2 / SigmaX ^ 2 + Ytheta ^ 2 / SigmaY ^ 2)) * Cos(2 * PI / Lambda * Xtheta + Psi)

                GaborFilter(X, Y, A) = GB

                CC = 0
                If GB < 0 Then CC = -GB: GB = 0

                'frmMAIN.Pic.Line (100 + X * 10, 100 + Y * 10)-(100 + (X + 1) * 10, 100 + (Y + 1) * 10), RGB(GB * 255, CC * 255, 0), BF


            Next
        Next
        'Stop

        A = A + 1

    Next

    For A = 0 To 15
        CC = 0
        For X = -nstds To nstds
            For Y = -nstds To nstds
                CC = CC + GaborFilter(X, Y, A)
            Next
        Next
        CC = CC / 49
        For X = -nstds To nstds
            For Y = -nstds To nstds
                GaborFilter(X, Y, A) = GaborFilter(X, Y, A) - CC
            Next
        Next
    Next

End Sub

Public Function Sketch()



'    InitGaborFilter 2, 5, 0, 0.25

    InitGaborFilter 1, 5, 2, 0.3
    ApplyGaborFilter


End Function

Private Sub ApplyGaborFilter()
    Dim A          As Long
    Dim XP         As Long
    Dim YP         As Long
    Dim MaxR       As Single
    Dim MaxG       As Single
    Dim MaxB       As Single

    Dim X          As Long
    Dim Y          As Long
    Dim ProgX      As Long        'For Progress Bar
    Dim ProgXStep  As Long        'For Progress Bar

    Dim Xfrom      As Long
    Dim Xto        As Long
    Dim Yfrom      As Long
    Dim Yto        As Long

    'ProgXstep = Round(pW / (100 / Iterations))
    ProgXStep = Round(2 * pW / 100)
    ProgX = 0

    For X = 3 To pW - 4

        Xfrom = X - 3
        Xto = X + 3

        For Y = 3 To pH - 4

            ResSingle(0, X, Y) = 0
            ResSingle(0, X, Y) = 0
            ResSingle(0, X, Y) = 0

            For A = 0 To 15
                ANGULAR(0, X, Y, A) = 0
                ANGULAR(1, X, Y, A) = 0
                ANGULAR(2, X, Y, A) = 0


                Yfrom = Y - 3
                Yto = Y + 3

                For XP = Xfrom To Xto
                    For YP = Yfrom To Yto


                        ANGULAR(0, X, Y, A) = ANGULAR(0, X, Y, A) + GaborFilter(X - XP, Y - YP, A) * Sbyte(0, XP, YP)
                        ANGULAR(1, X, Y, A) = ANGULAR(1, X, Y, A) + GaborFilter(X - XP, Y - YP, A) * Sbyte(1, XP, YP)
                        ANGULAR(2, X, Y, A) = ANGULAR(2, X, Y, A) + GaborFilter(X - XP, Y - YP, A) * Sbyte(2, XP, YP)

                    Next YP
                Next XP
            Next A


            MaxR = -9999999
            MaxG = -9999999
            MaxB = -9999999

            For A = 0 To 15
                If ANGULAR(2, X, Y, A) > MaxR Then MaxR = ANGULAR(2, X, Y, A)
                If ANGULAR(1, X, Y, A) > MaxG Then MaxG = ANGULAR(1, X, Y, A)
                If ANGULAR(0, X, Y, A) > MaxB Then MaxB = ANGULAR(0, X, Y, A)
            Next

            '            MaxR = MaxR * 0.5
            '            MaxG = MaxG * 0.5
            '            MaxB = MaxB * 0.5


            If MaxR < 0 Then MaxR = 0 Else If MaxR > 255 Then MaxR = 255
            If MaxG < 0 Then MaxG = 0 Else If MaxG > 255 Then MaxG = 255
            If MaxB < 0 Then MaxB = 0 Else If MaxB > 255 Then MaxB = 255



            MaxR = (MaxR + MaxG + MaxB) / 3
            MaxG = MaxR
            MaxB = MaxR

            ResSingle(2, X, Y) = Not (CByte(MaxR))
            ResSingle(1, X, Y) = Not (CByte(MaxG))
            ResSingle(0, X, Y) = Not (CByte(MaxB))


        Next Y
        '        frmMAIN.Caption = pW - X
        '        DoEvents
        If X > ProgX Then
            RaiseEvent PercDONE(X / pW, 0)
            ProgX = ProgX + ProgXStep
            '                Stop

        End If
    Next X
    RaiseEvent PercDONE(1, 1)

End Sub

