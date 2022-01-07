Attribute VB_Name = "Module1"
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hrgnParentRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Dim r0 As Long, lebarRgnParent As Long
Public Sub p(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    p1 = CreateRoundRectRgn(180, 120, 210, 130, 5, 5)
    p2 = CreateRoundRectRgn(180, 170, 200, 160, 5, 5)
    p3 = CreateRoundRectRgn(185, 120, 195, 170, 10, 15)
    
    p4 = CreateRoundRectRgn(204, 122, 212, 135, 10, 5)
    p5 = CreateRoundRectRgn(205, 125, 213, 137, 10, 5)
    p6 = CreateRoundRectRgn(205, 128, 214, 141, 10, 5)
    p7 = CreateRoundRectRgn(205, 131, 214, 141, 10, 5)
    p7a = CreateRoundRectRgn(205, 134, 214, 143, 10, 5)
    p7b = CreateRoundRectRgn(204, 137, 213, 145, 10, 5)
    p8 = CreateRoundRectRgn(203, 140, 212, 147, 10, 5)
    p9 = CreateRoundRectRgn(202, 143, 211, 149, 10, 5)
    p10 = CreateRoundRectRgn(201, 146, 210, 151, 10, 5)
    p11 = CreateRoundRectRgn(185, 143, 209, 151, 5, 5)
    
    CombineRgn bg, bg, p1, 2
    CombineRgn bg, bg, p2, 2
    CombineRgn bg, bg, p3, 2
    CombineRgn bg, bg, p4, 2
    CombineRgn bg, bg, p5, 2
    CombineRgn bg, bg, p6, 2
    CombineRgn bg, bg, p7, 2
    CombineRgn bg, bg, p7a, 2
    CombineRgn bg, bg, p7b, 2
    CombineRgn bg, bg, p8, 2
    CombineRgn bg, bg, p9, 2
    CombineRgn bg, bg, p10, 2
    CombineRgn bg, bg, p11, 2
    
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 180, y - 20 / 2 - 120
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub

Public Sub u(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    u1 = CreateRoundRectRgn(562, 120, 585, 130, 5, 5)
    u2 = CreateRoundRectRgn(567, 120, 580, 163, 10, 5)
    u1a = CreateRoundRectRgn(589, 120, 612, 130, 5, 5)
    u2a = CreateRoundRectRgn(594, 120, 607, 163, 10, 5)
    u3 = CreateRoundRectRgn(567, 155, 581, 163, 10, 5)
    u4 = CreateRoundRectRgn(568, 156, 582, 165, 10, 5)
    u5 = CreateRoundRectRgn(569, 157, 583, 166, 10, 5)
    u6 = CreateRoundRectRgn(570, 158, 584, 168, 10, 5)
    u7 = CreateRoundRectRgn(571, 159, 587, 170, 10, 5)
    u8 = CreateRoundRectRgn(592, 155, 607, 163, 10, 5)
    u9 = CreateRoundRectRgn(591, 156, 606, 165, 10, 5)
    u10 = CreateRoundRectRgn(590, 157, 605, 166, 10, 5)
    u11 = CreateRoundRectRgn(589, 158, 604, 168, 10, 5)
    u12 = CreateRoundRectRgn(578, 159, 603, 170, 10, 5)
    
    CombineRgn bg, bg, u1, 2
    CombineRgn bg, bg, u2, 2
    CombineRgn bg, bg, u1a, 2
    CombineRgn bg, bg, u2a, 2
    CombineRgn bg, bg, u3, 2
    CombineRgn bg, bg, u4, 2
    CombineRgn bg, bg, u5, 2
    CombineRgn bg, bg, u6, 2
    CombineRgn bg, bg, u7, 2
    CombineRgn bg, bg, u8, 2
    CombineRgn bg, bg, u9, 2
    CombineRgn bg, bg, u10, 2
    CombineRgn bg, bg, u11, 2
    CombineRgn bg, bg, u12, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 562, y - 20 / 2 - 120
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub

Public Sub t(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    t1 = CreateRoundRectRgn(485, 120, 525, 130, 5, 5)
    t2 = CreateRoundRectRgn(500, 120, 510, 170, 5, 5)
    t3 = CreateRoundRectRgn(495, 162, 515, 170, 5, 5)
    
    t4 = CreateRoundRectRgn(485, 126, 495, 132, 25, 1)
    t5 = CreateRoundRectRgn(485, 127, 494, 133, 25, 1)
    t6 = CreateRoundRectRgn(485, 128, 493, 134, 25, 1)
    
    t7 = CreateRoundRectRgn(515, 126, 525, 132, 25, 1)
    t8 = CreateRoundRectRgn(516, 127, 525, 133, 25, 1)
    t9 = CreateRoundRectRgn(517, 128, 525, 134, 25, 1)
    
    CombineRgn bg, bg, t1, 2
    CombineRgn bg, bg, t2, 2
    CombineRgn bg, bg, t3, 2
    CombineRgn bg, bg, t4, 2
    CombineRgn bg, bg, t5, 2
    CombineRgn bg, bg, t6, 2
    CombineRgn bg, bg, t7, 2
    CombineRgn bg, bg, t8, 2
    CombineRgn bg, bg, t9, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 485, y - 20 / 2 - 120
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub n(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    n1 = CreateRoundRectRgn(19, 120, 36, 130, 5, 5)
    n2 = CreateRoundRectRgn(24, 120, 34, 170, 5, 5)
    n3 = CreateRoundRectRgn(19, 170, 39, 160, 5, 5)
    
    n4 = CreateRoundRectRgn(42, 120, 62, 130, 5, 5)
    n5 = CreateRoundRectRgn(47, 120, 57, 170, 5, 5)
    n6 = CreateRoundRectRgn(47, 170, 56, 160, 5, 5)
    
    n7 = CreateRoundRectRgn(24, 122, 39, 136, 5, 5)
    n8 = CreateRoundRectRgn(25, 124, 40, 138, 5, 5)
    n9 = CreateRoundRectRgn(26, 126, 41, 140, 5, 5)
    n10 = CreateRoundRectRgn(27, 128, 42, 142, 5, 5)
    n11 = CreateRoundRectRgn(28, 130, 43, 144, 5, 5)
    n12 = CreateRoundRectRgn(29, 132, 44, 146, 5, 5)
    n13 = CreateRoundRectRgn(30, 134, 45, 148, 5, 5)
    n14 = CreateRoundRectRgn(33, 136, 46, 150, 5, 5)
    n15 = CreateRoundRectRgn(34, 138, 47, 152, 5, 5)
    n16 = CreateRoundRectRgn(35, 140, 48, 154, 5, 5)
    n17 = CreateRoundRectRgn(36, 142, 49, 156, 5, 5)
    n18 = CreateRoundRectRgn(37, 144, 50, 158, 5, 5)
    n19 = CreateRoundRectRgn(38, 146, 51, 160, 5, 5)
    n20 = CreateRoundRectRgn(39, 148, 52, 162, 5, 5)
    n21 = CreateRoundRectRgn(40, 150, 53, 164, 5, 5)
    n22 = CreateRoundRectRgn(41, 152, 54, 166, 5, 5)
    n23 = CreateRoundRectRgn(42, 154, 55, 168, 5, 5)
    n24 = CreateRoundRectRgn(43, 156, 56, 170, 5, 5)
    
    CombineRgn bg, bg, n1, 2
    CombineRgn bg, bg, n2, 2
    CombineRgn bg, bg, n3, 2
    CombineRgn bg, bg, n4, 2
    CombineRgn bg, bg, n5, 2
    CombineRgn bg, bg, n6, 2
    CombineRgn bg, bg, n7, 2
    CombineRgn bg, bg, n8, 2
    CombineRgn bg, bg, n9, 2
    CombineRgn bg, bg, n10, 2
    CombineRgn bg, bg, n11, 2
    CombineRgn bg, bg, n12, 2
    CombineRgn bg, bg, n13, 2
    CombineRgn bg, bg, n14, 2
    CombineRgn bg, bg, n15, 2
    CombineRgn bg, bg, n16, 2
    CombineRgn bg, bg, n17, 2
    CombineRgn bg, bg, n18, 2
    CombineRgn bg, bg, n19, 2
    CombineRgn bg, bg, n20, 2
    CombineRgn bg, bg, n21, 2
    CombineRgn bg, bg, n22, 2
    CombineRgn bg, bg, n23, 2
    CombineRgn bg, bg, n24, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 19, y - 20 / 2 - 120
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub a(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    a1 = CreateRoundRectRgn(30, 30, 50, 40, 10, 5)
    a2 = CreateRoundRectRgn(29, 35, 51, 45, 10, 5)
    a3 = CreateRoundRectRgn(28, 40, 39, 50, 10, 5)
    a4 = CreateRoundRectRgn(27, 45, 38, 55, 10, 5)
    a5 = CreateRoundRectRgn(26, 50, 37, 60, 10, 5)
    a6 = CreateRoundRectRgn(25, 55, 36, 65, 10, 5)
    a7 = CreateRoundRectRgn(24, 60, 56, 70, 10, 5)
    a8 = CreateRoundRectRgn(42, 40, 52, 50, 10, 5)
    a9 = CreateRoundRectRgn(43, 45, 53, 55, 10, 5)
    a10 = CreateRoundRectRgn(44, 50, 54, 60, 10, 5)
    a11 = CreateRoundRectRgn(45, 55, 55, 75, 10, 5)
    a12 = CreateRoundRectRgn(23, 65, 34, 75, 10, 5)
    a13 = CreateRoundRectRgn(22, 70, 33, 80, 10, 5)
    a14 = CreateRoundRectRgn(19, 75, 37, 82, 5, 5)
    a15 = CreateRoundRectRgn(47, 65, 57, 75, 10, 5)
    a16 = CreateRoundRectRgn(46, 70, 58, 80, 10, 5)
    a17 = CreateRoundRectRgn(43, 75, 61, 82, 5, 5)
    
    CombineRgn bg, bg, a1, 2
    CombineRgn bg, bg, a2, 2
    CombineRgn bg, bg, a3, 2
    CombineRgn bg, bg, a4, 2
    CombineRgn bg, bg, a5, 2
    CombineRgn bg, bg, a6, 2
    CombineRgn bg, bg, a7, 2
    CombineRgn bg, bg, a8, 2
    CombineRgn bg, bg, a9, 2
    CombineRgn bg, bg, a10, 2
    CombineRgn bg, bg, a11, 2
    CombineRgn bg, bg, a12, 2
    CombineRgn bg, bg, a13, 2
    CombineRgn bg, bg, a14, 2
    CombineRgn bg, bg, a15, 2
    CombineRgn bg, bg, a16, 2
    CombineRgn bg, bg, a17, 2
    
        
    lebarRgn = 40
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 19, y - 20 / 2 - 30
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub d(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    d1 = CreateRoundRectRgn(255, 30, 288, 40, 5, 5)
    d2 = CreateRoundRectRgn(257, 35, 273, 80, 10, 5)
    d3 = CreateRoundRectRgn(255, 80, 288, 70, 5, 5)
    
    d4 = CreateRoundRectRgn(270, 32, 282, 42, 10, 5)
    d5 = CreateRoundRectRgn(270, 76, 289, 66, 10, 5)
    
    d6 = CreateRoundRectRgn(275, 34, 290, 44, 10, 5)
    d7 = CreateRoundRectRgn(276, 36, 291, 46, 10, 5)
    d8 = CreateRoundRectRgn(277, 38, 292, 48, 10, 5)
    d9 = CreateRoundRectRgn(278, 40, 293, 50, 10, 5)
    d10 = CreateRoundRectRgn(279, 42, 294, 52, 10, 5)
    d11 = CreateRoundRectRgn(280, 44, 295, 54, 10, 5)
    d12 = CreateRoundRectRgn(281, 46, 296, 56, 10, 5)
    d13 = CreateRoundRectRgn(282, 48, 296, 58, 10, 5)
    d14 = CreateRoundRectRgn(282, 50, 296, 60, 10, 5)
    d15 = CreateRoundRectRgn(281, 52, 296, 62, 10, 5)
    d16 = CreateRoundRectRgn(280, 54, 296, 64, 10, 5)
    d17 = CreateRoundRectRgn(279, 56, 295, 66, 10, 5)
    d18 = CreateRoundRectRgn(278, 58, 294, 68, 10, 5)
    d19 = CreateRoundRectRgn(277, 60, 293, 70, 10, 5)
    d20 = CreateRoundRectRgn(276, 62, 292, 72, 10, 5)
    d21 = CreateRoundRectRgn(275, 64, 291, 74, 10, 5)
    d22 = CreateRoundRectRgn(274, 66, 290, 76, 10, 5)
    
    CombineRgn bg, bg, d1, 2
    CombineRgn bg, bg, d2, 2
    CombineRgn bg, bg, d3, 2
    CombineRgn bg, bg, d4, 2
    CombineRgn bg, bg, d5, 2
    CombineRgn bg, bg, d6, 2
    CombineRgn bg, bg, d7, 2
    CombineRgn bg, bg, d8, 2
    CombineRgn bg, bg, d9, 2
    CombineRgn bg, bg, d11, 2
    CombineRgn bg, bg, d12, 2
    CombineRgn bg, bg, d13, 2
    CombineRgn bg, bg, d14, 2
    CombineRgn bg, bg, d15, 2
    CombineRgn bg, bg, d16, 2
    CombineRgn bg, bg, d17, 2
    CombineRgn bg, bg, d18, 2
    CombineRgn bg, bg, d19, 2
    CombineRgn bg, bg, d20, 2
    CombineRgn bg, bg, d21, 2
    CombineRgn bg, bg, d22, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 255, y - 20 / 2 - 30
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub h(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    h1 = CreateRoundRectRgn(562, 30, 585, 40, 5, 5)
    h2 = CreateRoundRectRgn(566, 30, 581, 80, 10, 5)
    h3 = CreateRoundRectRgn(562, 80, 585, 70, 5, 5)
    
    h4 = CreateRoundRectRgn(568, 50, 607, 60, 10, 5)
    
    h5 = CreateRoundRectRgn(590, 30, 613, 40, 5, 5)
    h6 = CreateRoundRectRgn(594, 30, 609, 80, 10, 5)
    h7 = CreateRoundRectRgn(590, 80, 613, 70, 5, 5)
    
    CombineRgn bg, bg, h1, 2
    CombineRgn bg, bg, h2, 2
    CombineRgn bg, bg, h3, 2
    CombineRgn bg, bg, h4, 2
    CombineRgn bg, bg, h5, 2
    CombineRgn bg, bg, h6, 2
    CombineRgn bg, bg, h7, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 562, y - 20 / 2 - 30
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub i(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    i1 = CreateRoundRectRgn(643, 30, 666, 40, 5, 5)
    i2 = CreateRoundRectRgn(649, 30, 660, 80, 10, 5)
    i3 = CreateRoundRectRgn(643, 80, 666, 70, 5, 5)
    
    CombineRgn bg, bg, i1, 2
    CombineRgn bg, bg, i2, 2
    CombineRgn bg, bg, i3, 2
    
    lebarRgn = 20
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 643, y - 20 / 2 - 30

    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub k(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    k1 = CreateRoundRectRgn(765, 30, 786, 40, 5, 5)
    k2 = CreateRoundRectRgn(771, 30, 782, 80, 5, 5)
    k3 = CreateRoundRectRgn(765, 80, 786, 70, 5, 5)
    
    k4 = CreateRoundRectRgn(787, 30, 809, 40, 5, 5)
    k5 = CreateRoundRectRgn(790, 69, 809, 80, 5, 5)
    
    k6 = CreateRoundRectRgn(771, 50, 796, 60, 5, 5)
    
    k7 = CreateRoundRectRgn(783, 48, 796, 58, 10, 5)
    k8 = CreateRoundRectRgn(784, 46, 797, 56, 10, 5)
    k9 = CreateRoundRectRgn(785, 44, 798, 54, 10, 5)
    k10 = CreateRoundRectRgn(786, 42, 799, 52, 10, 5)
    k11 = CreateRoundRectRgn(787, 40, 800, 50, 10, 5)
    k12 = CreateRoundRectRgn(788, 38, 801, 48, 10, 5)
    k13 = CreateRoundRectRgn(789, 36, 802, 46, 10, 5)
    k14 = CreateRoundRectRgn(789, 34, 803, 44, 10, 5)
    
    k15 = CreateRoundRectRgn(783, 52, 797, 62, 10, 5)
    k16 = CreateRoundRectRgn(784, 54, 798, 64, 10, 5)
    k17 = CreateRoundRectRgn(785, 56, 799, 66, 10, 5)
    k18 = CreateRoundRectRgn(786, 58, 800, 68, 10, 5)
    k19 = CreateRoundRectRgn(787, 60, 801, 70, 10, 5)
    k20 = CreateRoundRectRgn(788, 62, 802, 72, 10, 5)
    k21 = CreateRoundRectRgn(789, 64, 803, 74, 10, 5)
    k22 = CreateRoundRectRgn(790, 66, 804, 76, 10, 5)
    
    CombineRgn bg, bg, k1, 2
    CombineRgn bg, bg, k2, 2
    CombineRgn bg, bg, k3, 2
    CombineRgn bg, bg, k4, 2
    CombineRgn bg, bg, k5, 2
    CombineRgn bg, bg, k6, 2
    CombineRgn bg, bg, k7, 2
    CombineRgn bg, bg, k8, 2
    CombineRgn bg, bg, k9, 2
    CombineRgn bg, bg, k10, 2
    CombineRgn bg, bg, k11, 2
    CombineRgn bg, bg, k12, 2
    CombineRgn bg, bg, k13, 2
    CombineRgn bg, bg, k14, 2
    CombineRgn bg, bg, k15, 2
    CombineRgn bg, bg, k16, 2
    CombineRgn bg, bg, k17, 2
    CombineRgn bg, bg, k18, 2
    CombineRgn bg, bg, k19, 2
    CombineRgn bg, bg, k20, 2
    CombineRgn bg, bg, k21, 2
    CombineRgn bg, bg, k22, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 765, y - 20 / 2 - 30
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub r(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    r1 = CreateRoundRectRgn(335, 120, 368, 130, 5, 5)
    r2 = CreateRoundRectRgn(335, 170, 355, 160, 5, 5)
    r3 = CreateRoundRectRgn(340, 120, 350, 170, 10, 15)
    
    r4 = CreateRoundRectRgn(360, 122, 370, 135, 10, 5)
    r5 = CreateRoundRectRgn(361, 125, 371, 137, 10, 5)
    r6 = CreateRoundRectRgn(361, 128, 372, 141, 10, 5)
    r7 = CreateRoundRectRgn(361, 131, 372, 141, 10, 5)
    r7a = CreateRoundRectRgn(361, 134, 372, 143, 10, 5)
    r7b = CreateRoundRectRgn(360, 137, 371, 145, 10, 5)
    r8 = CreateRoundRectRgn(359, 140, 370, 147, 10, 5)
    r9 = CreateRoundRectRgn(358, 143, 369, 149, 10, 5)
    r10 = CreateRoundRectRgn(357, 146, 368, 151, 10, 5)
    r11 = CreateRoundRectRgn(345, 143, 367, 151, 5, 5)
    
    r12 = CreateRoundRectRgn(358, 145, 368, 155, 5, 5)
    r13 = CreateRoundRectRgn(359, 150, 369, 160, 5, 5)
    r14 = CreateRoundRectRgn(360, 155, 370, 165, 5, 5)
    r15 = CreateRoundRectRgn(361, 160, 373, 170, 5, 5)
    
    CombineRgn bg, bg, r1, 2
    CombineRgn bg, bg, r2, 2
    CombineRgn bg, bg, r3, 2
    CombineRgn bg, bg, r4, 2
    CombineRgn bg, bg, r5, 2
    CombineRgn bg, bg, r6, 2
    CombineRgn bg, bg, r7, 2
    CombineRgn bg, bg, r7a, 2
    CombineRgn bg, bg, r7b, 2
    CombineRgn bg, bg, r8, 2
    CombineRgn bg, bg, r9, 2
    CombineRgn bg, bg, r10, 2
    CombineRgn bg, bg, r11, 2
    CombineRgn bg, bg, r12, 2
    CombineRgn bg, bg, r13, 2
    CombineRgn bg, bg, r14, 2
    CombineRgn bg, bg, r15, 2
    CombineRgn bg, bg, r16, 2
    CombineRgn bg, bg, r17, 2
    CombineRgn bg, bg, b13, 2
    
    lebarRgn = 20
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 335, y - 20 / 2 - 120

    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub m(ByVal x As Long, ByVal y As Long)
    r0 = getWinRgn
    bg = CreateRectRgn(0, 0, 0, 0)
    
    m1 = CreateRoundRectRgn(925, 30, 948, 40, 5, 5)
    m2 = CreateRoundRectRgn(931, 30, 942, 80, 10, 5)
    m3 = CreateRoundRectRgn(925, 80, 948, 70, 5, 5)
    
    m1a = CreateRoundRectRgn(956, 30, 980, 40, 5, 5)
    m2a = CreateRoundRectRgn(963, 30, 974, 80, 10, 5)
    m3a = CreateRoundRectRgn(956, 80, 980, 70, 5, 5)
    
    m4 = CreateRoundRectRgn(935, 35, 949, 45, 5, 5)
    m5 = CreateRoundRectRgn(943, 40, 950, 50, 5, 5)
    m6 = CreateRoundRectRgn(944, 45, 951, 55, 5, 5)
    m7 = CreateRoundRectRgn(945, 50, 952, 60, 5, 5)
    m8 = CreateRoundRectRgn(946, 55, 953, 65, 5, 5)
    m9 = CreateRoundRectRgn(947, 60, 954, 70, 5, 5)
    
    m10 = CreateRoundRectRgn(954, 35, 970, 45, 5, 5)
    m11 = CreateRoundRectRgn(953, 40, 962, 50, 5, 5)
    m12 = CreateRoundRectRgn(952, 45, 961, 55, 5, 5)
    m13 = CreateRoundRectRgn(951, 50, 960, 60, 5, 5)
    m14 = CreateRoundRectRgn(950, 55, 959, 65, 5, 5)
    m15 = CreateRoundRectRgn(949, 60, 958, 70, 5, 5)
    
    CombineRgn bg, bg, m1, 2
    CombineRgn bg, bg, m2, 2
    CombineRgn bg, bg, m3, 2
    CombineRgn bg, bg, m1a, 2
    CombineRgn bg, bg, m2a, 2
    CombineRgn bg, bg, m3a, 2
    CombineRgn bg, bg, m4, 2
    CombineRgn bg, bg, m5, 2
    CombineRgn bg, bg, m6, 2
    CombineRgn bg, bg, m7, 2
    CombineRgn bg, bg, m8, 2
    CombineRgn bg, bg, m9, 2
    CombineRgn bg, bg, m10, 2
    CombineRgn bg, bg, m11, 2
    CombineRgn bg, bg, m12, 2
    CombineRgn bg, bg, m13, 2
    CombineRgn bg, bg, m14, 2
    CombineRgn bg, bg, m15, 2
    CombineRgn bg, bg, m16, 2
        
    lebarRgn = 30
    
    lebarRgnParent = lebarRgnParent + 50
    
    OffsetRgn bg, x - (lebarRgn / 2) - 925, y - 20 / 2 - 30
    
    'Combine ke parent region
    CombineRgn r0, r0, bg, 2
    
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub init()
    r0 = CreateRectRgn(0, 0, 0, 0)
    SetWindowRgn Form1.hWnd, r0, True
End Sub
Public Sub delete()
    DeleteObject r0
    DeleteObject combined
End Sub
Public Sub geserKiri()
    If (lebarRgnParent <> 0) Then
        combined = getWinRgn
        OffsetRgn combined, -50, 0
        SetWindowRgn Form1.hWnd, combined, True
    End If
End Sub
Public Function getWinRgn() As Long
    combined = CreateRectRgn(0, 0, 0, 0)
    GetWindowRgn Form1.hWnd, combined
    getWinRgn = combined
End Function
Public Sub geserKananSet()
    combined = getWinRgn
    OffsetRgn combined, lebarRgnParent / 2, 0
    SetWindowRgn Form1.hWnd, combined, True
End Sub
Public Sub geserKiriSet()
    combined = getWinRgn
    OffsetRgn combined, -lebarRgnParent / 2, 0
    SetWindowRgn Form1.hWnd, combined, True
End Sub
