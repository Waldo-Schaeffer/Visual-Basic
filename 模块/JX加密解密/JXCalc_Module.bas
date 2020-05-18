Attribute VB_Name = "JXCalc_Module"

Dim JXByteList1() As Byte
Dim JXByteList2() As Byte

'--------字节解密全局函数--------
Public Function JXByteDecrypt(ByteNum As Byte) As Byte

On Error Resume Next

If LBound(JXByteList1) <> 0 Or UBound(JXByteList1) <> 255 Then Call JXInitialize
If LBound(JXByteList2) <> 0 Or UBound(JXByteList2) <> 255 Then Call JXInitialize

For i = 0 To 255
  If ByteNum = JXByteList2(i) Then
    JXByteDecrypt = JXByteList1(i)
    Exit For
  End If
Next i

End Function

'--------字节加密全局函数--------
Public Function JXByteEncrypt(ByteNum As Byte) As Byte

On Error Resume Next

If LBound(JXByteList1) <> 0 Or UBound(JXByteList1) <> 255 Then Call JXInitialize
If LBound(JXByteList2) <> 0 Or UBound(JXByteList2) <> 255 Then Call JXInitialize

For i = 0 To 255
  If ByteNum = JXByteList1(i) Then
    JXByteEncrypt = JXByteList2(i)
    Exit For
  End If
Next i

End Function

'--------数组解密全局函数--------
Public Function JXDecrypt(GetByteArray() As Byte, ByRef PutByteArray() As Byte) As Boolean

On Error GoTo Err

JXDecrypt = False

ReDim PutByteArray(LBound(GetByteArray) To UBound(GetByteArray)) As Byte

For i = LBound(PutByteArray) To UBound(PutByteArray)
  PutByteArray(i) = JXByteDecrypt(GetByteArray(i))
Next i

JXDecrypt = True

Exit Function

Err:
JXDecrypt = False

End Function

'--------数组加密全局函数--------
Public Function JXEncrypt(GetByteArray() As Byte, ByRef PutByteArray() As Byte) As Boolean

On Error GoTo Err

JXEncrypt = False

ReDim PutByteArray(LBound(GetByteArray) To UBound(GetByteArray)) As Byte

For i = LBound(PutByteArray) To UBound(PutByteArray)
  PutByteArray(i) = JXByteEncrypt(GetByteArray(i))
Next i

JXEncrypt = True

Exit Function

Err:
JXEncrypt = False

End Function

'--------变量初始全局过程--------
Public Sub JXInitialize()

On Error Resume Next

ReDim JXByteList1(0 To 255) As Byte
ReDim JXByteList2(0 To 255) As Byte

For i = 0 To 255
  JXByteList1(i) = i
Next i

JXByteList2(0) = 241
JXByteList2(1) = 249
JXByteList2(2) = 200
JXByteList2(3) = 21
JXByteList2(4) = 179
JXByteList2(5) = 60
JXByteList2(6) = 169
JXByteList2(7) = 183
JXByteList2(8) = 57
JXByteList2(9) = 162
JXByteList2(10) = 206
JXByteList2(11) = 138
JXByteList2(12) = 67
JXByteList2(13) = 63
JXByteList2(14) = 8
JXByteList2(15) = 228
JXByteList2(16) = 92
JXByteList2(17) = 190
JXByteList2(18) = 107
JXByteList2(19) = 91
JXByteList2(20) = 84
JXByteList2(21) = 0
JXByteList2(22) = 96
JXByteList2(23) = 135
JXByteList2(24) = 7
JXByteList2(25) = 246
JXByteList2(26) = 226
JXByteList2(27) = 19
JXByteList2(28) = 121
JXByteList2(29) = 203
JXByteList2(30) = 244
JXByteList2(31) = 10
JXByteList2(32) = 70
JXByteList2(33) = 136
JXByteList2(34) = 29
JXByteList2(35) = 221
JXByteList2(36) = 146
JXByteList2(37) = 69
JXByteList2(38) = 56
JXByteList2(39) = 201
JXByteList2(40) = 239
JXByteList2(41) = 39
JXByteList2(42) = 152
JXByteList2(43) = 45
JXByteList2(44) = 36
JXByteList2(45) = 204
JXByteList2(46) = 28
JXByteList2(47) = 134
JXByteList2(48) = 171
JXByteList2(49) = 193
JXByteList2(50) = 225
JXByteList2(51) = 137
JXByteList2(52) = 237
JXByteList2(53) = 88
JXByteList2(54) = 66
JXByteList2(55) = 52
JXByteList2(56) = 161
JXByteList2(57) = 53
JXByteList2(58) = 234
JXByteList2(59) = 220
JXByteList2(60) = 197
JXByteList2(61) = 71
JXByteList2(62) = 30
JXByteList2(63) = 213
JXByteList2(64) = 170
JXByteList2(65) = 47
JXByteList2(66) = 79
JXByteList2(67) = 155
JXByteList2(68) = 180
JXByteList2(69) = 182
JXByteList2(70) = 33
JXByteList2(71) = 181
JXByteList2(72) = 108
JXByteList2(73) = 35
JXByteList2(74) = 27
JXByteList2(75) = 205
JXByteList2(76) = 112
JXByteList2(77) = 94
JXByteList2(78) = 194
JXByteList2(79) = 62
JXByteList2(80) = 150
JXByteList2(81) = 122
JXByteList2(82) = 110
JXByteList2(83) = 68
JXByteList2(84) = 40
JXByteList2(85) = 9
JXByteList2(86) = 3
JXByteList2(87) = 117
JXByteList2(88) = 113
JXByteList2(89) = 59
JXByteList2(90) = 143
JXByteList2(91) = 141
JXByteList2(92) = 250
JXByteList2(93) = 130
JXByteList2(94) = 164
JXByteList2(95) = 15
JXByteList2(96) = 77
JXByteList2(97) = 178
JXByteList2(98) = 17
JXByteList2(99) = 252
JXByteList2(100) = 98
JXByteList2(101) = 148
JXByteList2(102) = 158
JXByteList2(103) = 133
JXByteList2(104) = 176
JXByteList2(105) = 253
JXByteList2(106) = 90
JXByteList2(107) = 104
JXByteList2(108) = 118
JXByteList2(109) = 31
JXByteList2(110) = 247
JXByteList2(111) = 85
JXByteList2(112) = 61
JXByteList2(113) = 128
JXByteList2(114) = 168
JXByteList2(115) = 93
JXByteList2(116) = 111
JXByteList2(117) = 192
JXByteList2(118) = 163
JXByteList2(119) = 22
JXByteList2(120) = 172
JXByteList2(121) = 223
JXByteList2(122) = 46
JXByteList2(123) = 142
JXByteList2(124) = 217
JXByteList2(125) = 26
JXByteList2(126) = 202
JXByteList2(127) = 48
JXByteList2(128) = 11
JXByteList2(129) = 4
JXByteList2(130) = 105
JXByteList2(131) = 123
JXByteList2(132) = 124
JXByteList2(133) = 101
JXByteList2(134) = 12
JXByteList2(135) = 131
JXByteList2(136) = 188
JXByteList2(137) = 251
JXByteList2(138) = 73
JXByteList2(139) = 184
JXByteList2(140) = 78
JXByteList2(141) = 54
JXByteList2(142) = 224
JXByteList2(143) = 186
JXByteList2(144) = 248
JXByteList2(145) = 115
JXByteList2(146) = 51
JXByteList2(147) = 109
JXByteList2(148) = 173
JXByteList2(149) = 154
JXByteList2(150) = 214
JXByteList2(151) = 100
JXByteList2(152) = 140
JXByteList2(153) = 210
JXByteList2(154) = 255
JXByteList2(155) = 103
JXByteList2(156) = 151
JXByteList2(157) = 199
JXByteList2(158) = 153
JXByteList2(159) = 58
JXByteList2(160) = 191
JXByteList2(161) = 230
JXByteList2(162) = 13
JXByteList2(163) = 20
JXByteList2(164) = 189
JXByteList2(165) = 72
JXByteList2(166) = 227
JXByteList2(167) = 32
JXByteList2(168) = 160
JXByteList2(169) = 212
JXByteList2(170) = 76
JXByteList2(171) = 144
JXByteList2(172) = 185
JXByteList2(173) = 211
JXByteList2(174) = 167
JXByteList2(175) = 87
JXByteList2(176) = 233
JXByteList2(177) = 16
JXByteList2(178) = 177
JXByteList2(179) = 65
JXByteList2(180) = 34
JXByteList2(181) = 116
JXByteList2(182) = 222
JXByteList2(183) = 196
JXByteList2(184) = 145
JXByteList2(185) = 95
JXByteList2(186) = 119
JXByteList2(187) = 215
JXByteList2(188) = 97
JXByteList2(189) = 99
JXByteList2(190) = 49
JXByteList2(191) = 236
JXByteList2(192) = 2
JXByteList2(193) = 126
JXByteList2(194) = 139
JXByteList2(195) = 132
JXByteList2(196) = 243
JXByteList2(197) = 25
JXByteList2(198) = 216
JXByteList2(199) = 6
JXByteList2(200) = 14
JXByteList2(201) = 174
JXByteList2(202) = 187
JXByteList2(203) = 156
JXByteList2(204) = 37
JXByteList2(205) = 50
JXByteList2(206) = 114
JXByteList2(207) = 89
JXByteList2(208) = 23
JXByteList2(209) = 232
JXByteList2(210) = 18
JXByteList2(211) = 129
JXByteList2(212) = 55
JXByteList2(213) = 254
JXByteList2(214) = 44
JXByteList2(215) = 209
JXByteList2(216) = 81
JXByteList2(217) = 120
JXByteList2(218) = 245
JXByteList2(219) = 165
JXByteList2(220) = 231
JXByteList2(221) = 24
JXByteList2(222) = 75
JXByteList2(223) = 125
JXByteList2(224) = 43
JXByteList2(225) = 82
JXByteList2(226) = 106
JXByteList2(227) = 195
JXByteList2(228) = 235
JXByteList2(229) = 242
JXByteList2(230) = 38
JXByteList2(231) = 229
JXByteList2(232) = 83
JXByteList2(233) = 238
JXByteList2(234) = 208
JXByteList2(235) = 166
JXByteList2(236) = 42
JXByteList2(237) = 157
JXByteList2(238) = 198
JXByteList2(239) = 86
JXByteList2(240) = 64
JXByteList2(241) = 159
JXByteList2(242) = 74
JXByteList2(243) = 219
JXByteList2(244) = 207
JXByteList2(245) = 41
JXByteList2(246) = 147
JXByteList2(247) = 218
JXByteList2(248) = 80
JXByteList2(249) = 5
JXByteList2(250) = 1
JXByteList2(251) = 149
JXByteList2(252) = 240
JXByteList2(253) = 102
JXByteList2(254) = 127
JXByteList2(255) = 175

End Sub
