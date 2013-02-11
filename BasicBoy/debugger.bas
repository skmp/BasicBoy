Attribute VB_Name = "debugger"
Option Explicit
Dim lasdr As Long
Public Function ToAsm1(ByVal val As Long, ByRef adr As Long) As String
Dim tmp As Long
Dim pre As String
pre = String(4 - Len(Hex(adr)), "0") & Hex(adr) & ": "
adr = adr + 1
 Select Case val
 Case &H0: ToAsm1 = "NOP"
 Case &H1: ToAsm1 = "LD BC," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &H2: ToAsm1 = "LD (BC),A"
 Case &H3: ToAsm1 = "INC BC"
 Case &H4: ToAsm1 = "INC B"
 Case &H5: ToAsm1 = "DEC B"
 Case &H6: ToAsm1 = "LD B," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &H7: ToAsm1 = "RLC A"
 Case &H8: ToAsm1 = "LD (" & Hex(readM(adr) + readM(adr + 1) * 256) & "H), SP": adr = adr + 1: adr = adr + 1
 Case &H9: ToAsm1 = "ADD HL,BC"
 Case &HA: ToAsm1 = "LD A,(BC)"
 Case &HB: ToAsm1 = "DEC BC"
 Case &HC: ToAsm1 = "INC C"
 Case &HD: ToAsm1 = "DEC C"
 Case &HE: ToAsm1 = "LD C," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HF: ToAsm1 = "RRC A"
 Case &H10: ToAsm1 = "STOP": adr = adr + 1
 Case &H11: ToAsm1 = "LD DE," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &H12: ToAsm1 = "LD (DE),A"
 Case &H13: ToAsm1 = "INC DE"
 Case &H14: ToAsm1 = "INC D"
 Case &H15: ToAsm1 = "DEC D"
 Case &H16: ToAsm1 = "LD D," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &H17: ToAsm1 = "RLA"
 Case &H18:
  tmp = readM(adr)
  If tmp And 128 Then
   tmp = 256 - tmp
   ToAsm1 = "JR -" & Hex(tmp) & "H"
  Else
   ToAsm1 = "JR " & Hex(tmp) & "H"
  End If
 Case &H19: ToAsm1 = "ADD HL,DE"
 Case &H1A: ToAsm1 = "LD A,(DE)"
 Case &H1B: ToAsm1 = "DEC DE"
 Case &H1C: ToAsm1 = "INC E"
 Case &H1D: ToAsm1 = "DEC E"
 Case &H1E: ToAsm1 = "LD E," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &H1F: ToAsm1 = "RRA"
 Case &H20:
  tmp = readM(adr)
  If tmp And 128 Then
   tmp = 256 - tmp
   ToAsm1 = "JR NZ,-" & Hex(tmp) & "H"
  Else
   ToAsm1 = "JR NZ," & Hex(tmp) & "H"
  End If
 Case &H21: ToAsm1 = "LD HL," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &H22: ToAsm1 = "LDI (HL), A"
 Case &H23: ToAsm1 = "INC HL"
 Case &H24: ToAsm1 = "INC H"
 Case &H25: ToAsm1 = "DEC H"
 Case &H26: ToAsm1 = "LD H," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &H27: ToAsm1 = "DAA"
 Case &H28:
  tmp = readM(adr)
  If tmp And 128 Then
   tmp = 256 - tmp
   ToAsm1 = "JR Z,-" & Hex(tmp) & "H"
  Else
   ToAsm1 = "JR Z," & Hex(tmp) & "H"
  End If
 Case &H29: ToAsm1 = "ADD HL,HL"
 Case &H2A: ToAsm1 = "LDI A,(HL)"
 Case &H2B: ToAsm1 = "DEC HL"
 Case &H2C: ToAsm1 = "INC L"
 Case &H2D: ToAsm1 = "DEC L"
 Case &H2E: ToAsm1 = "LD L," & Hex(readM(adr)) & "H": adr = adr + 1: adr = adr + 1
 Case &H2F: ToAsm1 = "CPL"
 Case &H30:
  tmp = readM(adr)
  If tmp And 128 Then
   tmp = 256 - tmp
   ToAsm1 = "JR NC,-" & Hex(tmp) & "H"
  Else
   ToAsm1 = "JR NC," & Hex(tmp) & "H"
  End If
 Case &H31: ToAsm1 = "LD SP," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &H32: ToAsm1 = "LDD (HL),A"
 Case &H33: ToAsm1 = "INC SP"
 Case &H34: ToAsm1 = "INC (HL)"
 Case &H35: ToAsm1 = "DEC (HL)"
 Case &H36: ToAsm1 = "LD (HL)," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &H37: ToAsm1 = "SCF"
 Case &H38:
  tmp = readM(adr)
  If tmp And 128 Then
   tmp = 256 - tmp
   ToAsm1 = "JR C,-" & Hex(tmp) & "H"
  Else
   ToAsm1 = "JR C," & Hex(tmp) & "H"
  End If
 Case &H39: ToAsm1 = "ADD HL,SP"
 Case &H3A: ToAsm1 = "LDD A,(HL)"
 Case &H3B: ToAsm1 = "DEC SP"
 Case &H3C: ToAsm1 = "INC A"
 Case &H3D: ToAsm1 = "DEC A"
 Case &H3E: ToAsm1 = "LD A," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &H3F: ToAsm1 = "CCF"
 Case &H40: ToAsm1 = "LD B,B"
 Case &H41: ToAsm1 = "LD B,C"
 Case &H42: ToAsm1 = "LD B,D"
 Case &H43: ToAsm1 = "LD B,E"
 Case &H44: ToAsm1 = "LD B,H"
 Case &H45: ToAsm1 = "LD B,L"
 Case &H46: ToAsm1 = "LD B,(HL)"
 Case &H47: ToAsm1 = "LD B,A"
 Case &H48: ToAsm1 = "LD C,B"
 Case &H49: ToAsm1 = "LD C,C"
 Case &H4A: ToAsm1 = "LD C,D"
 Case &H4B: ToAsm1 = "LD C,E"
 Case &H4C: ToAsm1 = "LD C,H"
 Case &H4D: ToAsm1 = "LD C,L"
 Case &H4E: ToAsm1 = "LD C,(HL)"
 Case &H4F: ToAsm1 = "LD C,A"
 Case &H50: ToAsm1 = "LD D,B"
 Case &H51: ToAsm1 = "LD D,C"
 Case &H52: ToAsm1 = "LD D,D"
 Case &H53: ToAsm1 = "LD D,E"
 Case &H54: ToAsm1 = "LD D,H"
 Case &H55: ToAsm1 = "LD D,L"
 Case &H56: ToAsm1 = "LD D,(HL)"
 Case &H57: ToAsm1 = "LD D,A"
 Case &H58: ToAsm1 = "LD E,B"
 Case &H59: ToAsm1 = "LD E,C"
 Case &H5A: ToAsm1 = "LD E,D"
 Case &H5B: ToAsm1 = "LD E,E"
 Case &H5C: ToAsm1 = "LD E,H"
 Case &H5D: ToAsm1 = "LD E,L"
 Case &H5E: ToAsm1 = "LD E,(HL)"
 Case &H5F: ToAsm1 = "LD E,A"
 Case &H60: ToAsm1 = "LD H,B"
 Case &H61: ToAsm1 = "LD H,C"
 Case &H62: ToAsm1 = "LD H,D"
 Case &H63: ToAsm1 = "LD H,E"
 Case &H64: ToAsm1 = "LD H,H"
 Case &H65: ToAsm1 = "LD H,L"
 Case &H66: ToAsm1 = "LD H,(HL)"
 Case &H67: ToAsm1 = "LD H,A"
 Case &H68: ToAsm1 = "LD L,B"
 Case &H69: ToAsm1 = "LD L,C"
 Case &H6A: ToAsm1 = "LD L,D"
 Case &H6B: ToAsm1 = "LD L,E"
 Case &H6C: ToAsm1 = "LD L,H"
 Case &H6D: ToAsm1 = "LD L,L"
 Case &H6E: ToAsm1 = "LD L,(HL)"
 Case &H6F: ToAsm1 = "LD L,A"
 Case &H70: ToAsm1 = "LD (HL),B"
 Case &H71: ToAsm1 = "LD (HL),C"
 Case &H72: ToAsm1 = "LD (HL),D"
 Case &H73: ToAsm1 = "LD (HL),E"
 Case &H74: ToAsm1 = "LD (HL),H"
 Case &H75: ToAsm1 = "LD (HL),L"
 Case &H76: ToAsm1 = "HALT"
 Case &H77: ToAsm1 = "LD (HL),A"
 Case &H78: ToAsm1 = "LD A,B"
 Case &H79: ToAsm1 = "LD A,C"
 Case &H7A: ToAsm1 = "LD A,D"
 Case &H7B: ToAsm1 = "LD A,E"
 Case &H7C: ToAsm1 = "LD A,H"
 Case &H7D: ToAsm1 = "LD A,L"
 Case &H7E: ToAsm1 = "LD A,(HL)"
 Case &H7F: ToAsm1 = "LD A,A"
 Case &H80: ToAsm1 = "ADD A,B"
 Case &H81: ToAsm1 = "ADD A,C"
 Case &H82: ToAsm1 = "ADD A,D"
 Case &H83: ToAsm1 = "ADD A,E"
 Case &H84: ToAsm1 = "ADD A,H"
 Case &H85: ToAsm1 = "ADD A,L"
 Case &H86: ToAsm1 = "ADD A,(HL)"
 Case &H87: ToAsm1 = "ADD A,A"
 Case &H88: ToAsm1 = "ADC A,B"
 Case &H89: ToAsm1 = "ADC A,C"
 Case &H8A: ToAsm1 = "ADC A,D"
 Case &H8B: ToAsm1 = "ADC A,E"
 Case &H8C: ToAsm1 = "ADC A,H"
 Case &H8D: ToAsm1 = "ADC A,L"
 Case &H8E: ToAsm1 = "ADC A,(HL)"
 Case &H8F: ToAsm1 = "ADC A,A"
 Case &H90: ToAsm1 = "SUB A,B"
 Case &H91: ToAsm1 = "SUB A,C"
 Case &H92: ToAsm1 = "SUB A,D"
 Case &H93: ToAsm1 = "SUB A,E"
 Case &H94: ToAsm1 = "SUB A,H"
 Case &H95: ToAsm1 = "SUB A,L"
 Case &H96: ToAsm1 = "SUB A,(HL)"
 Case &H97: ToAsm1 = "SUB A,A"
 Case &H98: ToAsm1 = "SBC A,B"
 Case &H99: ToAsm1 = "SBC A,C"
 Case &H9A: ToAsm1 = "SBC A,D"
 Case &H9B: ToAsm1 = "SBC A,E"
 Case &H9C: ToAsm1 = "SBC A,H"
 Case &H9D: ToAsm1 = "SBC A,L"
 Case &H9E: ToAsm1 = "SBC A,(HL)"
 Case &H9F: ToAsm1 = "SBC A,A"
 Case &HA0: ToAsm1 = "AND B"
 Case &HA1: ToAsm1 = "AND C"
 Case &HA2: ToAsm1 = "AND D"
 Case &HA3: ToAsm1 = "AND E"
 Case &HA4: ToAsm1 = "AND H"
 Case &HA5: ToAsm1 = "AND L"
 Case &HA6: ToAsm1 = "AND (HL)"
 Case &HA7: ToAsm1 = "AND A"
 Case &HA8: ToAsm1 = "XOR B"
 Case &HA9: ToAsm1 = "XOR C"
 Case &HAA: ToAsm1 = "XOR D"
 Case &HAB: ToAsm1 = "XOR E"
 Case &HAC: ToAsm1 = "XOR H"
 Case &HAD: ToAsm1 = "XOR L"
 Case &HAE: ToAsm1 = "XOR (HL)"
 Case &HAF: ToAsm1 = "XOR A"
 Case &HB0: ToAsm1 = "OR B"
 Case &HB1: ToAsm1 = "OR C"
 Case &HB2: ToAsm1 = "OR D"
 Case &HB3: ToAsm1 = "OR E"
 Case &HB4: ToAsm1 = "OR H"
 Case &HB5: ToAsm1 = "OR L"
 Case &HB6: ToAsm1 = "OR (HL)"
 Case &HB7: ToAsm1 = "OR A"
 Case &HB8: ToAsm1 = "CP B"
 Case &HB9: ToAsm1 = "CP C"
 Case &HBA: ToAsm1 = "CP D"
 Case &HBB: ToAsm1 = "CP E"
 Case &HBC: ToAsm1 = "CP H"
 Case &HBD: ToAsm1 = "CP L"
 Case &HBE: ToAsm1 = "CP (HL)"
 Case &HBF: ToAsm1 = "CP A"
 Case &HC0: ToAsm1 = "RET NZ"
 Case &HC1: ToAsm1 = "POP BC"
 Case &HC2: ToAsm1 = "JP NZ," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HC3: ToAsm1 = "JP " & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HC4: ToAsm1 = "CALL NZ," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HC5: ToAsm1 = "PUSH BC"
 Case &HC6: ToAsm1 = "ADD A," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HC7: ToAsm1 = "RST 00H"
 Case &HC8: ToAsm1 = "RET Z"
 Case &HC9: ToAsm1 = "RET"
 Case &HCA: ToAsm1 = "JP Z," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HCB: ToAsm1 = "CB "
 Case &HCC: ToAsm1 = "CALL Z," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HCD: ToAsm1 = "CALL " & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HCE: ToAsm1 = "ADC A," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HCF: ToAsm1 = "RST 08H"
 Case &HD0: ToAsm1 = "RET NC"
 Case &HD1: ToAsm1 = "POP DE"
 Case &HD2: ToAsm1 = "JP NC," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HD3: ToAsm1 = "Invalid Opcode"
 Case &HD4: ToAsm1 = "CALL NC," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HD5: ToAsm1 = "PUSH DE"
 Case &HD6: ToAsm1 = "SUB " & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HD7: ToAsm1 = "RST 10H"
 Case &HD8: ToAsm1 = "RET C"
 Case &HD9: ToAsm1 = "RETI"
 Case &HDA: ToAsm1 = "JP C," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HDB: ToAsm1 = "Invalid Opcode"
 Case &HDC: ToAsm1 = "CALL C," & Hex(readM(adr) + readM(adr + 1) * 256) & "H": adr = adr + 1: adr = adr + 1
 Case &HDD: ToAsm1 = "Invalid Opcode"
 Case &HDE: ToAsm1 = "SBC A," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HDF: ToAsm1 = "RST 18H"
 Case &HE0: ToAsm1 = "LDH (" & Hex(readM(adr)) & "H), A": adr = adr + 1
 Case &HE1: ToAsm1 = "POP HL"
 Case &HE2: ToAsm1 = "LDH (C), A"
 Case &HE3: ToAsm1 = "Invalid Opcode"
 Case &HE4: ToAsm1 = "Invalid Opcode"
 Case &HE5: ToAsm1 = "PUSH HL"
 Case &HE6: ToAsm1 = "AND " & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HE7: ToAsm1 = "RST 20H"
 Case &HE8: ToAsm1 = "ADD SP," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HE9: ToAsm1 = "JP HL"
 Case &HEA: ToAsm1 = "LD (" & Hex(readM(adr) + readM(adr + 1) * 256) & "H), A": adr = adr + 1: adr = adr + 1
 Case &HEB: ToAsm1 = "Invalid Opcode"
 Case &HEC: ToAsm1 = "Invalid Opcode"
 Case &HED: ToAsm1 = "Invalid Opcode"
 Case &HEE: ToAsm1 = "XOR " & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HEF: ToAsm1 = "RST 28H"
 Case &HF0: ToAsm1 = "LDH A,(" & Hex(readM(adr)) & "H)": adr = adr + 1
 Case &HF1: ToAsm1 = "POP AF"
 Case &HF2: ToAsm1 = "Invalid Opcode"
 Case &HF3: ToAsm1 = "DI"
 Case &HF4: ToAsm1 = "Invalid Opcode"
 Case &HF5: ToAsm1 = "PUSH AF"
 Case &HF6: ToAsm1 = "OR " & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HF7: ToAsm1 = "RST 30H"
 Case &HF8: ToAsm1 = "LDHL SP," & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HF9: ToAsm1 = "LD SP,HL"
 Case &HFA: ToAsm1 = "LD A,(" & Hex(readM(adr) + readM(adr + 1) * 256) & "H)": adr = adr + 1: adr = adr + 1
 Case &HFB: ToAsm1 = "EI"
 Case &HFC: ToAsm1 = "Invalid Opcode"
 Case &HFD: ToAsm1 = "Invalid Opcode"
 Case &HFE: ToAsm1 = "CP " & Hex(readM(adr)) & "H": adr = adr + 1
 Case &HFF: ToAsm1 = "RST 38H"
 End Select
 ToAsm1 = pre & ToAsm1
End Function
Sub dbinf(ByVal lon As Byte)
'If lon = &HCB Then
'lon = readM(PC)
'ic(lon + 256) = ic(lon + 256) + 1
'Else
Debug.Print ToAsm1(lon, 0 + PC)
'End If
End Sub
Sub DebuggerDiss(ByVal adr As Long)
If frmDebugger.w Then
Dim sa As Long
If adr <> lasdr Then
Dim i As Long, li As Long
frmDebugger.lstDiss.Clear
#If 0 Then
sa = adr ' save adr
ToAsm2 readM(adr), adr 'get one inst
adr = adr - 1 'dec one
For i = 0 To 5 ' get prev six
If adr < 65535 And adr > 0 Then frmDebugger.lstDiss.AddItem ToAsm2(readM(adr), adr)
adr = adr - 1
Next i
adr = sa
li = frmDebugger.lstDiss.ListCount - 1
#End If
If adr < 65535 And adr > 0 Then frmDebugger.lstDiss.AddItem ToAsm1(readM(adr), adr)
For i = 0 To 5
If adr < 65535 And adr > 0 Then frmDebugger.lstDiss.AddItem ToAsm1(readM(adr), adr)
Next i
frmDebugger.lstDiss.ListIndex = li
lasdr = adr
End If
End If
End Sub
