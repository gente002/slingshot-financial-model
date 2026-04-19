Attribute VB_Name = "KernelRandom"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelRandom.bas
' Purpose: Seeded pseudo-random number generator (Mersenne Twister MT19937)
'          for reproducible stochastic computation. Pure VBA, no dependencies.
' =============================================================================

' MT19937 constants
Private Const MT_N As Long = 624
Private Const MT_M As Long = 397

' MT state array (624 longs) and index
Private m_mt(0 To 623) As Long
Private m_mti As Long

' Stored seed for manifest/reproducibility
Private m_seed As Long
Private m_initialized As Boolean
Private m_callCount As Long


' =============================================================================
' InitSeed
' Initialize the MT state array from the given seed.
' =============================================================================
Public Sub InitSeed(seed As Long)
    m_seed = seed
    m_mt(0) = seed
    Dim i As Long
    For i = 1 To MT_N - 1
        ' m_mt(i) = 1812433253 * (m_mt(i-1) XOR (m_mt(i-1) >> 30)) + i
        Dim prev As Long
        prev = m_mt(i - 1)
        Dim shifted As Long
        shifted = UnsignedShiftRight(prev, 30)
        Dim xored As Long
        xored = prev Xor shifted
        m_mt(i) = AddLong(MultiplyLong(1812433253, xored), i)
    Next i
    m_mti = MT_N
    m_initialized = True
    m_callCount = 0
    KernelConfig.LogError SEV_INFO, "KernelRandom", "I-500", _
        "PRNG initialized with seed " & seed, ""
End Sub


' =============================================================================
' AutoSeed
' Generate a seed from Timer and Rnd for entropy, then call InitSeed.
' =============================================================================
Public Sub AutoSeed()
    Randomize Timer
    Dim autoSeedVal As Long
    autoSeedVal = CLng((Timer * 1000) Mod 2147483647) Xor CLng(Rnd * 100000)
    If autoSeedVal = 0 Then autoSeedVal = 5489
    KernelConfig.LogError SEV_INFO, "KernelRandom", "I-501", _
        "PRNG auto-seeded with value " & autoSeedVal, _
        "To reproduce this run, use InitSeed " & autoSeedVal
    InitSeed autoSeedVal
End Sub


' =============================================================================
' NextRandom
' Returns next Uniform(0,1) value from the MT sequence.
' =============================================================================
Public Function NextRandom() As Double
    If Not m_initialized Then
        KernelConfig.LogError SEV_WARN, "KernelRandom", "W-500", _
            "NextRandom called before InitSeed; auto-seeding", ""
        AutoSeed
    End If

    Dim y As Long

    ' Generate MT_N words at one time if needed
    If m_mti >= MT_N Then
        GenerateNumbers
    End If

    y = m_mt(m_mti)
    m_mti = m_mti + 1

    ' Tempering
    y = y Xor UnsignedShiftRight(y, 11)
    y = y Xor (ShiftLeft(y, 7) And &H9D2C5680)
    y = y Xor (ShiftLeft(y, 15) And &HEFC60000)
    y = y Xor UnsignedShiftRight(y, 18)

    m_callCount = m_callCount + 1

    ' Convert to [0, 1) double
    ' Use unsigned interpretation: if y < 0, add 2^32
    Dim dblVal As Double
    If y < 0 Then
        dblVal = CDbl(y) + 4294967296#
    Else
        dblVal = CDbl(y)
    End If
    NextRandom = dblVal / 4294967296#
End Function


' =============================================================================
' NextRandomRange
' Returns NextRandom() scaled to [low, high).
' =============================================================================
Public Function NextRandomRange(low As Double, high As Double) As Double
    NextRandomRange = low + NextRandom() * (high - low)
End Function


' =============================================================================
' GetSeed
' Returns the current seed value (for savepoint manifests).
' =============================================================================
Public Function GetSeed() As Long
    GetSeed = m_seed
End Function


' =============================================================================
' GetCallCount
' Returns how many times NextRandom has been called since InitSeed.
' =============================================================================
Public Function GetCallCount() As Long
    GetCallCount = m_callCount
End Function


' =============================================================================
' ResetToSeed
' Re-initializes with the stored seed. Resets call count.
' =============================================================================
Public Sub ResetToSeed()
    InitSeed m_seed
End Sub


' =============================================================================
' IsInitialized
' Returns whether the PRNG has been seeded.
' =============================================================================
Public Function IsInitialized() As Boolean
    IsInitialized = m_initialized
End Function


' =============================================================================
' GenerateNumbers (Private)
' Generates MT_N new values in the state array.
' =============================================================================
Private Sub GenerateNumbers()
    Dim kk As Long
    Dim y As Long

    For kk = 0 To MT_N - MT_M - 1
        y = (m_mt(kk) And &H80000000) Or (m_mt(kk + 1) And &H7FFFFFFF)
        m_mt(kk) = m_mt(kk + MT_M) Xor UnsignedShiftRight(y, 1)
        If (y And 1) <> 0 Then
            m_mt(kk) = m_mt(kk) Xor &H9908B0DF
        End If
    Next kk

    For kk = MT_N - MT_M To MT_N - 2
        y = (m_mt(kk) And &H80000000) Or (m_mt(kk + 1) And &H7FFFFFFF)
        m_mt(kk) = m_mt(kk + (MT_M - MT_N)) Xor UnsignedShiftRight(y, 1)
        If (y And 1) <> 0 Then
            m_mt(kk) = m_mt(kk) Xor &H9908B0DF
        End If
    Next kk

    y = (m_mt(MT_N - 1) And &H80000000) Or (m_mt(0) And &H7FFFFFFF)
    m_mt(MT_N - 1) = m_mt(MT_M - 1) Xor UnsignedShiftRight(y, 1)
    If (y And 1) <> 0 Then
        m_mt(MT_N - 1) = m_mt(MT_N - 1) Xor &H9908B0DF
    End If

    m_mti = 0
End Sub


' =============================================================================
' UnsignedShiftRight (Private)
' Performs unsigned right shift on a VBA Long (signed 32-bit).
' =============================================================================
Private Function UnsignedShiftRight(value As Long, shift As Long) As Long
    If shift = 0 Then
        UnsignedShiftRight = value
        Exit Function
    End If
    If shift >= 32 Then
        UnsignedShiftRight = 0
        Exit Function
    End If

    ' Handle sign bit separately for unsigned behavior
    If value >= 0 Then
        UnsignedShiftRight = value \ (2 ^ shift)
    Else
        ' Clear sign bit, shift, then add shifted sign bit back
        Dim cleared As Long
        cleared = value And &H7FFFFFFF
        UnsignedShiftRight = (cleared \ (2 ^ shift)) Or (2 ^ (31 - shift))
    End If
End Function


' =============================================================================
' ShiftLeft (Private)
' Performs left shift on a VBA Long. Overflow wraps via signed 32-bit.
' =============================================================================
Private Function ShiftLeft(value As Long, shift As Long) As Long
    If shift = 0 Then
        ShiftLeft = value
        Exit Function
    End If
    If shift >= 32 Then
        ShiftLeft = 0
        Exit Function
    End If

    ' Use Double for intermediate to avoid overflow
    Dim dblResult As Double
    Dim dblVal As Double
    If value < 0 Then
        dblVal = CDbl(value) + 4294967296#
    Else
        dblVal = CDbl(value)
    End If

    dblResult = dblVal * (2 ^ shift)

    ' Wrap to 32-bit unsigned range
    dblResult = dblResult - Fix(dblResult / 4294967296#) * 4294967296#

    ' Convert back to signed Long
    If dblResult >= 2147483648# Then
        ShiftLeft = CLng(dblResult - 4294967296#)
    Else
        ShiftLeft = CLng(dblResult)
    End If
End Function


' =============================================================================
' MultiplyLong (Private)
' Multiplies two Longs treating them as unsigned, returns lower 32 bits.
' =============================================================================
Private Function MultiplyLong(a As Long, b As Long) As Long
    Dim dblA As Double
    Dim dblB As Double

    If a < 0 Then
        dblA = CDbl(a) + 4294967296#
    Else
        dblA = CDbl(a)
    End If

    If b < 0 Then
        dblB = CDbl(b) + 4294967296#
    Else
        dblB = CDbl(b)
    End If

    Dim dblResult As Double
    dblResult = dblA * dblB

    ' Take mod 2^32
    dblResult = dblResult - Fix(dblResult / 4294967296#) * 4294967296#

    ' Convert back to signed Long
    If dblResult >= 2147483648# Then
        MultiplyLong = CLng(dblResult - 4294967296#)
    Else
        MultiplyLong = CLng(dblResult)
    End If
End Function


' =============================================================================
' AddLong (Private)
' Adds two Longs treating them as unsigned, returns lower 32 bits.
' =============================================================================
Private Function AddLong(a As Long, b As Long) As Long
    Dim dblA As Double
    Dim dblB As Double

    If a < 0 Then
        dblA = CDbl(a) + 4294967296#
    Else
        dblA = CDbl(a)
    End If

    If b < 0 Then
        dblB = CDbl(b) + 4294967296#
    Else
        dblB = CDbl(b)
    End If

    Dim dblResult As Double
    dblResult = dblA + dblB

    ' Take mod 2^32
    If dblResult >= 4294967296# Then
        dblResult = dblResult - 4294967296#
    End If

    ' Convert back to signed Long
    If dblResult >= 2147483648# Then
        AddLong = CLng(dblResult - 4294967296#)
    Else
        AddLong = CLng(dblResult)
    End If
End Function
