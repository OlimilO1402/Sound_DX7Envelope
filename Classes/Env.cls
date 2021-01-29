VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Env"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_params()    As Long '0-3: Level; 4-7: Rate;
Private m_level       As Long
Private m_ix          As Long
Private m_i           As Long
Private m_down        As Boolean
Private m_rising      As Boolean
Private m_targetlevel As Long
Private m_qr          As Long

Private outputlevel() As Long
Private envmask()     As Long

Public NoteOff_X As Long
Public NoteOff_Y As Long

Sub Init()
    outputlevel = CLngArr(Array(0, 5, 9, 13, 17, 20, 23, 25, 27, 29, _
                                31, 33, 35, 37, 39, 41, 42, 43, 45, 46, _
                                48, 49, 50, 51, 52, 53, 54, 55, 56, 57, _
                                58, 59, 60, 61, 62, 63, 64, 65, 66, 67, _
                                68, 69, 70, 71, 72, 73, 74, 75, 76, 77, _
                                78, 79, 80, 81, 82, 83, 84, 85, 86, 87, _
                                88, 89, 90, 91, 92, 93, 94, 95, 96, 97, _
                                98, 99, 100, 101, 102, 103, 104, 105, 106, 107, _
                                108, 109, 110, 111, 112, 113, 114, 115, 116, 117, _
                                118, 119, 120, 121, 122, 123, 124, 125, 126, 127))
    envmask = CLngArr2(3, 7, Array(0, 1, 0, 1, 0, 1, 0, 1, _
                                   0, 1, 0, 1, 0, 1, 1, 1, _
                                   0, 1, 1, 1, 0, 1, 1, 1, _
                                   0, 1, 1, 1, 1, 1, 1, 1))
    
End Sub

Private Sub Class_Initialize()
    m_down = True
    Init
End Sub

Friend Sub New_(params() As Long)
    m_params = params
    Reset
End Sub
Friend Function New_Env(params() As Long) As Env
    Set New_Env = New Env
    New_Env.New_ params
End Function

Private Sub Reset()
    m_level = 0
    m_ix = 0
    m_i = 0
    m_down = True
    advance 0
End Sub

Public Function getsample() As Long
    Dim lev As Long
    If (envenable(m_i, m_qr) And (m_ix < 3 Or (m_ix < 4 And Not m_down))) Then
        If (m_rising) Then
            lev = attackstep(m_level, m_i, m_qr)
            'Debug.Print lev 'console.Log (lev)
            If (lev >= m_targetlevel) Then
                lev = m_targetlevel
                Me.advance m_ix + 1
            End If
            m_level = lev
        Else
            lev = decaystep(m_level, m_i, m_qr)
            If (lev <= m_targetlevel) Then
                lev = m_targetlevel
                Me.advance m_ix + 1
            End If
            m_level = lev
        End If
    End If
    RecordNoteOff
    m_i = m_i + 1
    getsample = m_level
End Function
Sub RecordNoteOff()
    If m_ix = 3 And m_down And Not m_rising Then
        NoteOff_X = m_i
        NoteOff_Y = m_level
        'Debug.Print "RecordNoteOff"
    End If
End Sub
Public Function getsample2() As Long
    Dim lev As Long
    If (envenable(m_i, m_qr) And (m_ix < 3 Or (m_ix < 4 And Not m_down))) Then
        If (m_rising) Then
            lev = attackstep(m_level, m_i, m_qr)
            'Debug.Print lev 'console.Log (lev)
            If (lev >= m_targetlevel) Then
                lev = m_targetlevel
                Me.advance2 m_ix + 1
            End If
            m_level = lev
        Else
            lev = decaystep(m_level, m_i, m_qr)
            If (lev <= m_targetlevel) Then
                lev = m_targetlevel
                Me.advance2 m_ix + 1
            End If
            m_level = lev
        End If
    End If
    m_i = m_i + 1
    getsample2 = m_level
End Function

Public Sub advance(ByVal newix As Long)
    m_ix = newix
    If (m_ix < 4) Then
        Dim newlevel    As Long: newlevel = m_params(m_ix)
        Dim scaledlevel As Long: scaledlevel = Max(0, ShL(outputlevel(newlevel), 5) - 224)
        m_targetlevel = scaledlevel
        m_rising = (m_targetlevel - m_level) > 0
        Dim rate_scaling As Long ': rate_scaling = 0
        m_qr = Min(63, rate_scaling + ShR((m_params(m_ix + 4) * 41), 6))
    End If
       Debug.Print "advance ix=" & m_ix & ", qr=" & m_qr & ", target=" & m_targetlevel & ", rising=" & m_rising & ", down=" & m_down
    '//console.log("advance ix="+this.ix+ ", qr="+this.qr+ ", target="+this.targetlevel+ ", rising="+this.rising);
End Sub
Public Sub advance2(ByVal newix As Long)
    m_ix = newix
    If (m_ix < 4) Then
        Dim newlevel    As Long: newlevel = m_params(m_ix)
        Dim scaledlevel As Long: scaledlevel = Max(0, outputlevel(newlevel))
        m_targetlevel = scaledlevel
        m_rising = (m_targetlevel - m_level) > 0
        Dim rate_scaling As Long ': rate_scaling = 0
        m_qr = Min(63, rate_scaling + ShR((m_params(m_ix + 4) * 41), 6))
    End If
    '//console.log("advance ix="+this.ix+", qr="+this.qr+", target="+this.targetlevel+", rising="+this.rising);
End Sub

Public Function keyup()
    m_down = False
    Me.advance 3
End Function

Private Function envenable(ByVal i As Long, ByVal qr As Long) As Boolean
    Dim shift As Long: shift = ShR(qr, 2) - 11 'shift = (qr >> 2) - 11
    If (shift < 0) Then
        Dim sm As Long: sm = ShL(1, -shift) - 1  '(1 << -shift) - 1
        If ((i And sm) <> sm) Then envenable = False: Exit Function
        i = ShR(i, -shift)
    End If
    envenable = envmask(qr And 3, i And 7) <> 0
End Function

Private Function attackstep(ByVal lev As Long, ByVal i As Long, ByVal qr As Long) As Long
    Dim shift As Long: shift = ShR(qr, 2) - 11 'shift = (qr >> 2) - 11
    If (Not envenable(i, qr)) Then attackstep = lev: Exit Function
    Dim slope As Long: slope = 17 - ShR(lev, 8)  '(lev >> 8);
    lev = lev + ShL(slope, Max(shift, 0)) 'lev += slope << Math.max(shift, 0)
    attackstep = lev
End Function

Private Function decaystep(ByVal lev As Long, ByVal i As Long, ByVal qr As Long) As Long
    Dim shift As Long: shift = ShR(qr, 2) - 11  'shift = ShR(qr, 2) - 11
    If (Not envenable(i, qr)) Then decaystep = lev: Exit Function
    lev = lev - ShL(1, Max(shift, 0))  'lev -= 1 << Math.max(shift, 0)
    decaystep = lev
End Function

