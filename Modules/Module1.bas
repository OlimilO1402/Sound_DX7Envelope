Attribute VB_Name = "Module1"
Option Explicit
Public Type pointapi
    X As Long
    Y As Long
End Type
Public Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As pointapi, ByVal nCount As Long) As Long


Public Function New_Env(param() As Long) As Env
    Set New_Env = New Env
    New_Env.New_ param
End Function


'VBHelper
Public Function Max(v1, v2)
    If v1 > v2 Then Max = v1 Else Max = v2
End Function
Public Function Min(v1, v2)
    If v1 < v2 Then Min = v1 Else Min = v2
End Function

Public Function ShR(v As Long, s As Long) As Long
    ShR = v \ 2 ^ s
End Function
Public Function ShL(v As Long, s As Long) As Long
    ShL = v * 2 ^ s
End Function

Public Function CLngArr(vArr) As Long()
    Dim u As Long: u = UBound(vArr)
    ReDim lngarr(0 To u) As Long
    Dim i As Long: For i = 0 To u: lngarr(i) = CLng(vArr(i)): Next
    CLngArr = lngarr
End Function
Public Function CLngArr2(u1 As Long, u2 As Long, vArr) As Long()
    ReDim lngarr(0 To u1, 0 To u2) As Long
    Dim i As Long, j As Long, c As Long
    For i = 0 To u1
        For j = 0 To u2
            lngarr(i, j) = CLng(vArr(c))
            c = c + 1
        Next
    Next
    CLngArr2 = lngarr
End Function
'
'Public Sub Push(vArr(), data)
'    On Error Resume Next
'    Dim c: c = UBound(vArr) + 1
'    ReDim Preserve vArr(0 To c)
'    vArr(c) = data
'End Sub
'
'Public Function Push(arr() As Long, data As Long, ByVal i As Long)
'    arr(i) = data
'    Push = i + 1
'End Function
'Sub Push(ByRef vArr, data)
'    On Error Resume Next
'    Dim u1 As Long: u1 = UBound(vArr)
'    Dim u2 As Long, b As Boolean: b = IsArray(data, u2)
'    Dim u3 As Long: u3 = u1 + u2
'    ReDim Preserve vArr(0 To u3)
'    If b Then
'        Dim i As Long
'        For i = 0 To UBound(data)
'            vArr(i) = data(i)
'        Next
'    Else
'        vArr(u3) = data
'    End If
'End Sub
'
'Function IsArray(vArr, u) As Boolean
'    If IsEmpty(vArr) Then Exit Function
'    IsArray = (VarType(vArr) And vbArray) = vbArray
'    If IsArray Then u = UBound(vArr) Else u = 1
'End Function
'
'Sub Push2(ByRef vArr, u1 As Long, u2 As Long, data)
'    On Error Resume Next
'    ReDim Preserve vArr(0 To u1, 0 To u2)
'    Dim i As Long, j As Long, c As Long
'    For i = 0 To u1
'        For j = 0 To u2
'            vArr(i, j) = data(c): c = c + 1
'        Next
'    Next
'End Sub
'
'Sub Push2(ByRef vArr, u1 As Long, u2 As Long, data)
'    On Error Resume Next
'    'Dim u1 As Long: u1 = UBound(vArr)
'    'Dim u2 As Long, b As Boolean: b = IsArray(data, u2)
'    'Dim u3 As Long: u3 = u1 + u2
'    'ReDim Preserve vArr(0 To u3)
'    ReDim Preserve vArr(0 To u1, 0 To u2)
'    'If b Then
'    Dim i As Long, j As Long, c As Long
'    For i = 0 To u1
'        For j = 0 To u2
'            vArr(i, j) = data(c)
'            c = c + 1
'        Next
'    Next
'    'Else
'    '    vArr(u3) = data
'    'End If
'End Sub

'function draw_outputlevel(chart) {
'    var rawdata = [['value', 'gain']];
'    for (var i = 0; i < outputlevel.length; i++) {
'        var ol = outputlevel[i];
'        var db = 20 * Math.log(2) / Math.log(10) * (ol - 127) / 8;
'        rawdata.push([i, db]);
'    }
'    var data = google.visualization.arrayToDataTable(rawdata);
'    chart.draw(data, {title: 'Output level to dB scaling', hAxis: {
'                                                                    title: 'Output level value (0-99)'
'    }, vAxis: {
'                title: dB'
'    }});
'}
'function draw_rate(chart) {
'    var rawdata = [['value', 'rate']];
'    for (var i = 0; i < outputlevel.length; i++) {
'        var qr = i * 41 / 64;
'        var samplerate = 49096;
'        var baserate = samplerate / (1<<20) * 20 * Math.log(2) / Math.log(10);
'        var rate = baserate * (1 << (qr >> 2)) * (1 + .25 * (qr & 3));
'        rawdata.push([i, rate]);
'    }
'    var data = google.visualization.arrayToDataTable(rawdata);
'    chart.draw(data, {title: 'Rate values to actual level change', hAxis: {
'title:         'Rate value (0-99)'
'    }, vAxis: {
'title:         'dB/s',
'        logScale: true
'    }});
'}

