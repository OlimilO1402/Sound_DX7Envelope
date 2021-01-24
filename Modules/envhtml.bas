Attribute VB_Name = "envhtml"
Option Explicit

Sub LoadEnvHtml(params() As Long, ByVal nsamp As Long)
    Dim s As String
    Dim pre As String, par As String, post As String
    s = s & "<html>" & vbCrLf
    s = s & "  <head>" & vbCrLf
    s = s & "    <script type=""text/javascript"" src=""https://www.google.com/jsapi""></script>" & vbCrLf
    s = s & "    <script type=""text/javascript"">" & vbCrLf
    s = s & "      google.load(""visualization"", ""1"", {packages:[""corechart""]});" & vbCrLf
    s = s & "      google.setOnLoadCallback(drawChart);" & vbCrLf
    s = s & "      function drawChart() {" & vbCrLf
    s = s & "        var options = {" & vbCrLf
    s = s & "            title: 'Wave form qr=60'" & vbCrLf
    s = s & "        };" & vbCrLf
    s = s & "        var chart = new google.visualization.LineChart(document.getElementById('chart_div'));" & vbCrLf
    s = s & "        //chart.draw(data, options);" & vbCrLf
    s = s & "        function redraw() {" & vbCrLf
    s = s & "          var params = [];" & vbCrLf
    s = s & "          for (var i = 0; i < 8; i++) {" & vbCrLf
    s = s & "            params.push(parseInt(document.getElementById(""param"" + i).value));" & vbCrLf
    s = s & "          }" & vbCrLf
    s = s & "          var nsamp = parseInt(document.getElementById(""nsamp"").value);" & vbCrLf
    s = s & "          var rawdata = envdata(params, nsamp);" & vbCrLf
    s = s & "          var data = google.visualization.arrayToDataTable(rawdata);" & vbCrLf
    s = s & "          chart.draw(data, {title: 'Envelope'});" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "        for (var i = 0; i < 8; i++) {" & vbCrLf
    s = s & "          document.getElementById(""param"" + i).addEventListener(""change"", redraw);" & vbCrLf
    s = s & "        }" & vbCrLf
    s = s & "        document.getElementById(""nsamp"").addEventListener(""change"", redraw);" & vbCrLf
    s = s & "        redraw();" & vbCrLf
    s = s & "      }" & vbCrLf
    s = s & "    </script>" & vbCrLf
    s = s & "  </head>" & vbCrLf
    s = s & "  <body>" & vbCrLf
    pre = s: s = ""
    s = s & "    Level 1: <input id=""param0"" type=""text"" value=""" & params(0) & """>  Rate 1: <input id=""param4"" type=""text"" value=""" & params(4) & """><br/>" & vbCrLf
    s = s & "    Level 2: <input id=""param1"" type=""text"" value=""" & params(1) & """>  Rate 2: <input id=""param5"" type=""text"" value=""" & params(5) & """><br/>" & vbCrLf
    s = s & "    Level 3: <input id=""param2"" type=""text"" value=""" & params(2) & """>  Rate 3: <input id=""param6"" type=""text"" value=""" & params(6) & """><br/>" & vbCrLf
    s = s & "    Level 4: <input id=""param3"" type=""text"" value=""" & params(3) & """>  Rate 4: <input id=""param7"" type=""text"" value=""" & params(7) & """><br/>" & vbCrLf
    par = s: s = ""
    s = s & "    Number of samples: <input id=""nsamp"" type=""text"" value=""" & nsamp & """><br/>" & vbCrLf
    s = s & "    <div id=""chart_div"" style=""width: 900px; height: 500px;""></div>" & vbCrLf
    s = s & "  </body>" & vbCrLf
    s = s & "  <script>" & vbCrLf
    s = s & "var envmask = [[0, 1, 0, 1, 0, 1, 0, 1]," & vbCrLf
    s = s & "    [0, 1, 0, 1, 0, 1, 1, 1]," & vbCrLf
    s = s & "    [0, 1, 1, 1, 0, 1, 1, 1]," & vbCrLf
    s = s & "    [0, 1, 1, 1, 1, 1, 1, 1]];" & vbCrLf
    s = s & "var outputlevel = [0, 5, 9, 13, 17, 20, 23, 25, 27, 29, 31, 33, 35, 37, 39," & vbCrLf
    s = s & "41, 42, 43, 45, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61," & vbCrLf
    s = s & "62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80," & vbCrLf
    s = s & "81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99," & vbCrLf
    s = s & "100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114," & vbCrLf
    s = s & "115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127];" & vbCrLf
    s = s & "function envenable(i, qr) {" & vbCrLf
    s = s & "  var shift = (qr >> 2) - 11;" & vbCrLf
    s = s & "  if (shift < 0) {" & vbCrLf
    s = s & "    var sm = (1 << -shift) - 1;" & vbCrLf
    s = s & "    if ((i & sm) != sm) return false;" & vbCrLf
    s = s & "    i >>= -shift;" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  return envmask[qr & 3][i & 7] != 0;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function attackstep(lev, i, qr) {" & vbCrLf
    s = s & "  var shift = (qr >> 2) - 11;" & vbCrLf
    s = s & "  if (!envenable(i, qr)) return lev;" & vbCrLf
    s = s & "  var slope = 17 - (lev >> 8);" & vbCrLf
    s = s & "  lev += slope << Math.max(shift, 0);" & vbCrLf
    s = s & "  return lev;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function decaystep(lev, i, qr) {" & vbCrLf
    s = s & "  var shift = (qr >> 2) - 11;" & vbCrLf
    s = s & "  if (!envenable(i, qr)) return lev;" & vbCrLf
    s = s & "  lev -= 1 << Math.max(shift, 0);" & vbCrLf
    s = s & "  return lev;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function Env(params) {" & vbCrLf
    s = s & "  this.params = params;" & vbCrLf
    s = s & "  this.level = 0;" & vbCrLf
    s = s & "  this.ix = 0;" & vbCrLf
    s = s & "  this.i = 0;" & vbCrLf
    s = s & "  this.down = true;" & vbCrLf
    s = s & "  this.advance(0);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "Env.prototype.getsample = function() {" & vbCrLf
    s = s & "  if (envenable(this.i, this.qr) && (this.ix < 3 || (this.ix < 4 && !this.down))) {" & vbCrLf
    s = s & "    if (this.rising) {" & vbCrLf
    s = s & "      var lev = attackstep(this.level, this.i, this.qr);" & vbCrLf
    s = s & "      console.log(lev);" & vbCrLf
    s = s & "      if (lev >= this.targetlevel) {" & vbCrLf
    s = s & "        lev = this.targetlevel;" & vbCrLf
    s = s & "        this.advance(this.ix + 1);" & vbCrLf
    s = s & "      }" & vbCrLf
    s = s & "      this.level = lev;" & vbCrLf
    s = s & "    } else {" & vbCrLf
    s = s & "      var lev = decaystep(this.level, this.i, this.qr);" & vbCrLf
    s = s & "      if (lev <= this.targetlevel) {" & vbCrLf
    s = s & "        lev = this.targetlevel;" & vbCrLf
    s = s & "        this.advance(this.ix + 1);" & vbCrLf
    s = s & "      }" & vbCrLf
    s = s & "      this.level = lev;" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  this.i++;" & vbCrLf
    s = s & "  return this.level;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "Env.prototype.advance = function(newix) {" & vbCrLf
    s = s & "  this.ix = newix;" & vbCrLf
    s = s & "  if (this.ix < 4) {" & vbCrLf
    s = s & "    var newlevel = this.params[this.ix];" & vbCrLf
    s = s & "    var scaledlevel = Math.max(0, (outputlevel[newlevel] << 5) - 224);" & vbCrLf
    s = s & "    this.targetlevel = scaledlevel;" & vbCrLf
    s = s & "    this.rising = (this.targetlevel - this.level) > 0;" & vbCrLf
    s = s & "    var rate_scaling = 0;" & vbCrLf
    s = s & "    this.qr = Math.min(63, rate_scaling + ((this.params[this.ix + 4] * 41) >> 6));" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  //console.log(""advance ix=""+this.ix+"", qr=""+this.qr+"", target=""+this.targetlevel+"", rising=""+this.rising);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "Env.prototype.keyup = function() {" & vbCrLf
    s = s & "  this.down = false;" & vbCrLf
    s = s & "  this.advance(3);" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function attackdata(qr) {" & vbCrLf
    s = s & "  var result = [['samp', 'env']];" & vbCrLf
    s = s & "  var i = 0;" & vbCrLf
    s = s & "  var count = 0;" & vbCrLf
    s = s & "  var lev = 1716;" & vbCrLf
    s = s & "  while (true) {" & vbCrLf
    s = s & "    result.push([i, lev]);" & vbCrLf
    s = s & "    lev = attackstep(lev, i, qr);" & vbCrLf
    s = s & "    lev = Math.min(lev, 15 << 8);" & vbCrLf
    s = s & "    if (lev >= 15 << 8) {" & vbCrLf
    s = s & "      count++;" & vbCrLf
    s = s & "      if (count > 100) break;" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    i++;" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  return result;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "function envdata(params, nsamp) {" & vbCrLf
    s = s & "  console.log(nsamp);" & vbCrLf
    s = s & "  var result = [['samp', 'env']];" & vbCrLf
    s = s & "  var env = new Env(params);" & vbCrLf
    s = s & "  for (var i = 0; i < nsamp; i++) {" & vbCrLf
    s = s & "    if (i == 3 * nsamp / 4) {" & vbCrLf
    s = s & "      env.keyup();" & vbCrLf
    s = s & "    }" & vbCrLf
    s = s & "    result.push([i, env.getsample()]);" & vbCrLf
    s = s & "  }" & vbCrLf
    s = s & "  return result;" & vbCrLf
    s = s & "}" & vbCrLf
    s = s & "  </script>" & vbCrLf
    s = s & "</html>" & vbCrLf
    post = s
    s = WriteToFile(pre & par & post, App.Path & "\env_temp.html")
    Shell "explorer.exe " & s
End Sub

Function WriteToFile(filedata As String, PFN As String) As String
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary As FNr
    Put FNr, , filedata ', , FNr
    Debug.Print PFN
    WriteToFile = PFN
Finally:
    Close FNr
End Function
