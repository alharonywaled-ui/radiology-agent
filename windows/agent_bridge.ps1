param([string]$ConfigPath = "$env:LOCALAPPDATA\RadiologyAgent\config.json")
$ErrorActionPreference='Stop'
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
$Origin='https://alharonywaled-ui.github.io'
$Port=8765
$LogDir=Join-Path $env:LOCALAPPDATA 'RadiologyAgent'
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$Log=Join-Path $LogDir 'bridge.log'
function Record([string]$msg) { Add-Content -LiteralPath $Log -Value "$(Get-Date -Format o) $msg" -Encoding utf8 }
function Reply($stream,[int]$code,$data,[bool]$cors=$false) {
  $body=ConvertTo-Json -InputObject $data -Compress -Depth 15
  $bytes=[Text.Encoding]::UTF8.GetBytes($body)
  $h="HTTP/1.1 $code Response`r`nContent-Type: application/json; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n"
  if($cors){$h+="Access-Control-Allow-Origin: $Origin`r`nAccess-Control-Allow-Methods: POST, OPTIONS`r`nAccess-Control-Allow-Headers: content-type`r`nAccess-Control-Allow-Private-Network: true`r`nVary: Origin`r`n"}
  $header=[Text.Encoding]::ASCII.GetBytes($h+"`r`n")
  $stream.Write($header,0,$header.Length)
  $stream.Write($bytes,0,$bytes.Length)
  $stream.Flush()
}
function Process($client){
 $client.ReceiveTimeout=120000;$client.SendTimeout=120000
 $stream=$client.GetStream()
 $buffer=New-Object 'System.Collections.Generic.List[byte]'
 while($buffer.Count -lt 16384){
  $b=$stream.ReadByte();if($b -lt 0){return}
  $buffer.Add([byte]$b);$n=$buffer.Count
  if($n -ge 4 -and $buffer[$n-4] -eq 13 -and $buffer[$n-3] -eq 10 -and $buffer[$n-2] -eq 13 -and $buffer[$n-1] -eq 10){break}
 }
 if($buffer.Count -ge 16384){return}
 $lines=([Text.Encoding]::ASCII.GetString($buffer.ToArray())) -split "`r`n"
 $first=$lines[0] -split ' ';if($first.Length -lt 2){return}
 $method=$first[0];$path=$first[1];$headers=@{}
 foreach($line in $lines){
   $idx=$line.IndexOf(':')
   if($idx -gt 0){$headers[$line.Substring(0,$idx).ToLowerInvariant()]=$line.Substring($idx+1).Trim()}
 }
 if(([string]$headers['origin']) -cne $Origin){Reply $stream 403 @{error='Origin blocked'};return}
 if($method -eq 'OPTIONS'){Reply $stream 200 @{} $true;return}
 if($method -ne 'POST' -or $path -ne '/review'){Reply $stream 404 @{error='Not found'} $true;return}
 $length=0
 if(![int]::TryParse([string]$headers['content-length'],[ref]$length) -or $length -lt 2 -or $length -gt 85000){Reply $stream 413 @{error='Invalid request length'} $true;return}
 $bytes=New-Object byte[] $length;$count=0
 while($count -lt $length){$read=$stream.Read($bytes,$count,$length-$count);if($read -le 0){throw 'Incomplete request'};$count+=$read}
 $req=[Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
 $report=[string]$req.text
 if([string]::IsNullOrWhiteSpace($report) -or $report.Length -gt 18000){Reply $stream 400 @{error='Invalid report'} $true;return}
 if(!(Test-Path -LiteralPath $ConfigPath)){Reply $stream 500 @{error='Gemini API Key not configured'} $true;return}
 $config=Get-Content -LiteralPath $ConfigPath -Encoding UTF8 -Raw | ConvertFrom-Json
 $key=[string]$config.gemini_api_key
 if([string]::IsNullOrWhiteSpace($key)){Reply $stream 500 @{error='Missing API Key'} $true;return}
 $prompt=@'
Act as a cautious radiology-report proofreader. Identify grammatical errors and clear INTERNAL contradictions grounded exclusively in the report; do not diagnose, fabricate medical facts or present uncertain claims as facts. Return at most 30 items in valid JSON {"issues":[{"quote":"exact substring of report","category":"language or logic","explanation":"brief reason","suggestion":"brief proposed edit"}]}. Quotes must match the original report byte-for-byte where possible. Only classify logic when two statements contradict each other internally. Do not include patient identifiers in explanations.
REPORT:
'@
 $payload=@{contents=@(@{role='user';parts=@(@{text=($prompt+"`n"+$report)})});generationConfig=@{responseMimeType='application/json';temperature=0.1}} | ConvertTo-Json -Depth 14
 try {
  $uri='https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent'
  $response=Invoke-RestMethod -Uri $uri -Method Post -ContentType 'application/json; charset=utf-8' -Headers @{'x-goog-api-key'=$key} -Body ([Text.Encoding]::UTF8.GetBytes($payload)) -TimeoutSec 80
  $raw=[string]$response.candidates[0].content.parts[0].text
  $parsed=$raw | ConvertFrom-Json
  $issues=@()
  foreach($issue in @($parsed.issues)){
    if($issues.Count -ge 30){break}
    $quote=[string]$issue.quote;$category=[string]$issue.category
    if($quote.Length -gt 0 -and $quote.Length -le 300 -and $report.Contains($quote) -and @('language','logic') -ccontains $category){
      $explanation=[string]$issue.explanation
      $issues+=@{quote=$quote;category=$category;explanation=$explanation.Substring(0,[Math]::Min(650,$explanation.Length));suggestion=[string]$issue.suggestion}
    }
  }
  Reply $stream 200 @{issues=@($issues)} $true
 }catch{Record ('Gemini error: '+$_.Exception.GetType().Name);Reply $stream 502 @{error='Gemini failed. Verify API key, network and quota.'} $true}
}
$listener=[Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback,$Port)
try{$listener.Start();Record 'Bridge started'
 while($true){$client=$listener.AcceptTcpClient();try{Process $client}catch{Record ('Request failed: '+$_.Exception.Message)}finally{$client.Close()}}
}finally{$listener.Stop()}
