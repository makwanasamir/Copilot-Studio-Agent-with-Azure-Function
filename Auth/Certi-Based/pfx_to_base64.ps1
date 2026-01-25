$data = [Convert]::ToBase64String(
    [IO.File]::ReadAllBytes("C:\Azure\AD-App\Certi\EXO-AppOnly-MessageTrace.pfx")
)
 
Set-Content -Path "C:\Azure\AD-App\Certi\EXO-AppOnly-MessageTrace.txt" -Value $data