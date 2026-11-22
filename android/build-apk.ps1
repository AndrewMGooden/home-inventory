# Builds HomeInventory.apk with the Android SDK tools directly (no Gradle).
# Tools and the signing key live in %USERPROFILE%\home-inventory-android.
# Keep release.keystore: Android only installs updates signed with the same key.
$ErrorActionPreference = 'Stop'
$base = "$env:USERPROFILE\home-inventory-android"
$jar = "$base\sdk\platforms\android-35\android.jar"
$bt = "$base\sdk\build-tools\35.0.0"
$env:JAVA_HOME = "$base\jdk"
$env:PATH = "$env:JAVA_HOME\bin;$env:PATH"
$src = $PSScriptRoot
$out = "$base\build"
$versionCode = if ($args[0]) { $args[0] } else { 1 }

function Check($what) { if ($LASTEXITCODE -ne 0) { throw "$what failed" } }

if (Test-Path $out) { Remove-Item -Recurse -Force $out }
New-Item -ItemType Directory "$out\classes", "$out\dex" | Out-Null

& "$bt\aapt2.exe" compile --dir "$src\res" -o "$out\res.zip"; Check 'aapt2 compile'
& "$bt\aapt2.exe" link -o "$out\unsigned.apk" -I $jar --manifest "$src\AndroidManifest.xml" "$out\res.zip" `
  --min-sdk-version 26 --target-sdk-version 35 --version-code $versionCode --version-name "1.$versionCode"; Check 'aapt2 link'

javac -source 8 -target 8 -bootclasspath $jar -nowarn -Xlint:-options -d "$out\classes" `
  (Get-ChildItem "$src\src" -Recurse -Filter *.java).FullName; Check 'javac'
& "$bt\d8.bat" --release --min-api 26 --lib $jar --output "$out\dex" `
  (Get-ChildItem "$out\classes" -Recurse -Filter *.class).FullName; Check 'd8'
jar --update --no-manifest --no-compress --file "$out\unsigned.apk" -C "$out\dex" classes.dex; Check 'adding classes.dex'

& "$bt\zipalign.exe" -p -f 4 "$out\unsigned.apk" "$out\aligned.apk"; Check 'zipalign'

$keystore = "$base\release.keystore"
$passFile = "$base\keystore-password.txt"
if (-not (Test-Path $keystore)) {
  $chars = [char[]]'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'
  -join (1..24 | ForEach-Object { $chars | Get-Random }) | Set-Content -NoNewline -Encoding ascii $passFile
  $env:KS_PASS = Get-Content -Raw $passFile
  keytool -genkeypair -keystore $keystore -alias homeinventory -keyalg RSA -keysize 2048 -validity 10000 `
    -storepass:env KS_PASS -keypass:env KS_PASS -dname "CN=Home Inventory" | Out-Null; Check 'keytool'
}
$env:KS_PASS = Get-Content -Raw $passFile
& "$bt\apksigner.bat" sign --ks $keystore --ks-pass env:KS_PASS --key-pass env:KS_PASS `
  --out "$src\HomeInventory.apk" "$out\aligned.apk"; Check 'apksigner'
& "$bt\apksigner.bat" verify "$src\HomeInventory.apk"; Check 'apksigner verify'
Write-Host "Built $src\HomeInventory.apk"
