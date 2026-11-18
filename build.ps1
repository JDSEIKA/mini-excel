# Compila CalcRÃ¡pido con el compilador de .NET Framework incluido en Windows.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path $csc)) { throw 'No se encontrÃ³ csc.exe de .NET Framework 4.x.' }

New-Item -ItemType Directory -Force -Path (Join-Path $root 'dist') | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $root 'tests\bin') | Out-Null

$common = @(
    '/nologo', '/codepage:65001', '/optimize+', '/platform:anycpu', '/debug-',
    '/r:System.dll', '/r:System.Core.dll', '/r:System.Drawing.dll', '/r:System.Windows.Forms.dll',
    '/r:System.IO.Compression.dll', '/r:System.IO.Compression.FileSystem.dll',
    '/r:System.Xml.dll', '/r:System.Xml.Linq.dll', '/r:System.Security.dll'
)

$src = Get-ChildItem (Join-Path $root 'src\*.cs') |
    Where-Object { $_.Name -ne 'SelfTest.cs' } |
    Sort-Object Name | ForEach-Object { $_.FullName }
$exe = Join-Path $root 'dist\MiniExcel.exe'
$args = @('/target:winexe', ('/out:' + $exe), ('/win32icon:' + (Join-Path $root 'assets\app.ico')),
          ('/win32manifest:' + (Join-Path $root 'assets\app.manifest'))) + $common + $src
Write-Host 'Compilando MiniExcel.exe...'
& $csc $args
if ($LASTEXITCODE -ne 0) { throw "La compilaciÃ³n fallÃ³ (cÃ³digo $LASTEXITCODE)." }

$testFiles = @('Model.cs', 'FormulaEngine.cs', 'Functions.cs', 'ConditionalFormatting.cs', 'TableFormatting.cs', 'FileIO.cs', 'Ole2.cs', 'XlsIO.cs', 'XlsFormula.cs', 'XlsxIO.cs', 'Compatibility.cs', 'Charts.cs', 'SelfTest.cs') |
    ForEach-Object { Join-Path $root ('src\' + $_) }
$testMain = Join-Path $root 'tests\TestMain.cs'
$testExe = Join-Path $root 'tests\bin\CalcRapido.Tests.exe'
$targs = @('/target:exe', ('/out:' + $testExe)) + $common + $testFiles + @($testMain)
Write-Host 'Compilando pruebas del motor...'
& $csc $targs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas no compilaron (cÃ³digo $LASTEXITCODE)." }
Write-Host 'Ejecutando pruebas del motor...'
& $testExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas del motor (cÃ³digo $LASTEXITCODE)." }

$basicTestFiles = @('Model.cs', 'FormulaEngine.cs', 'Functions.cs', 'ConditionalFormatting.cs', 'TableFormatting.cs', 'Charts.cs') |
    ForEach-Object { Join-Path $root ('src\' + $_) }
$basicTestExe = Join-Path $root 'tests\bin\BasicCompatProbe.exe'
$bargs = @('/target:exe', ('/out:' + $basicTestExe), '/main:BasicCompatProbe') + $common +
    $basicTestFiles + @((Join-Path $root 'tests\BasicCompatProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de compatibilidad bÃ¡sica...'
& $csc $bargs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de compatibilidad no compilaron (cÃ³digo $LASTEXITCODE)." }
& $basicTestExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de compatibilidad (cÃ³digo $LASTEXITCODE)." }

$modernExe = Join-Path $root 'tests\bin\ModernFunctionsProbe.exe'
$margs = @('/target:exe', ('/out:' + $modernExe), '/main:ModernFunctionsProbe') + $common +
    $basicTestFiles + @((Join-Path $root 'tests\ModernFunctionsProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de funciones modernas...'
& $csc $margs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas modernas no compilaron (cÃ³digo $LASTEXITCODE)." }
& $modernExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas modernas (cÃ³digo $LASTEXITCODE)." }

$xlsCompatExe = Join-Path $root 'tests\bin\XlsNamesSharedTest.exe'
$xlsCompatArgs = @('/target:exe', ('/out:' + $xlsCompatExe), '/main:XlsNamesSharedTest') + $common +
    $src + @((Join-Path $root 'tests\XlsNamesSharedTest.cs'))
Write-Host 'Compilando y ejecutando pruebas de nombres/fÃ³rmulas .xls...'
& $csc $xlsCompatArgs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas .xls no compilaron (cÃ³digo $LASTEXITCODE)." }
& $xlsCompatExe (Join-Path $root 'tests\corpus\shared_formulas.xls') (Join-Path $root 'tests\corpus\StringContinueRecords.xls')
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas .xls (cÃ³digo $LASTEXITCODE)." }

$warningExe = Join-Path $root 'tests\bin\WarningProbe.exe'
$warningArgs = @('/target:exe', ('/out:' + $warningExe), '/main:WarningProbe') + $common +
    $src + @((Join-Path $root 'tests\WarningProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de avisos de compatibilidad...'
& $csc $warningArgs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de avisos no compilaron (cÃ³digo $LASTEXITCODE)." }
& $warningExe (Join-Path $root 'tests\bin\advanced.xlsx') (Join-Path $root 'tests\corpus\StringContinueRecords.xls')
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de avisos (cÃ³digo $LASTEXITCODE)." }

$printExe = Join-Path $root 'tests\bin\PrintProbe.exe'
$printArgs = @('/target:exe', ('/out:' + $printExe), '/main:PrintProbe') + $common +
    $src + @((Join-Path $root 'tests\PrintProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de impresiÃ³n...'
& $csc $printArgs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de impresiÃ³n no compilaron (cÃ³digo $LASTEXITCODE)." }
& $printExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de impresiÃ³n (cÃ³digo $LASTEXITCODE)." }

$chartExe = Join-Path $root 'tests\bin\ChartProbe.exe'
$chartArgs = @('/target:exe', ('/out:' + $chartExe), '/main:ChartProbe') + $common +
    $src + @((Join-Path $root 'tests\ChartProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de grÃ¡ficos...'
& $csc $chartArgs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de grÃ¡ficos no compilaron (cÃ³digo $LASTEXITCODE)." }
& $chartExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de grÃ¡ficos (cÃ³digo $LASTEXITCODE)." }

$replaceExe = Join-Path $root 'tests\bin\ReplaceProbe.exe'
$replaceArgs = @('/target:exe', ('/out:' + $replaceExe), '/main:ReplaceProbe') + $common +
    $src + @((Join-Path $root 'tests\ReplaceProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de buscar y reemplazar...'
& $csc $replaceArgs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de reemplazo no compilaron (código $LASTEXITCODE)." }
& $replaceExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de reemplazo (código $LASTEXITCODE)." }

$featureExe = Join-Path $root 'tests\bin\FeatureProbe.exe'
$featureArgs = @('/target:exe', ('/out:' + $featureExe), '/main:P') + $common +
    $src + @((Join-Path $root 'tests\FeatureProbe.cs'))
Write-Host 'Compilando y ejecutando pruebas de validación y strict...'
& $csc $featureArgs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de validación/strict no compilaron (código $LASTEXITCODE)." }
& $featureExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de validación/strict (código $LASTEXITCODE)." }

$gridTestFiles = @('Theme.cs', 'SheetGrid.cs', 'Model.cs', 'FormulaEngine.cs', 'Functions.cs', 'ConditionalFormatting.cs', 'TableFormatting.cs', 'Charts.cs') |
    ForEach-Object { Join-Path $root ('src\' + $_) }
$gridTestExe = Join-Path $root 'tests\bin\GridLogicTest.exe'
$gargs = @('/target:exe', ('/out:' + $gridTestExe), '/r:System.Windows.Forms.dll') + $common + $gridTestFiles + @((Join-Path $root 'tests\GridLogicTest.cs'))
Write-Host 'Compilando y ejecutando pruebas de la cuadrÃ­cula...'
& $csc $gargs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de cuadrÃ­cula no compilaron (cÃ³digo $LASTEXITCODE)." }
& $gridTestExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de cuadrÃ­cula (cÃ³digo $LASTEXITCODE)." }

$auditFiles = @('Theme.cs', 'SheetGrid.cs', 'Model.cs', 'FormulaEngine.cs', 'Functions.cs', 'ConditionalFormatting.cs', 'TableFormatting.cs',
    'FileIO.cs', 'Ole2.cs', 'XlsIO.cs', 'XlsFormula.cs', 'XlsxIO.cs', 'Compatibility.cs', 'Charts.cs') |
    ForEach-Object { Join-Path $root ('src\' + $_) }
$auditExe = Join-Path $root 'tests\bin\AuditRegressionTest.exe'
$aargs = @('/target:exe', ('/out:' + $auditExe), '/main:AuditRegressionTest') + $common +
    $auditFiles + @((Join-Path $root 'tests\AuditRegressionTest.cs'))
Write-Host 'Compilando y ejecutando pruebas de regresiÃ³n...'
& $csc $aargs
if ($LASTEXITCODE -ne 0) { throw "Las pruebas de regresiÃ³n no compilaron (cÃ³digo $LASTEXITCODE)." }
& $auditExe
if ($LASTEXITCODE -ne 0) { throw "Fallaron las pruebas de regresiÃ³n (cÃ³digo $LASTEXITCODE)." }

$undoExe = Join-Path $root 'tests\bin\UndoViewProbe.exe'
$uargs = @('/target:exe', ('/out:' + $undoExe), '/main:UndoViewProbe') + $common +
    $src + @((Join-Path $root 'tests\UndoViewProbe.cs'))
Write-Host 'Compilando y ejecutando prueba de deshacer...'
& $csc $uargs
if ($LASTEXITCODE -ne 0) { throw "La prueba de deshacer no compilÃ³ (cÃ³digo $LASTEXITCODE)." }
& $undoExe
if ($LASTEXITCODE -ne 0) { throw "FallÃ³ la prueba de deshacer (cÃ³digo $LASTEXITCODE)." }

$exeInfo = Get-Item $exe
Write-Host ("Listo: {0} ({1:N0} bytes)" -f $exeInfo.FullName, $exeInfo.Length)

