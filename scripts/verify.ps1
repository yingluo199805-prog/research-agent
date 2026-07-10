$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\\Scripts\\python.exe"
if (-not (Test-Path -LiteralPath $python)) { $python = "python" }

Set-Location -LiteralPath $repoRoot
& $python -X utf8 -c "import ast, pathlib; files=list(pathlib.Path('src').rglob('*.py'))+list(pathlib.Path('archive').rglob('*.py')); [ast.parse(p.read_text(encoding='utf-8-sig')) for p in files]; print(f'Python syntax: OK ({len(files)} files)')"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

git -C $repoRoot status --short
