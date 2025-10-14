param(
  [Parameter(Mandatory=$true)][string]$Message,
  [string]$Tag,
  [string]$Branch = "dev"
)
$ErrorActionPreference = "Stop"

# Cambia a la rama objetivo
git switch $Branch | Out-Null

# Trae últimos cambios
git pull --rebase

# Prepara cambios
git add -A

# ¿Hay algo que confirmar?
git diff --cached --quiet
$needCommit = $LASTEXITCODE -ne 0
if ($needCommit) {
  git commit -m $Message
} else {
  Write-Host "No hay cambios para confirmar."
}

# Sube
git push

# Tag opcional
if ($Tag) {
  git tag -a $Tag -m $Message
  git push origin $Tag
  Write-Host "Tag $Tag publicado."
}
