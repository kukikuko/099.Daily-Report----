Write-Host "Daily Report Automation 빌드를 시작합니다..." -ForegroundColor Green

Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

pyinstaller --clean DailyReport.spec

if (Test-Path "dist/DailyReport.exe") {
    Write-Host "✅ 빌드 성공: dist/DailyReport.exe" -ForegroundColor Green
} else {
    Write-Host "❌ 빌드 실패" -ForegroundColor Red
}
