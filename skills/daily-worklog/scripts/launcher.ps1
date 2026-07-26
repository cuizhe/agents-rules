$ChromePath = "D:\Program Files\RunningCheeseChrome\App\chrome.exe"
$ChromeArgs = "--remote-debugging-port=9222"

if (-not (Test-Path $ChromePath)) {
    Write-Host "[错误] Chrome 未找到: $ChromePath" -ForegroundColor Red
    exit 1
}

# 检查是否已有 Chrome 进程在运行
$running = Get-Process | Where-Object { $_.ProcessName -like "*chrome*" }
if ($running) {
    # 进一步检查是否已开启调试端口
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:9222/json/version" -TimeoutSec 2 -UseBasicParsing
        if ($response.StatusCode -eq 200) {
            Write-Host "[信息] Chrome 已运行且调试端口已开启（9222），无需重复启动。" -ForegroundColor Green
            exit 0
        }
    } catch {
        Write-Host "[警告] Chrome 正在运行但未开启调试端口，将重新启动..." -ForegroundColor Yellow
        Stop-Process -Name "chrome" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
}

Write-Host "[信息] 正在启动 Chrome（调试端口 9222）..." -ForegroundColor Green
Start-Process -FilePath $ChromePath -ArgumentList $ChromeArgs
Write-Host "[成功] Chrome 已启动，请手动登录系统后再执行填单脚本。" -ForegroundColor Green
