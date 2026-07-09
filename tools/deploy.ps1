$ErrorActionPreference = "Stop"

Write-Host "[1/5] Cleaning old Hexo files..."
hexo clean
if ($LASTEXITCODE -ne 0) {
    throw "hexo clean failed"
}

Write-Host "[2/5] Generating static website..."
hexo generate
if ($LASTEXITCODE -ne 0) {
    throw "hexo generate failed"
}

Write-Host "[3/5] Removing old upload directory..."
ssh ubuntu-lab "rm -rf ~/public"
if ($LASTEXITCODE -ne 0) {
    throw "remote cleanup failed"
}

Write-Host "[4/5] Uploading website to Ubuntu..."
scp -r .\public ubuntu-lab:/home/renyaocheng/
if ($LASTEXITCODE -ne 0) {
    throw "scp upload failed"
}

Write-Host "[5/5] Publishing website with rsync..."
ssh ubuntu-lab "rsync -av --delete ~/public/ /srv/software-tools-site/ && curl -I http://localhost"
if ($LASTEXITCODE -ne 0) {
    throw "remote deployment failed"
}

Write-Host "Deployment completed: http://192.168.56.101"
