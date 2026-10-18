#!/bin/bash
# QSlide dev server 启动包装（供 LaunchAgent 调用，确保 PATH 正确）
export PATH="/Users/enn/.workbuddy/binaries/node/versions/22.22.2/bin:/usr/local/bin:/usr/bin:/bin:$PATH"
cd /Users/enn/Github/QSlide
exec npm run dev
