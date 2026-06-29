#!/bin/bash
# AI-Quant GitHub SSH 自动推送脚本
# 用法：在终端运行 bash push.sh

cd /Users/wangfeynman/Desktop/AI-Quant

echo "📦 开始推送到 GitHub (SSH)..."
echo "仓库地址：git@github.com:FeynmanWang996/AI-Quant.git"
echo ""

git push origin main

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ 推送成功！"
    echo "查看仓库：https://github.com/FeynmanWang996/AI-Quant"
else
    echo ""
    echo "❌ 推送失败，请检查 SSH 配置。"
    echo "运行 ssh -T git@github.com 测试连接。"
fi
