#!/bin/bash

# 设置颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}================================${NC}"
echo -e "${BLUE}    订单分析系统启动中...${NC}"
echo -e "${BLUE}================================${NC}"
echo

# 检查虚拟环境是否存在
if [ ! -f ".venv/bin/activate" ]; then
    echo -e "${RED}[错误] 虚拟环境不存在！${NC}"
    echo "请先运行以下命令创建虚拟环境："
    echo "  python3 -m venv .venv"
    echo "  source .venv/bin/activate"
    echo "  pip install -r requirements.txt"
    echo
    exit 1
fi

# 激活虚拟环境
echo -e "${YELLOW}[步骤1] 激活虚拟环境...${NC}"
source .venv/bin/activate
if [ $? -ne 0 ]; then
    echo -e "${RED}[错误] 无法激活虚拟环境！${NC}"
    exit 1
fi
echo -e "${GREEN}[完成] 虚拟环境已激活${NC}"

# 检查依赖是否安装
echo
echo -e "${YELLOW}[步骤2] 检查依赖...${NC}"
python -c "import flask, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo -e "${YELLOW}[警告] 检测到缺少依赖，正在安装...${NC}"
    pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo -e "${RED}[错误] 依赖安装失败！${NC}"
        exit 1
    fi
fi
echo -e "${GREEN}[完成] 所有依赖已就绪${NC}"

# 启动Flask应用
echo
echo -e "${YELLOW}[步骤3] 启动Flask服务...${NC}"
echo -e "${GREEN}服务地址: http://localhost:5000${NC}"
echo -e "${YELLOW}按 Ctrl+C 停止服务${NC}"
echo

echo -e "${BLUE}================================${NC}"
echo -e "${BLUE}    服务启动完成！${NC}"
echo -e "${BLUE}================================${NC}"
echo

python app.py 