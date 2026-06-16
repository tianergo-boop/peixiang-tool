# 配箱工具 v1.0

## 简介
替代Excel配箱表的桌面工具，将VLOOKUP/COUNTIF公式逻辑转换为Python字典索引，实现毫秒级配箱计算。

## 功能
- 📋 订单导入（剪贴板粘贴 / CSV文件）
- 🚀 一键配箱（自动匹配箱号）
- 📊 结果导出（配箱表 / 发货通知单）
- 📥 ASN管理（插入发货通知数据）
- ⚙️ Excel双向同步 + 自动备份

## 运行要求
- Windows 7+ / macOS 10.12+
- 内存: ≥ 80MB
- 无需安装Python（exe版本）

## 开发运行
```bash
pip install openpyxl
python -m peixiang_tool
```

## 打包为EXE
```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "配箱工具" run.py
```
