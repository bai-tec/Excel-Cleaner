
# 📊 Excel 数据清洗工具

一个简单易用的 Excel 数据清洗和预处理工具，带有美观的 Web 界面。

## ✨ 功能特点

- 📤 上传 Excel 文件（支持 .xlsx 和 .xls 格式）
- 🔍 数据预览和基本统计
- 🧹 删除重复行、空行、空列
- 📝 缺失值填充（前向、后向、平均值、中位数、众数）
- 🔄 数据类型转换
- ✅ 手机号和邮箱格式验证
- 📈 数据可视化展示
- 💾 导出清洗后的结果（支持 Excel 和 CSV）
- 📋 操作历史记录

## 🚀 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行程序

```bash
streamlit run main.py
```

然后浏览器会自动打开 http://localhost:8501 ，就可以开始使用了！

## 📖 使用说明

1. **上传文件** - 点击上传按钮选择你的 Excel 文件
2. **查看数据** - 在数据预览标签页查看原始数据和统计信息
3. **清洗数据** - 在数据清洗标签页进行各种清洗操作
4. **验证数据** - 对手机号、邮箱等格式进行验证
5. **导出结果** - 导出清洗后的 Excel 或 CSV 文件

## 🛠️ 技术栈

- Python 3.8+
- Streamlit - Web 界面框架
- Pandas - 数据处理
- NumPy - 数值计算
- OpenPyXL - Excel 文件处理

## 📁 项目结构

```
excel-cleaner/
├── main.py           # 主程序文件
├── requirements.txt  # Python 依赖
└── README.md         # 项目说明文档
```

## 💡 使用场景

- HR 数据清洗
- 销售数据整理
- 学生成绩处理
- 客户数据清洗
- 调研数据分析

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

MIT License

## 🌟 支持

如果这个项目对你有帮助，欢迎给个 Star！

