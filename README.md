# ExcelToSQLite

一个将 Excel 数据导入到 SQLite 数据库的工具。

## 功能特性

一是，支持读取 `.xlsx` 和 `.xls` 格式的 Excel 文件；
二是，自动创建 SQLite 数据库表结构；
三是，可建立SQL分析指标，进行数据分析；
四是，通过Deepseek API接口，使用自然描述分析需求，建立分析指标，实现“0技术基础”分析建模；
五是，支持数据预览、数据导出；
五是，系统支持跨平台编译，可在 Windows、Linux等系统中运行使用。

## 技术栈

- .NET 8.0
- SQLite
- Excel 数据读取库

## 项目结构

```bash
ExcelToSQLite/
├── Program.cs          # 主程序入口
├── ExcelToSQLite.csproj # 项目配置文件
├── .gitignore          # Git 忽略文件配置
└── README.md           # 项目说明文档
```

## 许可证

本项目仅供个人学习和使用。

## 作者

GitHub: linydvanclean

## 快速开始

### 环境要求

- .NET 8.0 SDK
- Git（可选）

### 克隆项目

```bash
git clone https://github.com/linydvanclean/ExcelToSQLite.git
cd ExcelToSQLite
```

### 还原依赖并运行

```bash
dotnet restore
dotnet build
dotnet run
```

### 使用说明

1、读取Excel 文件；
2、运行程序，按提示操作；
3、数据将自动导入到 SQLite 数据库；


