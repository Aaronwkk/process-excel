# Excel批量处理工具

这是一个用于批量处理Excel文件的Python工具。它可以对指定文件夹中的所有Excel文件应用相同的计算公式，并将结果写入指定的列。

## 功能特点

- 支持批量处理多个Excel文件
- 支持自定义计算公式
- 支持多工作表处理
- 自动错误处理和报告

## 安装要求

- Python 3.8 或更高版本
- Poetry 包管理器

## 安装步骤

1. 克隆项目到本地
2. 在项目根目录运行以下命令安装依赖：
   ```bash
   brew install python@3.10
   poetry env use python3.10
   poetry lock --no-cache --regenerate
   poetry install
   ```

## 使用方法

1. 运行程序：
   ```bash
   poetry run convert /Users/a1/理赔文件
   poetry run my-cli
   ```

2. 按照提示输入：
   - Excel文件所在文件夹路径
   - 计算公式（使用列字母，如 A*B）
   - 输出列（字母）

## 示例

如果要在Excel中计算：赔偿金额 = 投保面积 × 赔偿系数
- 投保面积在A列
- 赔偿系数在B列
- 结果要输出到C列

则计算公式输入：A*B

## 注意事项

- 确保Excel文件未被其他程序打开
- 建议在处理前备份原始文件
- 公式中只能使用基本的数学运算符（+、-、*、/）