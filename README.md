# ADIFConvert

一个将业余无线电台日志从多维表格转换为标准 ADIF 格式的工具，方便与各种业余无线电日志软件兼容使用。

## 功能说明

该工具能够读取从[此模板](https://www.kdocs.cn/l/cqdXmjZlChyr)导出的 Excel 日志文件，并将其转换为符合 ADIF
标准的日志文件，适用于业余无线电爱好者的通联记录管理。

## 多维表格模板

使用前请先复制以下模板到您的金山文档，创建您的日志表格：https://www.kdocs.cn/l/cqdXmjZlChyr

## 使用方法

### 方法一：通过 Python 脚本运行

1. 确保已安装 Python 3.6+ 环境

2. 安装依赖包：

   ```bash
   pip install openpyxl
   ```

3. 查看脚本帮助：

   ```bash
   python main.py -h
   ```

4. 运行脚本

   ```bash
   python main.py
   ```

### 方法二：通过可执行文件运行

1. 双击运行 `ADIFConvert.exe` 即可

2. 使用命令行查看脚本帮助：

   ```bash
   ./ADIFConvert_V0.1.exe -h
   ```

3. 使用命令行运行脚本

   ```bash
   ./ADIFConvert_V0.1.exe
   ```

### 自定义参数

```plaintext
python main.py -i [输入Excel文件路径] -o [输出ADIF文件路径]
./ADIFConvert_V0.1.exe -i [输入Excel文件路径] -o [输出ADIF文件路径]
```

如果不指定参数，默认会：

- 读取 `用户下载文件夹/(日志)表格视图.xlsx`
- 输出到 `用户下载文件夹/tqsl.adi`

## 支持的字段映射

工具会将表格中的以下字段转换为 ADIF 标准字段：

| 表格字段   | ADIF 字段    |
|--------|------------|
| 呼号     | CALL       |
| 模式     | MODE       |
| 对方信号   | RST_RCVD   |
| 己方信号   | RST_SENT   |
| 对方设备   | RIG        |
| 对方天馈   | ANTENNA    |
| 对方功率   | RX_PWR     |
| 对方 QTH | QTH        |
| 己方设备   | MY_RIG     |
| 己方天馈   | MY_ANTENNA |
| 己方功率   | TX_PWR     |
| 己方 QTH | MY_CITY    |
| 补充     | COMMENTS   |

## 注意事项

1. 请确保日志表格包含必要的列：呼号、时间、频率
2. 时间会自动从东八区 (UTC+8) 转换为 ADIF 要求的 UTC 时间
3. 程序会跳过处理异常的记录，并在日志中显示错误信息
4. 生成的 ADIF 文件可直接用于 TQSL 等日志管理软件

## 关于作者

本工具由 BH5UQJ 开发，源码托管于：https://github.com/lingkai5wu/ADIFConvert