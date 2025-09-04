# Excel数据处理工具包

Unity编辑器插件，提供完整的Excel文件读写功能。支持Excel文件的读取、解析、创建和写入，具备强大的数据验证、类型转换和异常处理能力。

## 安装插件

### 方法一：Package Manager安装（推荐）
1. 在Unity编辑器中打开 `Window > Package Manager`
2. 点击左上角的 `+` 按钮，选择 `Add package from git URL`
3. 输入：`https://github.com/U-NV/ExcelDataParser.git`
4. 点击 `Add` 完成安装

### 方法二：手动安装
1. 下载最新版本的插件包
2. 解压到Unity项目的 `Assets/` 目录下
3. 重新打开Unity编辑器


## 功能特性

### 📖 ExcelReader - 数据读取器
- **智能解析**: 自动识别Excel文件结构，支持.xlsx和.xls格式
- **类型转换**: 支持string、int、float、double、bool、list、array、object等数据类型
- **关键词识别**: 内置var、type、default等关键词，支持自定义关键词扩展
- **数据验证**: 严格的类型验证和格式检查，确保数据完整性
- **异常处理**: 完善的错误处理机制，提供详细的错误信息

### ✏️ ExcelWriter - 数据写入器
- **灵活写入**: 支持按行列位置精确写入数据
- **自动创建**: 自动创建目录结构和Excel文件
- **格式支持**: 支持多种数据类型的写入
- **批量操作**: 支持批量数据写入和格式化

### ⚙️ 配置管理
- **自定义关键词**: 支持添加自定义关键词列表
- **大小写控制**: 可配置是否区分大小写
- **限制设置**: 可设置最大行数、列数和嵌套层级
- **验证模式**: 可选择严格验证或宽松模式
- **文件格式**: 支持多种Excel文件扩展名

## 使用场景

- **游戏配置管理**: 将游戏配置数据存储在Excel中，运行时动态加载
- **数据导入导出**: 在Unity编辑器中导入导出表格数据
- **批量数据处理**: 处理大量结构化数据
- **配置表生成**: 从Excel生成游戏配置表
- **数据分析**: 在Unity中分析Excel数据

## 系统要求

- **Unity版本**: Unity 2019.4.25f1 或更高版本
- **.NET Framework**: 4.7.1 或更高版本
- **依赖库**: EPPlus库（已包含在插件中）
- **支持格式**: .xlsx, .xls

## 快速开始

### 读取Excel文件

```csharp
using U0UGames.Excel;

// 读取Excel文件
var data = ExcelReader.GetRawData("Assets/Data/Config.xlsx");

// 遍历数据
foreach (var row in data)
{
    foreach (var kvp in row)
    {
        Debug.Log($"{kvp.Key}: {kvp.Value}");
    }
}
```

### 写入Excel文件

```csharp
using U0UGames.Excel;

// 创建Excel数据
var excelData = new ExcelWriter.ExcelData();

// 设置数据（按行列位置）
excelData[1, 1] = "Name";      // A1
excelData[1, 2] = "Age";       // B1
excelData[1, 3] = "Score";     // C1
excelData[2, 1] = "张三";      // A2
excelData[2, 2] = "25";        // B2
excelData[2, 3] = "95.5";      // C2
excelData[3, 1] = "李四";      // A3
excelData[3, 2] = "30";        // B3
excelData[3, 3] = "88.0";      // C3

// 保存文件
var filePath = "Assets/Data/Output.xlsx";
ExcelWriter.SaveFile(excelData, filePath);

// 或者使用自定义工作表名称
ExcelWriter.SaveFile(excelData, filePath, "MySheet");
```

### 配置设置

```csharp
// 配置ExcelReader
ExcelReaderConfig.CustomKeywords.Add("custom_key");
ExcelReaderConfig.CaseSensitive = true;
ExcelReaderConfig.MaxRows = 5000;
ExcelReaderConfig.StrictValidation = false;
```

### ExcelWriter详细用法

```csharp
using U0UGames.Excel;

// 创建Excel数据对象
var excelData = new ExcelWriter.ExcelData();

// 方法1: 使用索引器设置数据
excelData[1, 1] = "Header1";  // A1单元格
excelData[1, 2] = "Header2";  // B1单元格
excelData[2, 1] = "Data1";    // A2单元格
excelData[2, 2] = "Data2";    // B2单元格

// 方法2: 批量设置数据
for (int row = 1; row <= 10; row++)
{
    for (int col = 1; col <= 5; col++)
    {
        excelData[row, col] = $"R{row}C{col}";
    }
}

// 保存文件（使用默认工作表名称）
ExcelWriter.SaveFile(excelData, "Assets/Data/Output.xlsx");

// 保存文件（使用自定义工作表名称）
ExcelWriter.SaveFile(excelData, "Assets/Data/Output.xlsx", "MyCustomSheet");
```

## 支持的数据类型

| 类型 | 描述 | 示例 |
|------|------|------|
| string | 字符串 | "Hello World" |
| int | 整数 | 123 |
| float | 单精度浮点数 | 3.14f |
| double | 双精度浮点数 | 3.14159 |
| bool | 布尔值 | true/false |
| list | 列表 | [1, 2, 3] |
| array | 数组 | {1, 2, 3} |
| object | 对象 | {key: value} |

## 关键词说明

### 内置关键词
- **var**: 变量名定义
- **type**: 数据类型定义
- **default**: 默认值设置

### 自定义关键词
可以通过`ExcelReaderConfig.CustomKeywords`添加自定义关键词，扩展功能。

## 异常处理

插件提供完善的异常处理机制：

- **ExcelException**: Excel读取异常基类
- **ExcelWriteException**: Excel写入异常
- **详细错误信息**: 包含文件路径和具体错误描述

```csharp
try
{
    var data = ExcelReader.GetRawData("invalid_file.xlsx");
}
catch (ExcelException ex)
{
    Debug.LogError($"读取Excel文件失败: {ex.Message}, 文件: {ex.FilePath}");
}
```

## 性能优化

- **内存管理**: 优化的内存使用，支持大文件处理
- **批量处理**: 支持批量数据操作，提高处理效率
- **缓存机制**: 内置缓存机制，避免重复解析

## 故障排除

### 常见问题

**Q: 安装后无法找到ExcelDataParser？**
A: 确保已正确安装插件，检查Package Manager中是否显示插件已安装。

**Q: 读取Excel文件时出现权限错误？**
A: 确保Excel文件没有被其他程序占用，检查文件路径是否正确。

**Q: 数据验证失败？**
A: 检查Excel文件格式是否符合要求，确保关键词行以#开头。

**Q: 内存不足错误？**
A: 对于大文件，可以调整`ExcelReaderConfig.MaxRows`和`ExcelReaderConfig.MaxColumns`限制。

### 调试技巧

```csharp
// 启用详细日志
ExcelReaderConfig.StrictValidation = true;

// 检查文件是否支持
if (ExcelReader.IsSupportedFile("your_file.xlsx"))
{
    var data = ExcelReader.GetRawData("your_file.xlsx");
}

// 清理缓存
ExcelReader.ClearCache();
```

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request来改进这个工具包。

## 更新日志

查看 [CHANGELOG.md](CHANGELOG.md) 了解版本更新历史。

## 技术支持

如有问题，请通过以下方式联系：
- 邮箱: haowei1117@foxmail.com
- GitHub Issues: [提交问题](https://github.com/U-NV/ExcelDataParser/issues)
- GitHub仓库: [ExcelDataParser](https://github.com/U-NV/ExcelDataParser)
