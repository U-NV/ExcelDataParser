# Excel数据处理工具包

Unity编辑器插件，提供完整的Excel文件读写功能。支持Excel文件的读取、解析、创建和写入，具备强大的数据验证、类型转换和异常处理能力。

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

// 创建Excel文件
var filePath = "Assets/Data/Output.xlsx";
var data = new List<Dictionary<string, object>>
{
    new Dictionary<string, object> { {"Name", "张三"}, {"Age", 25}, {"Score", 95.5} },
    new Dictionary<string, object> { {"Name", "李四"}, {"Age", 30}, {"Score", 88.0} }
};

// 写入数据
ExcelWriter.WriteDataToFile(filePath, data);
```

### 配置设置

```csharp
// 配置ExcelReader
ExcelReaderConfig.CustomKeywords.Add("custom_key");
ExcelReaderConfig.CaseSensitive = true;
ExcelReaderConfig.MaxRows = 5000;
ExcelReaderConfig.StrictValidation = false;
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

## 系统要求

- Unity 2019.4.25f1 或更高版本
- .NET Framework 4.7.1 或更高版本
- EPPlus库支持

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request来改进这个工具包。

## 更新日志

查看 [CHANGELOG.md](CHANGELOG.md) 了解版本更新历史。

## 技术支持

如有问题，请通过以下方式联系：
- 邮箱: support@u0ugames.com
- GitHub Issues: [提交问题](https://github.com/U-NV/FeiShu-Unity-Integration/issues)
