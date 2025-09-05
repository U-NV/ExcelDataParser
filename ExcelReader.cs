using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using UnityEngine;

namespace U0UGames.ExcelDataParser
{
    /// <summary>
    /// Excel读取器配置类
    /// </summary>
    public static class ExcelReaderConfig
    {
        /// <summary>自定义关键词列表</summary>
        public static List<string> CustomKeywords { get; set; } = new List<string>();
        
        /// <summary>是否区分大小写</summary>
        public static bool CaseSensitive { get; set; } = false;
        
        /// <summary>最大行数限制</summary>
        public static int MaxRows { get; set; } = 10000;
        
        /// <summary>最大列数限制</summary>
        public static int MaxColumns { get; set; } = 1000;
        
        /// <summary>最大嵌套层级</summary>
        public static int MaxNestingLevel { get; set; } = 10;
        
        /// <summary>是否启用严格验证</summary>
        public static bool StrictValidation { get; set; } = true;
        
        /// <summary>支持的文件扩展名</summary>
        public static string[] SupportedExtensions { get; set; } = { ".xlsx", ".xls" };
        
        /// <summary>有效的变量名正则表达式</summary>
        public static string VariableNamePattern { get; set; } = @"^[a-zA-Z_][a-zA-Z0-9_]*$";
        
        /// <summary>有效的类型列表</summary>
        public static string[] ValidTypes { get; set; } = { "string", "int", "float", "double", "bool", "list", "array", "object" };
    }

    /// <summary>
    /// Excel读取异常基类
    /// </summary>
    public class ExcelException : Exception
    {
        public string FilePath { get; }
        public string SheetName { get; }
        
        public ExcelException(string message, string filePath = null, string sheetName = null) 
            : base(message)
        {
            FilePath = filePath;
            SheetName = sheetName;
        }
        
        public ExcelException(string message, Exception innerException, string filePath = null, string sheetName = null) 
            : base(message, innerException)
        {
            FilePath = filePath;
            SheetName = sheetName;
        }
    }
    
    /// <summary>
    /// Excel文件读取异常
    /// </summary>
    public class ExcelFileException : ExcelException
    {
        public ExcelFileException(string message, string filePath, Exception innerException = null) 
            : base(message, innerException, filePath)
        {
        }
    }
    
    /// <summary>
    /// Excel工作表读取异常
    /// </summary>
    public class ExcelSheetException : ExcelException
    {
        public ExcelSheetException(string message, string filePath, string sheetName, Exception innerException = null) 
            : base(message, innerException, filePath, sheetName)
        {
        }
    }
    
    /// <summary>
    /// Excel数据解析异常
    /// </summary>
    public class ExcelDataException : ExcelException
    {
        public int Row { get; }
        public int Column { get; }
        
        public ExcelDataException(string message, string filePath, string sheetName, int row = -1, int column = -1, Exception innerException = null) 
            : base(message, innerException, filePath, sheetName)
        {
            Row = row;
            Column = column;
        }
    }
    public static class ExcelReader
    {
        /// <summary>
        /// Excel解析关键词定义类
        /// 用于识别Excel表格中的特殊标记行
        /// </summary>
        private static class Keyword
        {
            /// <summary>变量名关键词，用于定义字段名称</summary>
            public static readonly string Var = "var";
            /// <summary>类型关键词，用于定义字段数据类型</summary>
            public static readonly string Type = "type";
            /// <summary>默认值关键词，用于定义字段默认值</summary>
            public static readonly string Default = "default";
            
            /// <summary>所有支持的关键词列表</summary>
            private static readonly List<string> BaseKeywordList = new List<string>()
            {
                Var,       
                Type,
                Default,
            };
            
            /// <summary>
            /// 获取所有关键词列表（包括自定义关键词）
            /// </summary>
            public static List<string> GetAllKeywords()
            {
                var allKeywords = new List<string>(BaseKeywordList);
                allKeywords.AddRange(ExcelReaderConfig.CustomKeywords);
                return allKeywords;
            }
            
            /// <summary>
            /// 检查指定值是否为有效关键词
            /// </summary>
            /// <param name="value">要检查的字符串</param>
            /// <returns>如果是关键词返回true，否则返回false</returns>
            public static bool Contains(string value)
            {
                if (string.IsNullOrEmpty(value)) return false;
                
                var keywords = GetAllKeywords();
                if (ExcelReaderConfig.CaseSensitive)
                {
                    return keywords.Contains(value);
                }
                else
                {
                    return keywords.Any(k => string.Equals(k, value, StringComparison.OrdinalIgnoreCase));
                }
            }
        }
        private class ColumnData
        {
            private Dictionary<string, List<string>> keyToKeyNameList = new Dictionary<string, List<string>>();
            // private Dictionary<string, List<string>> keyToKeyTypeList = new Dictionary<string, List<string>>();
            private string _defaultValue = null;
            public string DefaultValue => _defaultValue;
            public int ColumnIndex { get; }
            public ColumnData(int index)
            {
                ColumnIndex = index;
            }
            public List<string> Get(string keyword)
            {
                if (keyToKeyNameList.TryGetValue(keyword, out var dataList))
                {
                    return dataList;
                }
                
                return null;
            }
            public void Add(string keyword, string value)
            {
                if (!keyToKeyNameList.TryGetValue(keyword, out var valueList))
                {
                    valueList = new List<string>();
                    valueList.Add(value);
                    keyToKeyNameList[keyword] = valueList;
                }
                else
                {
                    valueList.Add(value);
                }
                
                if (keyword == Keyword.Default && !string.IsNullOrEmpty(value))
                {
                    _defaultValue = value;
                }
            }
        }
        /// <summary>列索引到列数据的映射字典</summary>
        [ThreadStatic]
        private static Dictionary<int, ColumnData> ColumDataLookup;
        
        /// <summary>字段路径到类型的映射字典</summary>
        [ThreadStatic]
        private static Dictionary<string, string> PathToType;
        
        /// <summary>字段路径到默认值的映射字典</summary>
        [ThreadStatic]
        private static Dictionary<string, string> PathToDefault;
        
        /// <summary>额外的数据字典，存储非标准关键词的数据</summary>
        [ThreadStatic]
        private static Dictionary<string, string> AdditionalData;
        
        /// <summary>当前文件路径（线程安全）</summary>
        [ThreadStatic]
        private static string _currFilePath;
        
        /// <summary>当前工作表名称（线程安全）</summary>
        [ThreadStatic]
        private static string _currSheetName;
        
        /// <summary>
        /// 初始化线程静态字段
        /// </summary>
        private static void InitializeThreadStaticFields()
        {
            if (ColumDataLookup == null) ColumDataLookup = new Dictionary<int, ColumnData>();
            if (PathToType == null) PathToType = new Dictionary<string, string>();
            if (PathToDefault == null) PathToDefault = new Dictionary<string, string>();
            if (AdditionalData == null) AdditionalData = new Dictionary<string, string>();
        }
        
        /// <summary>
        /// 清理静态缓存数据，防止内存泄漏
        /// </summary>
        public static void ClearCache()
        {
            ColumDataLookup?.Clear();
            PathToType?.Clear();
            PathToDefault?.Clear();
            AdditionalData?.Clear();
            _currFilePath = null;
            _currSheetName = null;
        }
        
        /// <summary>
        /// 验证列数据的有效性
        /// </summary>
        /// <param name="columnData">列数据</param>
        /// <returns>如果数据有效返回true，否则返回false</returns>
        private static bool ValidateColumnData(ColumnData columnData)
        {
            if (columnData == null) return false;
            
            var nameList = columnData.Get(Keyword.Var);
            var typeList = columnData.Get(Keyword.Type);
            
            if (nameList == null || typeList == null) return false;
            if (nameList.Count != typeList.Count) return false;
            
            if (ExcelReaderConfig.StrictValidation)
            {
                // 验证变量名格式
                foreach (var name in nameList)
                {
                    if (!IsValidVariableName(name))
                    {
                        Debug.LogError($"第{columnData.ColumnIndex}列包含无效的变量名: {name}");
                        return false;
                    }
                }
                
                // 验证类型格式
                foreach (var type in typeList)
                {
                    if (!IsValidType(type))
                    {
                        Debug.LogError($"第{columnData.ColumnIndex}列包含无效的类型: {type}");
                        return false;
                    }
                }
            }
            
            return true;
        }
        
        /// <summary>
        /// 验证变量名是否有效
        /// </summary>
        /// <param name="name">变量名</param>
        /// <returns>如果有效返回true，否则返回false</returns>
        private static bool IsValidVariableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return false;
            
            try
            {
                return Regex.IsMatch(name, ExcelReaderConfig.VariableNamePattern);
            }
            catch
            {
                // 如果正则表达式无效，使用简单的验证
                return char.IsLetter(name[0]) && name.All(c => char.IsLetterOrDigit(c) || c == '_');
            }
        }
        
        /// <summary>
        /// 验证类型是否有效
        /// </summary>
        /// <param name="type">类型字符串</param>
        /// <returns>如果有效返回true，否则返回false</returns>
        private static bool IsValidType(string type)
        {
            if (string.IsNullOrWhiteSpace(type)) return false;
            
            var normalizedType = ExcelReaderConfig.CaseSensitive ? type : type.ToLower();
            return ExcelReaderConfig.ValidTypes.Any(t => 
                ExcelReaderConfig.CaseSensitive ? t == normalizedType : t.ToLower() == normalizedType);
        }
        
        /// <summary>
        /// 检查文件格式是否支持
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>如果支持返回true，否则返回false</returns>
        public static bool IsSupportedFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return false;
            
            var extension = Path.GetExtension(filePath).ToLower();
            return ExcelReaderConfig.SupportedExtensions.Contains(extension);
        }
        
        /// <summary>
        /// 构建字段路径，优化字符串操作
        /// </summary>
        /// <param name="names">名称列表</param>
        /// <returns>构建的路径字符串</returns>
        private static string BuildPath(List<string> names)
        {
            if (names == null || names.Count == 0) return string.Empty;
            
            var pathBuilder = new StringBuilder();
            for (int i = 0; i < names.Count; i++)
            {
                if (i > 0) pathBuilder.Append('.');
                pathBuilder.Append(names[i]);
            }
            return pathBuilder.ToString();
        }
        /// <summary>
        /// 更新变量查找表，将列数据中的变量名和类型信息添加到全局查找表中
        /// </summary>
        /// <param name="columnData">列数据</param>
        private static void UpdateVarLookup(ColumnData columnData)
        {
            // 验证列数据
            if (!ValidateColumnData(columnData))
            {
                return;
            }
            
            // 获取变量名列表
            var nameList = columnData.Get(Keyword.Var);
            var typeList = columnData.Get(Keyword.Type);
            
            // 构建字段路径并添加到查找表
            string defaultNamePath = BuildPath(nameList);
            
            for (int i = 0; i < nameList.Count; i++)
            {
                string name = nameList[i];
                string type = typeList[i];
                
                // 构建当前层级的路径
                string currentPath = BuildPath(nameList.Take(i + 1).ToList());
                PathToType[currentPath] = type;
            }
            
            // 添加默认值到查找表
            string defaultValue = columnData.DefaultValue;
            if (!string.IsNullOrEmpty(defaultValue))
            {
                PathToDefault[defaultNamePath] = defaultValue;
            }
        }
        /// <summary>
        /// 尝试获取Excel工作表中的列关键词信息
        /// 解析以#开头的关键词行，构建数据结构定义
        /// </summary>
        /// <param name="ws">Excel工作表</param>
        /// <param name="keywordRowNum">关键词行数量</param>
        /// <returns>如果成功解析关键词返回true，否则返回false</returns>
        private static bool TryGetColumKeyword(ExcelWorksheet ws, out int keywordRowNum)
        {
            // 初始化线程静态字段
            InitializeThreadStaticFields();
            
            keywordRowNum = 0;
            Dictionary<int, string> rowToKeyWord = new Dictionary<int, string>();
            AdditionalData.Clear();

            // 验证工作表有效性
            if (ws == null) 
            {
                throw new ExcelSheetException("工作表对象为空", _currFilePath, "未知");
            }
            if (ws.Dimension == null) 
            {
                throw new ExcelSheetException("工作表维度信息为空", _currFilePath, ws.Name);
            }
            if (ws.Dimension.End == null) 
            {
                throw new ExcelSheetException("工作表结束位置信息为空", _currFilePath, ws.Name);
            }
            
            int columnCount = ws.Dimension.End.Column;
            int rowCount = ws.Dimension.End.Row;
            
            // 检查行数和列数限制
            if (rowCount > ExcelReaderConfig.MaxRows)
            {
                Debug.LogWarning($"工作表 '{ws.Name}' 行数超过限制 ({rowCount} > {ExcelReaderConfig.MaxRows})，将只处理前 {ExcelReaderConfig.MaxRows} 行");
                rowCount = ExcelReaderConfig.MaxRows;
            }
            
            if (columnCount > ExcelReaderConfig.MaxColumns)
            {
                Debug.LogWarning($"工作表 '{ws.Name}' 列数超过限制 ({columnCount} > {ExcelReaderConfig.MaxColumns})，将只处理前 {ExcelReaderConfig.MaxColumns} 列");
                columnCount = ExcelReaderConfig.MaxColumns;
            }
            
            // 遍历每一行寻找#开头的数据
            for (int rowNum = 1; rowNum <= rowCount; rowNum++)
            {
                var keyword = ws.Cells[rowNum, 1].Value?.ToString();
                if (keyword != null && keyword.StartsWith("#"))
                {
                    keywordRowNum++;
                    keyword = keyword.TrimStart('#');
                    
                    // 检查是否为标准关键词
                    if (Keyword.Contains(keyword))
                    {
                        rowToKeyWord[rowNum] = keyword;
                    }
                    else
                    {
                        // 处理非标准关键词，将其作为额外数据存储
                        string value = null;
                        for (int i = 2; i < columnCount; i++)
                        {
                            try
                            {
                                value = ws.Cells[rowNum, i].Value?.ToString();
                            }
                            catch (Exception e)
                            {
                                string errorMsg = $"工作表 '{ws.Name}' 第 {rowNum} 行第 {i} 列读取关键词数据时发生错误";
                                Debug.LogError($"{errorMsg}: {e.Message}");
                                throw new ExcelDataException(errorMsg, _currFilePath, ws.Name, rowNum, i, e);
                            }
                            
                            if (!string.IsNullOrEmpty(value))
                            {
                                break;
                            }
                        }

                        AdditionalData[keyword] = value;
                    }
                }
                else
                {
                    // 遇到非#开头的行，停止解析
                    break;
                }
            }
            
            // 检查是否包含必需的关键词 Var 和 Type
            if (rowToKeyWord.Count > 0)
            {
                var keywordsList = rowToKeyWord.Values;
                bool haveKeywordVar = keywordsList.Contains(Keyword.Var);
                bool haveKeywordType = keywordsList.Contains(Keyword.Type);
                if (!haveKeywordVar || !haveKeywordType)
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
            
            // 解析每一列的数据
            PathToType.Clear();
            PathToDefault.Clear();
            ColumDataLookup.Clear();
            
            for (int column = 2; column <= columnCount; column++)
            {
                bool haveData = false;
                ColumnData columnData = new ColumnData(column);
                
                // 遍历关键词行，收集该列的数据
                for (int row = 1; row <= keywordRowNum; row++)
                {
                    // 跳过不需要的关键词
                    if (!rowToKeyWord.TryGetValue(row, out var keyword))
                    {
                        continue;
                    }

                    var cell = ws.Cells[row, column];
                    object cellValue;
                    
                    // 处理合并单元格
                    if (cell.Merge)
                    {
                        string range = ws.MergedCells[row, column];
                        var cellRange = new ExcelAddress(range);
                        cellValue = ws.Cells[cellRange.Start.Row, cellRange.Start.Column].Value;
                    }
                    else
                    {
                        cellValue = cell.Value;
                    }

                    string stringValue = cellValue?.ToString();
                    bool isValidValue = !string.IsNullOrEmpty(stringValue);
                    if (isValidValue)
                    {
                        haveData = true;
                        columnData.Add(keyword, stringValue);
                    }
                }

                // 如果该列有数据，则添加到查找表
                if (haveData)
                {
                    UpdateVarLookup(columnData);
                    ColumDataLookup[column] = columnData;
                }
            }

            return true;
        }
        /// <summary>
        /// 尝试合并两个数据对象，主要用于列表数据的智能合并
        /// 当相邻行的数据结构相似时，将新数据合并到现有列表中
        /// </summary>
        /// <param name="lastValue">上一个数据对象</param>
        /// <param name="newValue">新的数据对象</param>
        /// <param name="parentPath">父级路径</param>
        /// <returns>如果合并成功返回true，否则返回false</returns>
        private static bool TryMerge(Dictionary<string, object> lastValue, Dictionary<string, object> newValue, string parentPath)
        {
            // 安全检查
            if (lastValue == null || newValue == null) return false;
            if (lastValue.Count == 0 || newValue.Count == 0) return false;
            
            // 检查嵌套层级限制
            int nestingLevel = string.IsNullOrEmpty(parentPath) ? 0 : parentPath.Split('.').Length;
            if (nestingLevel > ExcelReaderConfig.MaxNestingLevel)
            {
                Debug.LogWarning($"数据嵌套层级过深 ({nestingLevel} > {ExcelReaderConfig.MaxNestingLevel})，跳过合并");
                return false;
            }
            
            // 获取所有的列表名称
            List<string> listNames = new List<string>();
            foreach (var kvp in newValue)
            {
                if (kvp.Value is List<object>)
                {
                    listNames.Add(kvp.Key);
                }
            }

            // 找到周围都是空的列表名称（可以合并的列表）
            List<string> targetListNames = new List<string>();
            foreach (var varName in listNames)
            {
                bool isTarget = true;
                foreach (var kvp in newValue)
                {
                    // 跳过自己
                    if (kvp.Key == varName)
                    {
                        continue;
                    }

                    // 如果名称为列表则跳过
                    if(listNames.Contains(kvp.Key))
                    {
                        continue;
                    }
                    
                    // 如果有非空值，说明不能被合并，停止循环
                    if (kvp.Value != null)
                    {
                        isTarget = false;
                        break;
                    }
                }
                // 添加周围都是空或列表的列表名称
                if (isTarget)
                {
                    targetListNames.Add(varName);
                }
            }

            // 没有可以合并的列表，返回false
            if (targetListNames.Count == 0)
            {
                return false;
            }
            
            // 遍历每一个可以合并的列表名称
            foreach (var listName in targetListNames)
            {
                var targetObject = lastValue[listName];
                var newValueObject = newValue[listName];
                var currPath = parentPath + listName;

                // 只合并列表对象
                if (targetObject is not List<object> targetList || newValueObject is not List<object> newValueList)
                {
                    continue;
                }
                
                // 新对象内容为空跳过
                if (newValueList.Count == 0) continue;
                var newValueItem = newValueList[0];
                
                // 如果目标列表中目前没有元素，则直接合并
                if (targetList.Count == 0)
                {
                    ApplyDefaultValue(ref newValueItem, currPath);
                    targetList.Add(newValueItem);
                    continue;
                }
                
                var lastValueItem = targetList.Last();
                // 如果两者都有值，尝试递归合并
                if (lastValueItem is Dictionary<string, object> targetClassValue 
                    && newValueItem is Dictionary<string, object> newClassValue)
                {
                    // 合并失败就应用默认值并添加新行
                    if (!TryMerge(targetClassValue, newClassValue, currPath))
                    {
                        ApplyDefaultValueToClass(newClassValue, currPath);
                        targetList.Add(newClassValue);
                    }
                }
                // 如果是值的话，尝试应用默认值
                else
                {
                    ApplyDefaultValue(ref newValueItem, currPath);
                    targetList.Add(newValueItem);
                }
            }
            // 合并完成，返回true
            return true;
        }
        private static Dictionary<string, object> GetRawLineData(ExcelWorksheet ws, int row)
        {
            int columnCount = ws.Dimension.End.Column;
            // 遍历这一行的每一列单元格， 并将这一行转换为数据
            Dictionary<string, object> keyValue = new Dictionary<string, object>();
            for (int column = 2; column <= columnCount; column++)
            {
                // 如果这行没有配置数据，则跳过
                if (!ColumDataLookup.TryGetValue(column, out var columnData))
                {
                    continue;
                }

                string cellValue = null;
                try
                {
                    var cell = ws.Cells[row, column];
                    var rawCellValue = cell.Value;
                    cellValue = rawCellValue?.ToString();
                }
                catch(Exception e)
                {
                    string errorMsg = $"工作表 '{ws.Name}' 第 {row} 行第 {column} 列读取单元格数据时发生错误";
                    Debug.LogError($"{errorMsg}: {e.Message}");
                    throw new ExcelDataException(errorMsg, _currFilePath, ws.Name, row, column, e);
                }
                // 将空字符串设置 null
                if (string.IsNullOrEmpty(cellValue))
                {
                    cellValue = null;
                }
                
                var nameList = columnData.Get(Keyword.Var);
                if (nameList == null || nameList.Count == 0)
                {
                    Debug.LogWarning($"{_currFilePath}:{ws.Name} 的{column}行数据错误，无法找到变量名称");
                    return null;
                }

                // 名称列表代表变量路径
                if (nameList.Count > 1)
                {
                    var currDict = keyValue;
                    for (int i = 0; i < nameList.Count; i++)
                    {
                        bool isLast = i == nameList.Count - 1;
                        string varName = nameList[i];
                        // 最后一个名称为实际要赋值的变量，其他的名称都是上级路径
                        if (isLast)
                        {
                            currDict[varName] = cellValue;
                        }
                        // 不是最后一个名称就建立嵌套结构
                        else
                        {
                            if (!currDict.TryGetValue(varName, out object childObject))
                            {
                                childObject = new Dictionary<string, object>();
                                currDict[varName] = childObject;
                            }
                            currDict = (Dictionary<string, object>)childObject;
                        }
                    }
                }
                else if (nameList.Count == 1)
                {
                    keyValue[nameList[0]] = cellValue;
                }
            }

            return keyValue;
        }
        
        
        
        private static void ProcessRawLineData(Dictionary<string, object> classData, string parentPath)
        {
            Dictionary<string, object> needModifiedDict = new Dictionary<string, object>();
            foreach (var kvp in classData)
            {
                string varName = kvp.Key;
                string varPath;
                if (!string.IsNullOrEmpty(parentPath))
                {
                    varPath =  parentPath +"."+ varName;
                }
                else
                {
                    varPath = varName;
                }

                if (!PathToType.TryGetValue(varPath, out string varType))
                {
                    var fileAssetPath = _currFilePath.Replace(Application.dataPath,"");
                    string errorMsg = $"找不到属性 '{varPath}' 的类型定义";
                    Debug.LogError($"ExcelReader: {errorMsg} \n {_currSheetName}:{fileAssetPath}");
                    throw new ExcelDataException(errorMsg, _currFilePath, _currSheetName);
                }
                
                if (kvp.Value is Dictionary<string, object> childClass)
                {
                    ProcessRawLineData(childClass,varPath);
                    bool isEmpty = true;
                    foreach (var childKvp in childClass)
                    {
                        if (childKvp.Value != null)
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        needModifiedDict[varName] = null;
                    }
                }
                // 将成员判空
                else
                {
                    bool isEmpty = true;
                    // 遍历看类中是否有任意有效内容
                    if (kvp.Value != null)
                    {
                        switch (kvp.Value)
                        {
                            case List<object> list:
                                isEmpty = list.Count == 0;
                                break;
                            case string stringValue:
                                isEmpty = string.IsNullOrEmpty(stringValue);
                                break;
                            default:
                                isEmpty = false;
                                break;
                        }
                    }
                    
                    if (isEmpty)
                    {
                        needModifiedDict[varName] = null;
                    }
                }
                
                varType = varType.ToLower();
                // 如果被标记为list，尝试转换对象为list
                if (varType.Contains("list"))
                {
                    // 如果之前被登记修改了，设置为登记的值。
                    var currValue = needModifiedDict.TryGetValue(varName, out var modifiedValue)? modifiedValue:kvp.Value;
                    if (currValue != null)
                    {
                        needModifiedDict[varName] = new List<object>()
                        {
                            currValue
                        };
                    }
                }
            }
            // 应用登记了的修改
            foreach (var kvp in needModifiedDict)
            {
                classData[kvp.Key] = kvp.Value;
            }
        }


        private static void ApplyDefaultValueToClass(Dictionary<string, object> classData, string parentNamePath)
        {
            Dictionary<string, object> needModifiedDict = new Dictionary<string, object>();
            foreach (var kvp in classData)
            {
                string varName = kvp.Key;
                var varValue = kvp.Value;
                string currNamePath = parentNamePath + "." + varName;
                currNamePath = currNamePath.Trim('.');

                // 如果能得到默认值，说明到达了名称末尾
                if (PathToDefault.TryGetValue(currNamePath, out var defaultValue))
                {
                    // 如果不为空，则不应用默认值
                    if(varValue!=null)continue;
                    
                    // 为空则应用，并查找下一个属性
                    needModifiedDict[varName] = defaultValue;
                    continue;
                }
                
                // 如果没到达末尾，向内处理数据
                switch (varValue)
                {
                    case List<object> list:
                        foreach (var listElement in list)
                        {
                            if (listElement is Dictionary<string, object> dict)
                            {
                                ApplyDefaultValueToClass(dict, currNamePath);
                            }
                        }

                        break;
                    case Dictionary<string, object> dict:
                        ApplyDefaultValueToClass(dict, currNamePath);
                        break;
                }
            }

            foreach (var kvp in needModifiedDict)
            {
                classData[kvp.Key] = kvp.Value;
            }
        }
        private static void ApplyDefaultValue(ref object target, string namePath)
        {
            switch (target)
            {
                case Dictionary<string, object> classData:
                    ApplyDefaultValueToClass(classData, namePath);
                    break;
                case List<object> list:
                    if (list.Count == 1)
                    {
                        var listItem = list[0];
                        ApplyDefaultValue(ref listItem, namePath);
                        list[0] = listItem;
                    }
                    break;
                case string value:
                    if (PathToDefault.TryGetValue(namePath, out var defaultValue))
                    {
                        target = defaultValue;
                    }
                    break;
                default:
                    break;
            }
            
        }

        private static List<Dictionary<string, object>> GetSheetVarList(ExcelWorksheet worksheet,int keywordRowNum)
        {
            List<Dictionary<string, object>> valueList = new List<Dictionary<string, object>>();
            // 遍历每一行单元格
            int rowCount = worksheet.Dimension.End.Row;
            for (int row = keywordRowNum + 1; row <= rowCount; row++)
            {
                var lineData = GetRawLineData(worksheet, row);
                if (lineData == null)
                {
                    continue;
                }
                // 将空类变成null 并 将list类型转换为列表对象
                ProcessRawLineData(lineData,"");
                
                //判断数据结构是否为空，如果为空则跳过
                bool isEmpty = true;
                foreach (var kvp in lineData)
                {
                    if (kvp.Value != null)
                    {
                        isEmpty = false;
                        break;
                    }
                }
                if (isEmpty)
                {
                    continue;
                }
                
                // 当列表有一个以上的元素时，尝试合并数据
                if (valueList.Count > 0 && TryMerge(valueList.Last(), lineData,""))
                {
                    continue;
                }
                
                // 没有合并成功则重置数据，并添加
                lineData = GetRawLineData(worksheet, row);
                ApplyDefaultValueToClass(lineData,"");
                ProcessRawLineData(lineData,"");

                valueList.Add(lineData);
            }
            return valueList;
        }
  

        private static Dictionary<string, object> SheetToRawData(ExcelWorksheet worksheet)
        {
            // 获得每一列的关键词数据
            if (!TryGetColumKeyword(worksheet,out int keywordRowNum))
            {
                return null;
            }
            
            Dictionary<string, object> result = new Dictionary<string, object>();
            foreach (var data in AdditionalData)
            {
                result[data.Key] = data.Value;
            }

            result[RawDataKey.typeLookup] = PathToType;
            result[RawDataKey.dataList] = GetSheetVarList(worksheet, keywordRowNum);

            return result;
        }



        public static class RawDataKey
        {
            public const string className = "_className";
            public const string keyName = "_keyName";
            public const string sheetName = "_sheetName";
            public const string dataList = "_dataList";
            public const string typeLookup = "_typeLookup";
            
            // public const string saveAssetsFolderPath = "_saveAssetsFolderPath";
            // public const string mutilFile = "_mutilFile";
        }

        
        public static List<object> GetRawDataAtPath(List<string> currPath, object dataRoot)
        {
            List<object> dataList = new List<object>();

            // 如果抵达了最终路径，对内容进行解析
            if (currPath == null || currPath.Count == 0)
            {
                switch (dataRoot)
                {
                    case List<object> listRawData:
                        dataList.AddRange(listRawData);
                        break;
                    default:
                        dataList.Add(dataRoot);
                        break;
                }
                return dataList;
            }
            
            // 如果没有，则递归直到抵达最终路径
            string currValueName = currPath[0];
            List<string> nextPath = new List<string>(currPath.Count);
            nextPath.AddRange(currPath);
            nextPath.RemoveAt(0);
            
            switch (dataRoot)
            {
                case List<object> listRawData:
                    {
                        foreach (var childRawDataObject in listRawData)
                        {
                            if (childRawDataObject is Dictionary<string, object> classRawData && 
                                classRawData.TryGetValue(currValueName, out object childClass))
                            {
                                var childDataList = GetRawDataAtPath(nextPath, childClass);
                                dataList.AddRange(childDataList);
                            }
                        }
                    }
                    break;
                case Dictionary<string, object> classRawData:
                    {
                        if (classRawData.TryGetValue(currValueName, out object childClass))
                        {
                            var childDataList = GetRawDataAtPath(nextPath, childClass);
                            dataList.AddRange(childDataList);
                        }
                    }
                    break;
            }
            return dataList;
        }

        public static List<Dictionary<string, object>> GetRawData(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ExcelFileException("文件路径不能为空", filePath);
            }
            
            if (!File.Exists(filePath))
            {
                throw new ExcelFileException($"Excel文件不存在: {filePath}", filePath);
            }
            
            if (!IsSupportedFile(filePath))
            {
                throw new ExcelFileException($"不支持的文件格式: {filePath}，支持的格式: {string.Join(", ", ExcelReaderConfig.SupportedExtensions)}", filePath);
            }
            
            _currFilePath = filePath;
            List<Dictionary<string, object>> resultList = new List<Dictionary<string, object>>();
            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage package = null;
            
            try
            {
                package = new ExcelPackage(fileInfo);
            }
            catch (UnauthorizedAccessException ex)
            {
                string errorMsg = $"没有权限访问Excel文件: {filePath}";
                Debug.LogError($"{errorMsg}: {ex.Message}");
                throw new ExcelFileException(errorMsg, filePath, ex);
            }
            catch (IOException ex)
            {
                string errorMsg = $"读取Excel文件时发生IO错误: {filePath}";
                Debug.LogError($"{errorMsg}: {ex.Message}");
                throw new ExcelFileException(errorMsg, filePath, ex);
            }
            catch (Exception ex)
            {
                string errorMsg = $"无法读取Excel文件: {filePath}";
                Debug.LogError($"{errorMsg}: {ex.Message}");
                throw new ExcelFileException(errorMsg, filePath, ex);
            }
            
            using (package)
            {
                ExcelWorksheets excelSheets = null;
                try
                {
                    ExcelWorkbook excelWorkbook = package.Workbook;
                    excelSheets = excelWorkbook.Worksheets;
                    
                    if (excelSheets == null || excelSheets.Count == 0)
                    {
                        throw new ExcelFileException($"Excel文件中没有工作表: {filePath}", filePath);
                    }
                }
                catch (Exception ex)
                {
                    string errorMsg = $"无法读取Excel文件的工作表数据: {filePath}";
                    Debug.LogError($"{errorMsg}: {ex.Message}");
                    throw new ExcelFileException(errorMsg, filePath, ex);
                }
                
                foreach (var worksheet in excelSheets)
                {
                    try
                    {
                        Dictionary<string, object> rawData = new Dictionary<string, object>();
                        rawData[RawDataKey.sheetName] = worksheet.Name;
                        _currSheetName = worksheet.Name;
                        resultList.Add(rawData);

                        Dictionary<string, object> sheetData = SheetToRawData(worksheet);

                        if (sheetData != null && sheetData.Count > 0)
                        {
                            foreach (var kvp in sheetData)
                            {
                                rawData[kvp.Key] = kvp.Value;
                            }
                        }
                    }
                    catch (ExcelDataException)
                    {
                        // 重新抛出数据异常，保持原有的异常信息
                        throw;
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"处理工作表 '{worksheet.Name}' 时发生错误";
                        Debug.LogError($"{errorMsg}: {ex.Message}");
                        throw new ExcelSheetException(errorMsg, filePath, worksheet.Name, ex);
                    }
                }
            }
            return resultList;
        }
    }
}



