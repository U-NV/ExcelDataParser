using System;
using System.Collections.Generic;
using UnityEngine;

namespace U0UGames.Excel
{
    /// <summary>
    /// ExcelReader使用示例
    /// 展示如何使用优化后的ExcelReader功能
    /// </summary>
    public static class ExcelReaderExample
    {
        /// <summary>
        /// 基础使用示例
        /// </summary>
        public static void BasicUsageExample()
        {
            try
            {
                // 配置ExcelReader
                ExcelReaderConfig.CaseSensitive = false;
                ExcelReaderConfig.MaxRows = 5000;
                ExcelReaderConfig.MaxColumns = 100;
                ExcelReaderConfig.StrictValidation = true;
                
                // 添加自定义关键词
                ExcelReaderConfig.CustomKeywords.Add("description");
                ExcelReaderConfig.CustomKeywords.Add("category");
                
                // 读取Excel文件
                string filePath = "Assets/Data/GameConfig.xlsx";
                var rawData = ExcelReader.GetRawData(filePath);
                
                Debug.Log($"成功读取 {rawData.Count} 个工作表");
                
                // 处理数据
                foreach (var sheetData in rawData)
                {
                    string sheetName = sheetData[ExcelReader.RawDataKey.sheetName] as string;
                    Debug.Log($"处理工作表: {sheetName}");
                    
                    var dataList = sheetData[ExcelReader.RawDataKey.dataList] as List<Dictionary<string, object>>;
                    Debug.Log($"工作表包含 {dataList.Count} 行数据");
                }
            }
            catch (ExcelFileException ex)
            {
                Debug.LogError($"文件错误: {ex.Message}");
            }
            catch (ExcelSheetException ex)
            {
                Debug.LogError($"工作表错误: {ex.Message}");
            }
            catch (ExcelDataException ex)
            {
                Debug.LogError($"数据错误: {ex.Message} (行: {ex.Row}, 列: {ex.Column})");
            }
            finally
            {
                // 清理缓存
                ExcelReader.ClearCache();
            }
        }
        
        /// <summary>
        /// 高级配置示例
        /// </summary>
        public static void AdvancedConfigurationExample()
        {
            // 自定义配置
            ExcelReaderConfig.CaseSensitive = true;
            ExcelReaderConfig.MaxRows = 10000;
            ExcelReaderConfig.MaxColumns = 200;
            ExcelReaderConfig.MaxNestingLevel = 5;
            ExcelReaderConfig.StrictValidation = false;
            
            // 自定义支持的文件格式
            ExcelReaderConfig.SupportedExtensions = new[] { ".xlsx", ".xls", ".csv" };
            
            // 自定义变量名验证规则
            ExcelReaderConfig.VariableNamePattern = @"^[a-zA-Z][a-zA-Z0-9_]*$";
            
            // 自定义类型列表
            ExcelReaderConfig.ValidTypes = new[] { 
                "string", "int", "float", "double", "bool", 
                "list", "array", "object", "vector3", "color" 
            };
            
            // 添加更多自定义关键词
            ExcelReaderConfig.CustomKeywords.AddRange(new[] {
                "description", "category", "tags", "metadata", "config"
            });
            
            Debug.Log("ExcelReader配置已更新");
        }
        
        /// <summary>
        /// 错误处理示例
        /// </summary>
        public static void ErrorHandlingExample()
        {
            try
            {
                // 尝试读取不存在的文件
                ExcelReader.GetRawData("NonExistentFile.xlsx");
            }
            catch (ExcelFileException ex)
            {
                Debug.LogError($"文件错误: {ex.Message}");
                Debug.LogError($"文件路径: {ex.FilePath}");
            }
            
            try
            {
                // 尝试读取不支持的文件格式
                ExcelReader.GetRawData("Document.pdf");
            }
            catch (ExcelFileException ex)
            {
                Debug.LogError($"格式错误: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 性能优化示例
        /// </summary>
        public static void PerformanceExample()
        {
            // 批量处理多个文件
            string[] filePaths = {
                "Assets/Data/Items.xlsx",
                "Assets/Data/Skills.xlsx",
                "Assets/Data/Levels.xlsx"
            };
            
            foreach (string filePath in filePaths)
            {
                try
                {
                    var startTime = DateTime.Now;
                    var data = ExcelReader.GetRawData(filePath);
                    var endTime = DateTime.Now;
                    
                    Debug.Log($"文件 {filePath} 处理完成，耗时: {(endTime - startTime).TotalMilliseconds}ms");
                    
                    // 处理完成后清理缓存
                    ExcelReader.ClearCache();
                }
                catch (Exception ex)
                {
                    Debug.LogError($"处理文件 {filePath} 时出错: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 数据验证示例
        /// </summary>
        public static void DataValidationExample()
        {
            // 启用严格验证
            ExcelReaderConfig.StrictValidation = true;
            
            try
            {
                var data = ExcelReader.GetRawData("Assets/Data/ValidatedData.xlsx");
                
                // 验证数据完整性
                foreach (var sheetData in data)
                {
                    var typeLookup = sheetData[ExcelReader.RawDataKey.typeLookup] as Dictionary<string, string>;
                    var dataList = sheetData[ExcelReader.RawDataKey.dataList] as List<Dictionary<string, object>>;
                    
                    Debug.Log($"工作表类型定义数量: {typeLookup.Count}");
                    Debug.Log($"工作表数据行数: {dataList.Count}");
                }
            }
            catch (ExcelDataException ex)
            {
                Debug.LogError($"数据验证失败: {ex.Message}");
                Debug.LogError($"位置: 工作表 {ex.SheetName}, 行 {ex.Row}, 列 {ex.Column}");
            }
        }
        
        /// <summary>
        /// ExcelWriter使用示例
        /// </summary>
        public static void ExcelWriterExample()
        {
            try
            {
                // 创建Excel数据
                var excelData = new ExcelWriter.ExcelData();
                
                // 设置表头
                excelData[1, 1] = "ID";
                excelData[1, 2] = "Name";
                excelData[1, 3] = "Level";
                excelData[1, 4] = "Score";
                
                // 设置数据行
                excelData[2, 1] = "1";
                excelData[2, 2] = "Player1";
                excelData[2, 3] = "10";
                excelData[2, 4] = "1500";
                
                excelData[3, 1] = "2";
                excelData[3, 2] = "Player2";
                excelData[3, 3] = "15";
                excelData[3, 4] = "2200";
                
                // 保存文件
                var filePath = "Assets/Data/PlayerData.xlsx";
                ExcelWriter.SaveFile(excelData, filePath, "Players");
                
                Debug.Log($"Excel文件已保存到: {filePath}");
            }
            catch (ExcelWriteException ex)
            {
                Debug.LogError($"保存Excel文件失败: {ex.Message}");
                Debug.LogError($"文件路径: {ex.FilePath}");
            }
        }
    }
}
