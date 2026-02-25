using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using UnityEngine;

namespace U0UGames.ExcelDataParser
{
    /// <summary>
    /// Excel写入异常
    /// </summary>
    public class ExcelWriteException : Exception
    {
        public string FilePath { get; }
        
        public ExcelWriteException(string message, string filePath = null, Exception innerException = null) 
            : base(message, innerException)
        {
            FilePath = filePath;
        }
    }
    /// <summary>
    /// Excel文件写入器，提供创建和保存Excel文件的功能
    /// 主要功能：
    /// 1. 创建Excel文件并写入数据
    /// 2. 支持按行列位置写入数据
    /// 3. 自动创建目录结构
    /// 4. 完善的异常处理机制
    /// </summary>
    public static class ExcelWriter
    {
        /// <summary>
        /// Excel数据位置结构体，表示Excel中的行列位置
        /// </summary>
        public struct ExcelDataPos : IEquatable<ExcelDataPos>
        {
            /// <summary>行号（从1开始）</summary>
            public readonly int row;
            /// <summary>列号（从1开始）</summary>
            public readonly int col;

            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="row">行号</param>
            /// <param name="col">列号</param>
            public ExcelDataPos(int row, int col)
            {
                this.row = row;
                this.col = col;
            }

            /// <summary>
            /// 比较两个位置是否相等
            /// </summary>
            /// <param name="other">另一个位置</param>
            /// <returns>如果相等返回true，否则返回false</returns>
            public bool Equals(ExcelDataPos other)
            {
                return this.row == other.row && this.col == other.col;
            }

            /// <summary>
            /// 重写Equals方法
            /// </summary>
            /// <param name="obj">要比较的对象</param>
            /// <returns>如果相等返回true，否则返回false</returns>
            public override bool Equals(object obj)
            {
                if (obj is ExcelDataPos)
                {
                    return Equals((ExcelDataPos)obj);
                }
                return false;
            }

            /// <summary>
            /// 重写GetHashCode方法
            /// </summary>
            /// <returns>哈希码</returns>
            public override int GetHashCode()
            {
                return HashCode.Combine(row, col);
            }
        }
        /// <summary>
        /// Excel数据容器类，用于存储要写入Excel的数据
        /// 支持按行列位置存储和访问数据
        /// </summary>
        public class ExcelData:IEnumerable<KeyValuePair<ExcelDataPos, string>>
        {
            /// <summary>位置到数据的映射字典</summary>
            private Dictionary<ExcelDataPos, string> _dataLookup = new Dictionary<ExcelDataPos, string>();
            
            /// <summary>
            /// 索引器，用于按行列位置访问数据
            /// </summary>
            /// <param name="row">行号（从1开始）</param>
            /// <param name="col">列号（从1开始）</param>
            /// <returns>指定位置的数据，如果不存在返回null</returns>
            public string this[int row, int col]
            {
                get
                {
                    if (_dataLookup.TryGetValue(new ExcelDataPos(row,col), out string target))
                    {
                        return target;
                    }
                    return null;
                }
                set
                {
                    if (value != null)
                    {
                        _dataLookup[new ExcelDataPos(row, col)] = value;
                    }
                    else
                    {
                        _dataLookup.Remove(new ExcelDataPos(row, col));
                    }
                }
            }

            /// <summary>
            /// 获取枚举器
            /// </summary>
            /// <returns>数据枚举器</returns>
            public IEnumerator<KeyValuePair<ExcelDataPos, string>> GetEnumerator()
            {
                return _dataLookup.GetEnumerator();
            }

            /// <summary>
            /// 获取非泛型枚举器
            /// </summary>
            /// <returns>枚举器</returns>
            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }
        
        
        /// <summary>
        /// Excel样式数据容器，用于存储写入Excel时需要应用的样式信息。
        /// 与 ExcelData 完全分离，作为可选参数传入 SaveFile 重载。
        /// 传入 null 时行为与无样式版本一致。
        /// </summary>
        public class ExcelStyleData
        {
            private struct CellPos : IEquatable<CellPos>
            {
                public readonly int row;
                public readonly int col;

                public CellPos(int row, int col) { this.row = row; this.col = col; }
                public bool Equals(CellPos other) => row == other.row && col == other.col;
                public override bool Equals(object obj) => obj is CellPos other && Equals(other);
                public override int GetHashCode() => HashCode.Combine(row, col);
            }

            private readonly Dictionary<int, Color> _rowBackgroundColors = new Dictionary<int, Color>();
            private readonly Dictionary<CellPos, Color> _cellBackgroundColors = new Dictionary<CellPos, Color>();
            private readonly Dictionary<int, double> _columnWidths = new Dictionary<int, double>();
            private readonly Dictionary<int, bool> _columnWrapTexts = new Dictionary<int, bool>();
            private readonly Dictionary<CellPos, bool> _cellWrapTexts = new Dictionary<CellPos, bool>();
            private int _freezeRowCount = 0;

            /// <summary>
            /// 设置整行背景色（row 从 1 开始）
            /// </summary>
            public void SetRowBackgroundColor(int row, Color color)
            {
                _rowBackgroundColors[row] = color;
            }

            /// <summary>
            /// 设置单个单元格背景色（row、col 均从 1 开始）
            /// </summary>
            public void SetCellBackgroundColor(int row, int col, Color color)
            {
                _cellBackgroundColors[new CellPos(row, col)] = color;
            }

            /// <summary>
            /// 设置列宽（col 从 1 开始，width 单位与 EPPlus 一致，通常为字符数）
            /// </summary>
            public void SetColumnWidth(int col, double width)
            {
                _columnWidths[col] = width;
            }

            /// <summary>
            /// 对整列设置自动换行（col 从 1 开始）
            /// </summary>
            public void SetColumnWrapText(int col, bool wrapText)
            {
                _columnWrapTexts[col] = wrapText;
            }

            /// <summary>
            /// 对单个单元格设置自动换行（row、col 均从 1 开始）
            /// </summary>
            public void SetCellWrapText(int row, int col, bool wrapText)
            {
                _cellWrapTexts[new CellPos(row, col)] = wrapText;
            }

            /// <summary>
            /// 冻结前 N 行（即从第 N+1 行开始滚动，rowCount 为 0 时不冻结）
            /// </summary>
            public void SetFreezeRows(int rowCount)
            {
                _freezeRowCount = rowCount;
            }

            /// <summary>
            /// 将所有样式应用到指定工作表
            /// </summary>
            internal void ApplyTo(ExcelWorksheet worksheet)
            {
                foreach (var kvp in _rowBackgroundColors)
                {
                    var rowStyle = worksheet.Row(kvp.Key).Style;
                    rowStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    rowStyle.Fill.BackgroundColor.SetColor(ToDrawingColor(kvp.Value));
                }

                foreach (var kvp in _cellBackgroundColors)
                {
                    var cellStyle = worksheet.Cells[kvp.Key.row, kvp.Key.col].Style;
                    cellStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    cellStyle.Fill.BackgroundColor.SetColor(ToDrawingColor(kvp.Value));
                }

                foreach (var kvp in _columnWidths)
                {
                    worksheet.Column(kvp.Key).Width = kvp.Value;
                }

                foreach (var kvp in _columnWrapTexts)
                {
                    worksheet.Column(kvp.Key).Style.WrapText = kvp.Value;
                }

                foreach (var kvp in _cellWrapTexts)
                {
                    worksheet.Cells[kvp.Key.row, kvp.Key.col].Style.WrapText = kvp.Value;
                }

                if (_freezeRowCount > 0)
                {
                    worksheet.View.FreezePanes(_freezeRowCount + 1, 1);
                }
            }

            private static System.Drawing.Color ToDrawingColor(Color color)
            {
                return System.Drawing.Color.FromArgb(
                    (int)(color.a * 255),
                    (int)(color.r * 255),
                    (int)(color.g * 255),
                    (int)(color.b * 255));
            }
        }

        /// <summary>
        /// 将ExcelData保存为Excel文件
        /// </summary>
        /// <param name="data">要保存的Excel数据</param>
        /// <param name="path">保存路径</param>
        /// <exception cref="ExcelWriteException">当保存过程中发生错误时抛出</exception>
        public static void SaveFile(ExcelData data, string path)
        {
            SaveFileInternal(data, null, path, "Sheet1");
        }
        
        /// <summary>
        /// 保存Excel数据到指定路径，支持自定义工作表名称
        /// </summary>
        /// <param name="data">Excel数据</param>
        /// <param name="path">保存路径</param>
        /// <param name="sheetName">工作表名称</param>
        public static void SaveFile(ExcelData data, string path, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ExcelWriteException("工作表名称不能为空", path);
            }
            SaveFileInternal(data, null, path, sheetName);
        }

        /// <summary>
        /// 将ExcelData连同样式保存为Excel文件（默认使用 Sheet1）
        /// </summary>
        /// <param name="data">要保存的Excel数据</param>
        /// <param name="styles">样式数据，传入 null 时与无样式版本行为一致</param>
        /// <param name="path">保存路径</param>
        public static void SaveFile(ExcelData data, ExcelStyleData styles, string path)
        {
            SaveFileInternal(data, styles, path, "Sheet1");
        }

        /// <summary>
        /// 将ExcelData连同样式保存为Excel文件，支持自定义工作表名称
        /// </summary>
        /// <param name="data">要保存的Excel数据</param>
        /// <param name="styles">样式数据，传入 null 时与无样式版本行为一致</param>
        /// <param name="path">保存路径</param>
        /// <param name="sheetName">工作表名称</param>
        public static void SaveFile(ExcelData data, ExcelStyleData styles, string path, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new ExcelWriteException("工作表名称不能为空", path);
            }
            SaveFileInternal(data, styles, path, sheetName);
        }

        private static void SaveFileInternal(ExcelData data, ExcelStyleData styles, string path, string sheetName)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ExcelWriteException("保存路径不能为空", path);
            }

            if (data == null)
            {
                throw new ExcelWriteException("Excel数据不能为空", path);
            }

            ExcelPackage excelPackage = null;
            try
            {
                excelPackage = new ExcelPackage();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);

                foreach (var kvp in data)
                {
                    if (string.IsNullOrEmpty(kvp.Value)) continue;

                    var pos = kvp.Key;
                    try
                    {
                        worksheet.Cells[pos.row, pos.col].Value = kvp.Value;
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"写入第 {pos.row} 行第 {pos.col} 列数据时发生错误";
                        Debug.LogError($"{errorMsg}: {ex.Message}");
                        throw new ExcelWriteException(errorMsg, path, ex);
                    }
                }

                styles?.ApplyTo(worksheet);

                string folderPath = Path.GetDirectoryName(path);
                if (folderPath != null && !Directory.Exists(folderPath))
                {
                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        string errorMsg = $"没有权限创建目录: {folderPath}";
                        Debug.LogError($"{errorMsg}: {ex.Message}");
                        throw new ExcelWriteException(errorMsg, path, ex);
                    }
                    catch (IOException ex)
                    {
                        string errorMsg = $"创建目录时发生IO错误: {folderPath}";
                        Debug.LogError($"{errorMsg}: {ex.Message}");
                        throw new ExcelWriteException(errorMsg, path, ex);
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"创建目录时发生未知错误: {folderPath}";
                        Debug.LogError($"{errorMsg}: {ex.Message}");
                        throw new ExcelWriteException(errorMsg, path, ex);
                    }
                }

                FileInfo file = new FileInfo(path);
                try
                {
                    excelPackage.SaveAs(file);
                }
                catch (UnauthorizedAccessException ex)
                {
                    string errorMsg = $"没有权限保存文件: {path}";
                    Debug.LogError($"{errorMsg}: {ex.Message}");
                    throw new ExcelWriteException(errorMsg, path, ex);
                }
                catch (IOException ex)
                {
                    string errorMsg = $"保存文件时发生IO错误: {path}";
                    Debug.LogError($"{errorMsg}: {ex.Message}");
                    throw new ExcelWriteException(errorMsg, path, ex);
                }
                catch (Exception ex)
                {
                    string errorMsg = $"保存Excel文件时发生未知错误: {path}";
                    Debug.LogError($"{errorMsg}: {ex.Message}");
                    throw new ExcelWriteException(errorMsg, path, ex);
                }
            }
            catch (ExcelWriteException)
            {
                throw;
            }
            catch (Exception ex)
            {
                string errorMsg = $"创建或操作Excel包时发生错误: {path}";
                Debug.LogError($"{errorMsg}: {ex.Message}");
                throw new ExcelWriteException(errorMsg, path, ex);
            }
            finally
            {
                excelPackage?.Dispose();
            }
        }
    }
}