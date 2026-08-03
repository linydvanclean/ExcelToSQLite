using ExcelToSQLite.Models;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToSQLite.Services
{
    public class FuelCardParserService
    {
        private int _recordCount = 0;
        private int _skippedCount = 0;
        private string _customerName = string.Empty;
        private string _networkName = string.Empty;

        public FuelCardParserService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task<List<FuelCardRecord>> ParseFuelCardAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                _recordCount = 0;
                _skippedCount = 0;
                _customerName = string.Empty;
                _networkName = string.Empty;
                var records = new List<FuelCardRecord>();

                try
                {
                    // 根据文件扩展名选择解析方式
                    string extension = Path.GetExtension(filePath).ToLower();
                    
                    if (extension == ".xlsx")
                    {
                        records = ParseXlsxFile(filePath);
                    }
                    else if (extension == ".xls")
                    {
                        records = ParseXlsFile(filePath);
                    }
                    else
                    {
                        throw new NotSupportedException($"不支持的文件格式: {extension}。请使用 .xlsx 或 .xls 格式。");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"解析加油卡文件失败: {ex.Message}", ex);
                }

                return records;
            });
        }

        /// <summary>
        /// 解析 .xlsx 文件 (使用 EPPlus)
        /// </summary>
        private List<FuelCardRecord> ParseXlsxFile(string filePath)
        {
            var allRecords = new List<FuelCardRecord>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            if (worksheet == null || worksheet.Dimension == null)
                return allRecords;

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // 提取表头信息（客户名称、网点名称）
            ExtractHeaderInfoXlsx(worksheet);

            // 找到数据表头行
            int headerRow = FindDataHeaderRowXlsx(worksheet, rowCount);
            if (headerRow == -1)
                return allRecords;

            // 建立列索引映射
            var columnIndexes = BuildColumnIndexesXlsx(worksheet, headerRow, colCount);

            // 逐行解析数据
            int dataStartRow = headerRow + 1;
            string currentCardNumber = string.Empty;

            for (int row = dataStartRow; row <= rowCount; row++)
            {
                try
                {
                    var colAValue = GetCellValueXlsx(worksheet, row, 0);
                    var colBValue = GetCellValueXlsx(worksheet, row, 1);
                    var colDValue = GetCellValueXlsx(worksheet, row, 3);

                    // 跳过汇总行
                    if (!string.IsNullOrEmpty(colAValue) && 
                        (colAValue.Contains("总账户") || 
                         colAValue.StartsWith("卡号:") || 
                         colAValue.Contains("由于可能存在的在途交易")))
                    {
                        if (colAValue.Contains("由于可能存在的在途交易"))
                            break;
                        continue;
                    }

                    // 跳过小计和总计行
                    if (!string.IsNullOrEmpty(colDValue) && 
                        (colDValue.Contains("小计") || colDValue.Contains("总计")))
                        continue;

                    // 获取卡号
                    string rowCardNumber = string.Empty;
                    if (!string.IsNullOrEmpty(colAValue) && Regex.IsMatch(colAValue, @"^\d{16,}$"))
                    {
                        rowCardNumber = colAValue.Trim();
                        currentCardNumber = rowCardNumber;
                    }
                    else if (!string.IsNullOrEmpty(colBValue) && Regex.IsMatch(colBValue, @"^\d{16,}$"))
                    {
                        rowCardNumber = colBValue.Trim();
                        currentCardNumber = rowCardNumber;
                    }

                    // 获取业务类型
                    string businessType = string.Empty;
                    if (columnIndexes.TryGetValue("BusinessType", out int businessCol))
                    {
                        businessType = GetCellValueXlsx(worksheet, row, businessCol)?.Trim() ?? string.Empty;
                    }

                    // 只处理加油记录
                    if (businessType == "加油" && !string.IsNullOrEmpty(currentCardNumber))
                    {
                        var record = ParseFuelRecordXlsx(worksheet, row, columnIndexes, currentCardNumber);
                        if (record != null)
                        {
                            allRecords.Add(record);
                            _recordCount++;
                        }
                    }
                    else if (!string.IsNullOrEmpty(businessType))
                    {
                        _skippedCount++;
                    }
                }
                catch
                {
                    _skippedCount++;
                }
            }

            return allRecords;
        }

        /// <summary>
        /// 解析 .xls 文件 (使用 NPOI)
        /// </summary>
        private List<FuelCardRecord> ParseXlsFile(string filePath)
        {
            var allRecords = new List<FuelCardRecord>();

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);

            if (sheet == null)
                return allRecords;

            int rowCount = sheet.LastRowNum + 1;
            int colCount = 0;

            // 获取最大列数
            for (int i = 0; i < rowCount; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null && row.LastCellNum > colCount)
                {
                    colCount = (int)row.LastCellNum;
                }
            }

            // 提取表头信息（客户名称、网点名称）
            ExtractHeaderInfoXls(sheet, rowCount);

            // 找到数据表头行
            int headerRow = FindDataHeaderRowXls(sheet, rowCount);
            if (headerRow == -1)
                return allRecords;

            // 建立列索引映射
            var columnIndexes = BuildColumnIndexesXls(sheet, headerRow, colCount);

            // 逐行解析数据
            int dataStartRow = headerRow + 1;
            string currentCardNumber = string.Empty;

            for (int row = dataStartRow; row < rowCount; row++)
            {
                try
                {
                    var colAValue = GetCellValueXls(sheet, row, 0);
                    var colBValue = GetCellValueXls(sheet, row, 1);
                    var colDValue = GetCellValueXls(sheet, row, 3);

                    // 跳过汇总行
                    if (!string.IsNullOrEmpty(colAValue) && 
                        (colAValue.Contains("总账户") || 
                         colAValue.StartsWith("卡号:") || 
                         colAValue.Contains("由于可能存在的在途交易")))
                    {
                        if (colAValue.Contains("由于可能存在的在途交易"))
                            break;
                        continue;
                    }

                    // 跳过小计和总计行
                    if (!string.IsNullOrEmpty(colDValue) && 
                        (colDValue.Contains("小计") || colDValue.Contains("总计")))
                        continue;

                    // 获取卡号
                    string rowCardNumber = string.Empty;
                    if (!string.IsNullOrEmpty(colAValue) && Regex.IsMatch(colAValue, @"^\d{16,}$"))
                    {
                        rowCardNumber = colAValue.Trim();
                        currentCardNumber = rowCardNumber;
                    }
                    else if (!string.IsNullOrEmpty(colBValue) && Regex.IsMatch(colBValue, @"^\d{16,}$"))
                    {
                        rowCardNumber = colBValue.Trim();
                        currentCardNumber = rowCardNumber;
                    }

                    // 获取业务类型
                    string businessType = string.Empty;
                    if (columnIndexes.TryGetValue("BusinessType", out int businessCol))
                    {
                        businessType = GetCellValueXls(sheet, row, businessCol)?.Trim() ?? string.Empty;
                    }

                    // 只处理加油记录
                    if (businessType == "加油" && !string.IsNullOrEmpty(currentCardNumber))
                    {
                        var record = ParseFuelRecordXls(sheet, row, columnIndexes, currentCardNumber);
                        if (record != null)
                        {
                            allRecords.Add(record);
                            _recordCount++;
                        }
                    }
                    else if (!string.IsNullOrEmpty(businessType))
                    {
                        _skippedCount++;
                    }
                }
                catch
                {
                    _skippedCount++;
                }
            }

            return allRecords;
        }

        #region XLSX 辅助方法

        private void ExtractHeaderInfoXlsx(ExcelWorksheet worksheet)
        {
            for (int row = 1; row <= 10; row++)
            {
                for (int col = 1; col <= 20; col++)
                {
                    var value = GetCellValueXlsx(worksheet, row, col);
                    if (!string.IsNullOrEmpty(value))
                    {
                        if (value.Contains("客户名称:"))
                        {
                            var nameValue = GetCellValueXlsx(worksheet, row, col + 1);
                            if (!string.IsNullOrEmpty(nameValue))
                                _customerName = nameValue.Trim();
                            else
                            {
                                nameValue = GetCellValueXlsx(worksheet, row, col + 2);
                                if (!string.IsNullOrEmpty(nameValue))
                                    _customerName = nameValue.Trim();
                            }
                        }
                        else if (value.Contains("网点名称:"))
                        {
                            var nameValue = GetCellValueXlsx(worksheet, row, col + 1);
                            if (!string.IsNullOrEmpty(nameValue))
                                _networkName = nameValue.Trim();
                            else
                            {
                                nameValue = GetCellValueXlsx(worksheet, row, col + 2);
                                if (!string.IsNullOrEmpty(nameValue))
                                    _networkName = nameValue.Trim();
                            }
                        }
                    }
                }
            }
        }

        private int FindDataHeaderRowXlsx(ExcelWorksheet worksheet, int rowCount)
        {
            for (int row = 1; row <= Math.Min(rowCount, 20); row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    var value = GetCellValueXlsx(worksheet, row, col);
                    if (!string.IsNullOrEmpty(value) && value.Trim() == "卡号")
                        return row;
                }
            }

            // 备用方案
            for (int row = 1; row <= Math.Min(rowCount, 20); row++)
            {
                int matchCount = 0;
                for (int col = 0; col < 20; col++)
                {
                    var value = GetCellValueXlsx(worksheet, row, col);
                    if (string.IsNullOrEmpty(value)) continue;

                    if (value.Contains("卡号")) matchCount++;
                    else if (value.Contains("时间")) matchCount++;
                    else if (value.Contains("业务类型")) matchCount++;
                    else if (value.Contains("品种")) matchCount++;
                    else if (value.Contains("数量")) matchCount++;
                    else if (value.Contains("单价")) matchCount++;
                    else if (value.Contains("金额")) matchCount++;
                    else if (value.Contains("余额")) matchCount++;
                }

                if (matchCount >= 4)
                    return row;
            }

            return -1;
        }

        private Dictionary<string, int> BuildColumnIndexesXlsx(ExcelWorksheet worksheet, int headerRow, int colCount)
        {
            var indexes = new Dictionary<string, int>();

            for (int col = 0; col < colCount; col++)
            {
                var value = GetCellValueXlsx(worksheet, headerRow, col);
                if (string.IsNullOrEmpty(value)) continue;

                value = value.Trim();

                switch (value)
                {
                    case "卡号": indexes["CardNumber"] = col; break;
                    case "时间": indexes["TransactionTime"] = col; break;
                    case "业务类型": indexes["BusinessType"] = col; break;
                    case "品种": indexes["FuelType"] = col; break;
                    case "数量": indexes["Quantity"] = col; break;
                    case "单价": indexes["UnitPrice"] = col; break;
                    case "金额(分值)":
                    case "金额": indexes["Amount"] = col; break;
                    case "奖励分值": indexes["BonusPoints"] = col; break;
                    case "优惠价": indexes["DiscountPrice"] = col; break;
                    case "余额": indexes["Balance"] = col; break;
                    case "地点": indexes["Location"] = col; break;
                    case "操作员": indexes["Operator"] = col; break;
                    case "备注": indexes["Remarks"] = col; break;
                }
            }

            return indexes;
        }

        private FuelCardRecord? ParseFuelRecordXlsx(ExcelWorksheet worksheet, int row, Dictionary<string, int> indexes, string cardNumber)
        {
            try
            {
                var record = new FuelCardRecord
                {
                    CardNumber = cardNumber,
                    CustomerName = _customerName,
                    NetworkName = _networkName,
                    CreatedAt = DateTime.Now,
                    BusinessType = "加油"
                };

                if (indexes.TryGetValue("TransactionTime", out int timeCol))
                {
                    var timeValue = GetCellValueXlsx(worksheet, row, timeCol);
                    if (!string.IsNullOrEmpty(timeValue) && DateTime.TryParse(timeValue, out DateTime time))
                        record.TransactionTime = time;
                }

                if (indexes.TryGetValue("FuelType", out int fuelCol))
                {
                    var fuelValue = GetCellValueXlsx(worksheet, row, fuelCol);
                    record.FuelType = fuelValue?.Trim() ?? string.Empty;
                }

                if (indexes.TryGetValue("Quantity", out int qtyCol))
                {
                    var qtyValue = GetCellValueXlsx(worksheet, row, qtyCol);
                    if (decimal.TryParse(qtyValue, out decimal qty))
                        record.Quantity = qty;
                }

                if (indexes.TryGetValue("UnitPrice", out int priceCol))
                {
                    var priceValue = GetCellValueXlsx(worksheet, row, priceCol);
                    if (decimal.TryParse(priceValue, out decimal price))
                        record.UnitPrice = price;
                }

                if (indexes.TryGetValue("Amount", out int amountCol))
                {
                    var amountValue = GetCellValueXlsx(worksheet, row, amountCol);
                    if (decimal.TryParse(amountValue, out decimal amount))
                        record.Amount = amount;
                }

                if (indexes.TryGetValue("BonusPoints", out int bonusCol))
                {
                    var bonusValue = GetCellValueXlsx(worksheet, row, bonusCol);
                    if (decimal.TryParse(bonusValue, out decimal bonus))
                        record.BonusPoints = bonus;
                }

                if (indexes.TryGetValue("DiscountPrice", out int discountCol))
                {
                    var discountValue = GetCellValueXlsx(worksheet, row, discountCol);
                    if (decimal.TryParse(discountValue, out decimal discount))
                        record.DiscountPrice = discount;
                }

                if (indexes.TryGetValue("Balance", out int balanceCol))
                {
                    var balanceValue = GetCellValueXlsx(worksheet, row, balanceCol);
                    if (decimal.TryParse(balanceValue, out decimal balance))
                        record.Balance = balance;
                }

                if (indexes.TryGetValue("Location", out int locationCol))
                {
                    var locationValue = GetCellValueXlsx(worksheet, row, locationCol);
                    record.Location = locationValue?.Trim() ?? string.Empty;
                }

                if (indexes.TryGetValue("Operator", out int operatorCol))
                {
                    var operatorValue = GetCellValueXlsx(worksheet, row, operatorCol);
                    record.Operator = operatorValue?.Trim() ?? string.Empty;
                }

                if (indexes.TryGetValue("Remarks", out int remarksCol))
                {
                    var remarksValue = GetCellValueXlsx(worksheet, row, remarksCol);
                    record.Remarks = remarksValue?.Trim() ?? string.Empty;
                }

                if (record.TransactionTime == default || (record.Amount == 0 && record.Quantity == 0))
                    return null;

                return record;
            }
            catch
            {
                return null;
            }
        }

        private string GetCellValueXlsx(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                var cell = worksheet.Cells[row + 1, col + 1];
                var value = cell.Value?.ToString()?.Trim() ?? string.Empty;
                
                if (string.IsNullOrEmpty(value) && cell.Merge)
                {
                    try
                    {
                        var startRow = cell.Start.Row;
                        var startCol = cell.Start.Column;
                        var startCell = worksheet.Cells[startRow, startCol];
                        return startCell.Value?.ToString()?.Trim() ?? string.Empty;
                    }
                    catch
                    {
                        return string.Empty;
                    }
                }
                
                return value;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        #region XLS 辅助方法

        private void ExtractHeaderInfoXls(ISheet sheet, int rowCount)
        {
            for (int row = 0; row < Math.Min(rowCount, 10); row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    var value = GetCellValueXls(sheet, row, col);
                    if (!string.IsNullOrEmpty(value))
                    {
                        if (value.Contains("客户名称:"))
                        {
                            var nameValue = GetCellValueXls(sheet, row, col + 1);
                            if (!string.IsNullOrEmpty(nameValue))
                                _customerName = nameValue.Trim();
                            else
                            {
                                nameValue = GetCellValueXls(sheet, row, col + 2);
                                if (!string.IsNullOrEmpty(nameValue))
                                    _customerName = nameValue.Trim();
                            }
                        }
                        else if (value.Contains("网点名称:"))
                        {
                            var nameValue = GetCellValueXls(sheet, row, col + 1);
                            if (!string.IsNullOrEmpty(nameValue))
                                _networkName = nameValue.Trim();
                            else
                            {
                                nameValue = GetCellValueXls(sheet, row, col + 2);
                                if (!string.IsNullOrEmpty(nameValue))
                                    _networkName = nameValue.Trim();
                            }
                        }
                    }
                }
            }
        }

        private int FindDataHeaderRowXls(ISheet sheet, int rowCount)
        {
            for (int row = 0; row < Math.Min(rowCount, 20); row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    var value = GetCellValueXls(sheet, row, col);
                    if (!string.IsNullOrEmpty(value) && value.Trim() == "卡号")
                        return row;
                }
            }

            for (int row = 0; row < Math.Min(rowCount, 20); row++)
            {
                int matchCount = 0;
                for (int col = 0; col < 20; col++)
                {
                    var value = GetCellValueXls(sheet, row, col);
                    if (string.IsNullOrEmpty(value)) continue;

                    if (value.Contains("卡号")) matchCount++;
                    else if (value.Contains("时间")) matchCount++;
                    else if (value.Contains("业务类型")) matchCount++;
                    else if (value.Contains("品种")) matchCount++;
                    else if (value.Contains("数量")) matchCount++;
                    else if (value.Contains("单价")) matchCount++;
                    else if (value.Contains("金额")) matchCount++;
                    else if (value.Contains("余额")) matchCount++;
                }

                if (matchCount >= 4)
                    return row;
            }

            return -1;
        }

        private Dictionary<string, int> BuildColumnIndexesXls(ISheet sheet, int headerRow, int colCount)
        {
            var indexes = new Dictionary<string, int>();

            for (int col = 0; col < colCount; col++)
            {
                var value = GetCellValueXls(sheet, headerRow, col);
                if (string.IsNullOrEmpty(value)) continue;

                value = value.Trim();

                switch (value)
                {
                    case "卡号": indexes["CardNumber"] = col; break;
                    case "时间": indexes["TransactionTime"] = col; break;
                    case "业务类型": indexes["BusinessType"] = col; break;
                    case "品种": indexes["FuelType"] = col; break;
                    case "数量": indexes["Quantity"] = col; break;
                    case "单价": indexes["UnitPrice"] = col; break;
                    case "金额(分值)":
                    case "金额": indexes["Amount"] = col; break;
                    case "奖励分值": indexes["BonusPoints"] = col; break;
                    case "优惠价": indexes["DiscountPrice"] = col; break;
                    case "余额": indexes["Balance"] = col; break;
                    case "地点": indexes["Location"] = col; break;
                    case "操作员": indexes["Operator"] = col; break;
                    case "备注": indexes["Remarks"] = col; break;
                }
            }

            return indexes;
        }

        private FuelCardRecord? ParseFuelRecordXls(ISheet sheet, int row, Dictionary<string, int> indexes, string cardNumber)
        {
            try
            {
                var record = new FuelCardRecord
                {
                    CardNumber = cardNumber,
                    CustomerName = _customerName,
                    NetworkName = _networkName,
                    CreatedAt = DateTime.Now,
                    BusinessType = "加油"
                };

                if (indexes.TryGetValue("TransactionTime", out int timeCol))
                {
                    var timeValue = GetCellValueXls(sheet, row, timeCol);
                    if (!string.IsNullOrEmpty(timeValue) && DateTime.TryParse(timeValue, out DateTime time))
                        record.TransactionTime = time;
                }

                if (indexes.TryGetValue("FuelType", out int fuelCol))
                {
                    var fuelValue = GetCellValueXls(sheet, row, fuelCol);
                    record.FuelType = fuelValue?.Trim() ?? string.Empty;
                }

                if (indexes.TryGetValue("Quantity", out int qtyCol))
                {
                    var qtyValue = GetCellValueXls(sheet, row, qtyCol);
                    if (decimal.TryParse(qtyValue, out decimal qty))
                        record.Quantity = qty;
                }

                if (indexes.TryGetValue("UnitPrice", out int priceCol))
                {
                    var priceValue = GetCellValueXls(sheet, row, priceCol);
                    if (decimal.TryParse(priceValue, out decimal price))
                        record.UnitPrice = price;
                }

                if (indexes.TryGetValue("Amount", out int amountCol))
                {
                    var amountValue = GetCellValueXls(sheet, row, amountCol);
                    if (decimal.TryParse(amountValue, out decimal amount))
                        record.Amount = amount;
                }

                if (indexes.TryGetValue("BonusPoints", out int bonusCol))
                {
                    var bonusValue = GetCellValueXls(sheet, row, bonusCol);
                    if (decimal.TryParse(bonusValue, out decimal bonus))
                        record.BonusPoints = bonus;
                }

                if (indexes.TryGetValue("DiscountPrice", out int discountCol))
                {
                    var discountValue = GetCellValueXls(sheet, row, discountCol);
                    if (decimal.TryParse(discountValue, out decimal discount))
                        record.DiscountPrice = discount;
                }

                if (indexes.TryGetValue("Balance", out int balanceCol))
                {
                    var balanceValue = GetCellValueXls(sheet, row, balanceCol);
                    if (decimal.TryParse(balanceValue, out decimal balance))
                        record.Balance = balance;
                }

                if (indexes.TryGetValue("Location", out int locationCol))
                {
                    var locationValue = GetCellValueXls(sheet, row, locationCol);
                    record.Location = locationValue?.Trim() ?? string.Empty;
                }

                if (indexes.TryGetValue("Operator", out int operatorCol))
                {
                    var operatorValue = GetCellValueXls(sheet, row, operatorCol);
                    record.Operator = operatorValue?.Trim() ?? string.Empty;
                }

                if (indexes.TryGetValue("Remarks", out int remarksCol))
                {
                    var remarksValue = GetCellValueXls(sheet, row, remarksCol);
                    record.Remarks = remarksValue?.Trim() ?? string.Empty;
                }

                if (record.TransactionTime == default || (record.Amount == 0 && record.Quantity == 0))
                    return null;

                return record;
            }
            catch
            {
                return null;
            }
        }

        private string GetCellValueXls(ISheet sheet, int row, int col)
        {
            try
            {
                var rowData = sheet.GetRow(row);
                if (rowData == null) return string.Empty;

                var cell = rowData.GetCell(col);
                if (cell == null) return string.Empty;

                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue?.Trim() ?? string.Empty;
                    case CellType.Numeric:
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            var dateValue = cell.DateCellValue;
                            return dateValue?.ToString("yyyy-MM-dd HH:mm:ss") ?? string.Empty;
                        }
                        return cell.NumericCellValue.ToString();
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Formula:
                        try
                        {
                            return cell.StringCellValue?.Trim() ?? string.Empty;
                        }
                        catch
                        {
                            return cell.NumericCellValue.ToString();
                        }
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion

        public string GetSummaryInfo()
        {
            return $"客户: {_customerName}, 网点: {_networkName}, 记录数: {_recordCount}";
        }
    }
}