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
        private string _operatorName = string.Empty;
        private string _printDate = string.Empty;
        private DateTime _startDate;
        private DateTime _endDate;

        // 列索引映射 - 根据样表
        private class ColumnIndexes
        {
            public int CardNumber { get; set; } = -1;      // A列
            public int TransactionTime { get; set; } = -1; // B列
            public int BusinessType { get; set; } = -1;    // D列
            public int FuelType { get; set; } = -1;        // E列
            public int Quantity { get; set; } = -1;        // F列
            public int UnitPrice { get; set; } = -1;       // H列
            public int Amount { get; set; } = -1;          // J列 (金额(分值))
            public int BonusPoints { get; set; } = -1;     // K列 (奖励分值)
            public int DiscountPrice { get; set; } = -1;   // M列 (优惠价)
            public int Balance { get; set; } = -1;         // O列 (余额)
            public int Location { get; set; } = -1;        // Q列 (地点)
            public int Operator { get; set; } = -1;        // S列 (操作员)
            public int Remarks { get; set; } = -1;         // T列 (备注)
        }

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
                _operatorName = string.Empty;
                _printDate = string.Empty;
                var records = new List<FuelCardRecord>();

                try
                {
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

        #region XLSX 解析

        private List<FuelCardRecord> ParseXlsxFile(string filePath)
        {
            var allRecords = new List<FuelCardRecord>();

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            if (worksheet == null || worksheet.Dimension == null)
                return allRecords;

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // 提取表头信息
            ExtractHeaderInfoXlsx(worksheet);

            // 查找数据表头行（包含"卡号"的行）
            int headerRow = FindDataHeaderRowXlsx(worksheet, rowCount);
            if (headerRow == -1)
            {
                throw new Exception("未找到数据表头行（包含'卡号'列）");
            }

            // 建立列索引映射
            var indexes = BuildColumnIndexesXlsx(worksheet, headerRow, colCount);
            ValidateColumnIndexes(indexes);

            // 解析数据行
            string currentCardNumber = string.Empty;
            int dataStartRow = headerRow + 1;

            for (int row = dataStartRow; row <= rowCount; row++)
            {
                try
                {
                    // 获取关键列的值
                    var colAValue = GetCellValueXlsx(worksheet, row, 0); // A列
                    var colBValue = GetCellValueXlsx(worksheet, row, 1); // B列
                    var colDValue = GetCellValueXlsx(worksheet, row, 3); // D列

                    // 检查是否到达无效数据行
                    if (colAValue.Contains("由于可能存在的在途交易"))
                        break;

                    // 跳过卡汇总行（包含"卡号:"的行）
                    if (colAValue.StartsWith("卡号:") || colBValue.StartsWith("卡号:"))
                        continue;

                    // 跳过总账户行
                    if (colAValue == "总账户" || colAValue.Contains("总账户"))
                        continue;

                    // 跳过小计和总计行
                    if (colDValue.Contains("小计") || colDValue.Contains("总计"))
                        continue;

                    // 获取卡号（优先从A列获取16位以上数字）
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
                    string businessType = GetCellValueXlsx(worksheet, row, indexes.BusinessType)?.Trim() ?? string.Empty;

                    // 只处理圈存和加油记录
                    if ((businessType == "圈存" || businessType == "加油") && !string.IsNullOrEmpty(currentCardNumber))
                    {
                        var record = ParseFuelRecordXlsx(worksheet, row, indexes, currentCardNumber, businessType);
                        if (record != null)
                        {
                            allRecords.Add(record);
                            _recordCount++;
                        }
                    }
                    else if (!string.IsNullOrEmpty(businessType) && businessType != "圈提")
                    {
                        _skippedCount++;
                    }
                }
                catch
                {
                    _skippedCount++;
                    // 继续处理下一行
                }
            }

            return allRecords;
        }

        private void ExtractHeaderInfoXlsx(ExcelWorksheet worksheet)
        {
            for (int row = 1; row <= 10; row++)
            {
                for (int col = 1; col <= 20; col++)
                {
                    var value = GetCellValueXlsx(worksheet, row, col);
                    if (string.IsNullOrEmpty(value)) continue;

                    if (value.Contains("客户名称:"))
                    {
                        var nameValue = GetCellValueXlsx(worksheet, row, col + 1);
                        if (!string.IsNullOrEmpty(nameValue))
                            _customerName = nameValue.Trim();
                    }
                    else if (value.Contains("网点名称:"))
                    {
                        var nameValue = GetCellValueXlsx(worksheet, row, col + 1);
                        if (!string.IsNullOrEmpty(nameValue))
                            _networkName = nameValue.Trim();
                    }
                    else if (value.Contains("操 作 员:"))
                    {
                        var nameValue = GetCellValueXlsx(worksheet, row, col + 1);
                        if (!string.IsNullOrEmpty(nameValue))
                            _operatorName = nameValue.Trim();
                    }
                    else if (value.Contains("打印日期:"))
                    {
                        var dateValue = GetCellValueXlsx(worksheet, row, col + 1);
                        if (!string.IsNullOrEmpty(dateValue))
                            _printDate = dateValue.Trim();
                    }
                    else if (value.Contains("起止时间:"))
                    {
                        var timeValue = GetCellValueXlsx(worksheet, row, col + 1);
                        if (!string.IsNullOrEmpty(timeValue))
                        {
                            var parts = timeValue.Split(new[] { "----" }, StringSplitOptions.None);
                            if (parts.Length == 2)
                            {
                                if (DateTime.TryParse(parts[0], out DateTime start))
                                    _startDate = start;
                                if (DateTime.TryParse(parts[1], out DateTime end))
                                    _endDate = end;
                            }
                        }
                    }
                }
            }
        }

        private int FindDataHeaderRowXlsx(ExcelWorksheet worksheet, int rowCount)
        {
            // 精确查找包含"卡号"且后续行有数据的行
            for (int row = 1; row <= Math.Min(rowCount, 30); row++)
            {
                var value = GetCellValueXlsx(worksheet, row, 0);
                if (!string.IsNullOrEmpty(value) && value.Trim() == "卡号")
                {
                    // 验证下一行是否包含数据
                    if (row + 1 <= rowCount)
                    {
                        var nextRowValue = GetCellValueXlsx(worksheet, row + 1, 3);
                        if (!string.IsNullOrEmpty(nextRowValue) && 
                            (nextRowValue.Contains("圈存") || nextRowValue.Contains("加油")))
                        {
                            return row;
                        }
                    }
                }
            }

            // 备用方案：查找包含多个表头关键词的行
            for (int row = 1; row <= Math.Min(rowCount, 30); row++)
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

                if (matchCount >= 5)
                    return row;
            }

            return -1;
        }

        private ColumnIndexes BuildColumnIndexesXlsx(ExcelWorksheet worksheet, int headerRow, int colCount)
        {
            var indexes = new ColumnIndexes();

            for (int col = 0; col < colCount; col++)
            {
                var value = GetCellValueXlsx(worksheet, headerRow, col);
                if (string.IsNullOrEmpty(value)) continue;

                value = value.Trim();

                switch (value)
                {
                    case "卡号":
                        indexes.CardNumber = col;
                        break;
                    case "时间":
                        indexes.TransactionTime = col;
                        break;
                    case "业务类型":
                        indexes.BusinessType = col;
                        break;
                    case "品种":
                        indexes.FuelType = col;
                        break;
                    case "数量":
                        indexes.Quantity = col;
                        break;
                    case "单价":
                        indexes.UnitPrice = col;
                        break;
                    case "金额(分值)":
                        indexes.Amount = col;
                        break;
                    case "奖励分值":
                        indexes.BonusPoints = col;
                        break;
                    case "优惠价":
                        indexes.DiscountPrice = col;
                        break;
                    case "余额":
                        indexes.Balance = col;
                        break;
                    case "地点":
                        indexes.Location = col;
                        break;
                    case "操作员":
                        indexes.Operator = col;
                        break;
                    case "备注":
                        indexes.Remarks = col;
                        break;
                }
            }

            return indexes;
        }

        private void ValidateColumnIndexes(ColumnIndexes indexes)
        {
            var missingColumns = new List<string>();
            if (indexes.CardNumber == -1) missingColumns.Add("卡号");
            if (indexes.TransactionTime == -1) missingColumns.Add("时间");
            if (indexes.BusinessType == -1) missingColumns.Add("业务类型");
            if (indexes.Amount == -1) missingColumns.Add("金额(分值)");

            if (missingColumns.Count > 0)
            {
                throw new Exception($"未找到必需的列: {string.Join(", ", missingColumns)}");
            }
        }

        private FuelCardRecord? ParseFuelRecordXlsx(
            ExcelWorksheet worksheet, 
            int row, 
            ColumnIndexes indexes, 
            string cardNumber,
            string businessType)
        {
            try
            {
                var record = new FuelCardRecord
                {
                    CardNumber = cardNumber,
                    CustomerName = _customerName,
                    NetworkName = _networkName,
                    CreatedAt = DateTime.Now,
                    BusinessType = businessType
                };

                // 交易时间
                if (indexes.TransactionTime != -1)
                {
                    var timeValue = GetCellValueXlsx(worksheet, row, indexes.TransactionTime);
                    if (!string.IsNullOrEmpty(timeValue) && DateTime.TryParse(timeValue, out DateTime time))
                        record.TransactionTime = time;
                }

                // 油品类型
                if (indexes.FuelType != -1)
                {
                    var fuelValue = GetCellValueXlsx(worksheet, row, indexes.FuelType);
                    record.FuelType = fuelValue?.Trim() ?? string.Empty;
                }

                // 数量
                if (indexes.Quantity != -1)
                {
                    var qtyValue = GetCellValueXlsx(worksheet, row, indexes.Quantity);
                    if (decimal.TryParse(qtyValue, out decimal qty))
                        record.Quantity = qty;
                }

                // 单价
                if (indexes.UnitPrice != -1)
                {
                    var priceValue = GetCellValueXlsx(worksheet, row, indexes.UnitPrice);
                    if (decimal.TryParse(priceValue, out decimal price))
                        record.UnitPrice = price;
                }

                // 金额（分值）
                if (indexes.Amount != -1)
                {
                    var amountValue = GetCellValueXlsx(worksheet, row, indexes.Amount);
                    if (decimal.TryParse(amountValue, out decimal amount))
                    {
                        // 金额以分为单位，转换为元
                        record.Amount = amount / 100;
                    }
                }

                // 奖励分值
                if (indexes.BonusPoints != -1)
                {
                    var bonusValue = GetCellValueXlsx(worksheet, row, indexes.BonusPoints);
                    if (decimal.TryParse(bonusValue, out decimal bonus))
                        record.BonusPoints = bonus;
                }

                // 优惠价
                if (indexes.DiscountPrice != -1)
                {
                    var discountValue = GetCellValueXlsx(worksheet, row, indexes.DiscountPrice);
                    if (decimal.TryParse(discountValue, out decimal discount))
                        record.DiscountPrice = discount;
                }

                // 余额
                if (indexes.Balance != -1)
                {
                    var balanceValue = GetCellValueXlsx(worksheet, row, indexes.Balance);
                    if (decimal.TryParse(balanceValue, out decimal balance))
                        record.Balance = balance;
                }

                // 地点
                if (indexes.Location != -1)
                {
                    var locationValue = GetCellValueXlsx(worksheet, row, indexes.Location);
                    record.Location = locationValue?.Trim() ?? string.Empty;
                }

                // 操作员
                if (indexes.Operator != -1)
                {
                    var operatorValue = GetCellValueXlsx(worksheet, row, indexes.Operator);
                    record.Operator = operatorValue?.Trim() ?? string.Empty;
                }

                // 备注
                if (indexes.Remarks != -1)
                {
                    var remarksValue = GetCellValueXlsx(worksheet, row, indexes.Remarks);
                    record.Remarks = remarksValue?.Trim() ?? string.Empty;
                }

                // 验证记录有效性
                if (record.TransactionTime == default || record.Amount == 0)
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
                if (row < 0 || col < 0) return string.Empty;
                
                var cell = worksheet.Cells[row + 1, col + 1];
                var value = cell.Value?.ToString()?.Trim() ?? string.Empty;
                
                // 处理合并单元格
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

        #region XLS 解析（使用NPOI）

        private List<FuelCardRecord> ParseXlsFile(string filePath)
        {
            var allRecords = new List<FuelCardRecord>();

            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);

            if (sheet == null)
                return allRecords;

            int rowCount = sheet.LastRowNum + 1;
            int colCount = GetMaxColumnCountXls(sheet, rowCount);

            // 提取表头信息
            ExtractHeaderInfoXls(sheet, rowCount);

            // 查找数据表头行
            int headerRow = FindDataHeaderRowXls(sheet, rowCount);
            if (headerRow == -1)
                throw new Exception("未找到数据表头行（包含'卡号'列）");

            // 建立列索引映射
            var indexes = BuildColumnIndexesXls(sheet, headerRow, colCount);
            ValidateColumnIndexes(indexes);

            // 解析数据行
            string currentCardNumber = string.Empty;
            int dataStartRow = headerRow + 1;

            for (int row = dataStartRow; row < rowCount; row++)
            {
                try
                {
                    var colAValue = GetCellValueXls(sheet, row, 0);
                    var colBValue = GetCellValueXls(sheet, row, 1);
                    var colDValue = GetCellValueXls(sheet, row, 3);

                    // 检查是否到达无效数据行
                    if (colAValue.Contains("由于可能存在的在途交易"))
                        break;

                    // 跳过卡汇总行
                    if (colAValue.StartsWith("卡号:") || colBValue.StartsWith("卡号:"))
                        continue;

                    // 跳过总账户行
                    if (colAValue == "总账户" || colAValue.Contains("总账户"))
                        continue;

                    // 跳过小计和总计行
                    if (colDValue.Contains("小计") || colDValue.Contains("总计"))
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
                    string businessType = GetCellValueXls(sheet, row, indexes.BusinessType)?.Trim() ?? string.Empty;

                    // 只处理圈存和加油记录
                    if ((businessType == "圈存" || businessType == "加油") && !string.IsNullOrEmpty(currentCardNumber))
                    {
                        var record = ParseFuelRecordXls(sheet, row, indexes, currentCardNumber, businessType);
                        if (record != null)
                        {
                            allRecords.Add(record);
                            _recordCount++;
                        }
                    }
                }
                catch
                {
                    _skippedCount++;
                }
            }

            return allRecords;
        }

        private int GetMaxColumnCountXls(ISheet sheet, int rowCount)
        {
            int maxCols = 0;
            for (int i = 0; i < rowCount; i++)
            {
                var row = sheet.GetRow(i);
                if (row != null && row.LastCellNum > maxCols)
                    maxCols = (int)row.LastCellNum;
            }
            return maxCols;
        }

        private void ExtractHeaderInfoXls(ISheet sheet, int rowCount)
        {
            for (int row = 0; row < Math.Min(rowCount, 10); row++)
            {
                for (int col = 0; col < 20; col++)
                {
                    var value = GetCellValueXls(sheet, row, col);
                    if (string.IsNullOrEmpty(value)) continue;

                    if (value.Contains("客户名称:"))
                    {
                        var nameValue = GetCellValueXls(sheet, row, col + 1);
                        if (!string.IsNullOrEmpty(nameValue))
                            _customerName = nameValue.Trim();
                    }
                    else if (value.Contains("网点名称:"))
                    {
                        var nameValue = GetCellValueXls(sheet, row, col + 1);
                        if (!string.IsNullOrEmpty(nameValue))
                            _networkName = nameValue.Trim();
                    }
                    else if (value.Contains("操 作 员:"))
                    {
                        var nameValue = GetCellValueXls(sheet, row, col + 1);
                        if (!string.IsNullOrEmpty(nameValue))
                            _operatorName = nameValue.Trim();
                    }
                    else if (value.Contains("打印日期:"))
                    {
                        var dateValue = GetCellValueXls(sheet, row, col + 1);
                        if (!string.IsNullOrEmpty(dateValue))
                            _printDate = dateValue.Trim();
                    }
                }
            }
        }

        private int FindDataHeaderRowXls(ISheet sheet, int rowCount)
        {
            for (int row = 0; row < Math.Min(rowCount, 30); row++)
            {
                var value = GetCellValueXls(sheet, row, 0);
                if (!string.IsNullOrEmpty(value) && value.Trim() == "卡号")
                {
                    if (row + 1 < rowCount)
                    {
                        var nextRowValue = GetCellValueXls(sheet, row + 1, 3);
                        if (!string.IsNullOrEmpty(nextRowValue) && 
                            (nextRowValue.Contains("圈存") || nextRowValue.Contains("加油")))
                        {
                            return row;
                        }
                    }
                }
            }

            // 备用方案
            for (int row = 0; row < Math.Min(rowCount, 30); row++)
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

                if (matchCount >= 5)
                    return row;
            }

            return -1;
        }

        private ColumnIndexes BuildColumnIndexesXls(ISheet sheet, int headerRow, int colCount)
        {
            var indexes = new ColumnIndexes();

            for (int col = 0; col < colCount; col++)
            {
                var value = GetCellValueXls(sheet, headerRow, col);
                if (string.IsNullOrEmpty(value)) continue;

                value = value.Trim();

                switch (value)
                {
                    case "卡号":
                        indexes.CardNumber = col;
                        break;
                    case "时间":
                        indexes.TransactionTime = col;
                        break;
                    case "业务类型":
                        indexes.BusinessType = col;
                        break;
                    case "品种":
                        indexes.FuelType = col;
                        break;
                    case "数量":
                        indexes.Quantity = col;
                        break;
                    case "单价":
                        indexes.UnitPrice = col;
                        break;
                    case "金额(分值)":
                        indexes.Amount = col;
                        break;
                    case "奖励分值":
                        indexes.BonusPoints = col;
                        break;
                    case "优惠价":
                        indexes.DiscountPrice = col;
                        break;
                    case "余额":
                        indexes.Balance = col;
                        break;
                    case "地点":
                        indexes.Location = col;
                        break;
                    case "操作员":
                        indexes.Operator = col;
                        break;
                    case "备注":
                        indexes.Remarks = col;
                        break;
                }
            }

            return indexes;
        }

        private FuelCardRecord? ParseFuelRecordXls(
            ISheet sheet, 
            int row, 
            ColumnIndexes indexes, 
            string cardNumber,
            string businessType)
        {
            try
            {
                var record = new FuelCardRecord
                {
                    CardNumber = cardNumber,
                    CustomerName = _customerName,
                    NetworkName = _networkName,
                    CreatedAt = DateTime.Now,
                    BusinessType = businessType
                };

                if (indexes.TransactionTime != -1)
                {
                    var timeValue = GetCellValueXls(sheet, row, indexes.TransactionTime);
                    if (!string.IsNullOrEmpty(timeValue) && DateTime.TryParse(timeValue, out DateTime time))
                        record.TransactionTime = time;
                }

                if (indexes.FuelType != -1)
                {
                    var fuelValue = GetCellValueXls(sheet, row, indexes.FuelType);
                    record.FuelType = fuelValue?.Trim() ?? string.Empty;
                }

                if (indexes.Quantity != -1)
                {
                    var qtyValue = GetCellValueXls(sheet, row, indexes.Quantity);
                    if (decimal.TryParse(qtyValue, out decimal qty))
                        record.Quantity = qty;
                }

                if (indexes.UnitPrice != -1)
                {
                    var priceValue = GetCellValueXls(sheet, row, indexes.UnitPrice);
                    if (decimal.TryParse(priceValue, out decimal price))
                        record.UnitPrice = price;
                }

                if (indexes.Amount != -1)
                {
                    var amountValue = GetCellValueXls(sheet, row, indexes.Amount);
                    if (decimal.TryParse(amountValue, out decimal amount))
                    {
                        record.Amount = amount / 100; // 分转元
                    }
                }

                if (indexes.BonusPoints != -1)
                {
                    var bonusValue = GetCellValueXls(sheet, row, indexes.BonusPoints);
                    if (decimal.TryParse(bonusValue, out decimal bonus))
                        record.BonusPoints = bonus;
                }

                if (indexes.DiscountPrice != -1)
                {
                    var discountValue = GetCellValueXls(sheet, row, indexes.DiscountPrice);
                    if (decimal.TryParse(discountValue, out decimal discount))
                        record.DiscountPrice = discount;
                }

                if (indexes.Balance != -1)
                {
                    var balanceValue = GetCellValueXls(sheet, row, indexes.Balance);
                    if (decimal.TryParse(balanceValue, out decimal balance))
                        record.Balance = balance;
                }

                if (indexes.Location != -1)
                {
                    var locationValue = GetCellValueXls(sheet, row, indexes.Location);
                    record.Location = locationValue?.Trim() ?? string.Empty;
                }

                if (indexes.Operator != -1)
                {
                    var operatorValue = GetCellValueXls(sheet, row, indexes.Operator);
                    record.Operator = operatorValue?.Trim() ?? string.Empty;
                }

                if (indexes.Remarks != -1)
                {
                    var remarksValue = GetCellValueXls(sheet, row, indexes.Remarks);
                    record.Remarks = remarksValue?.Trim() ?? string.Empty;
                }

                if (record.TransactionTime == default || record.Amount == 0)
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

        #region 公共方法

        public string GetSummaryInfo()
        {
            return $"客户: {_customerName}, 网点: {_networkName}, 记录数: {_recordCount}, 跳过: {_skippedCount}";
        }

        #endregion
    }
}