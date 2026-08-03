using System;

namespace ExcelToSQLite.Models
{
    public class FuelCardRecord
    {
        public int Id { get; set; }
        public string? CardNumber { get; set; }        // 卡号
        public DateTime TransactionTime { get; set; }   // 交易时间
        public string? BusinessType { get; set; }       // 业务类型：圈存、圈提、加油等
        public string? FuelType { get; set; }           // 油品类型
        public decimal Quantity { get; set; }           // 数量（升）
        public decimal UnitPrice { get; set; }          // 单价
        public decimal Amount { get; set; }             // 金额
        public decimal BonusPoints { get; set; }        // 奖励分值
        public decimal DiscountPrice { get; set; }      // 优惠价
        public decimal Balance { get; set; }            // 余额
        public string? Location { get; set; }           // 地点
        public string? Operator { get; set; }           // 操作员
        public string? Remarks { get; set; }            // 备注
        public string? CustomerName { get; set; }       // 客户名称（从表头提取）
        public string? NetworkName { get; set; }        // 网点名称（从表头提取）
        public DateTime CreatedAt { get; set; }         // 导入时间
        
        // 辅助属性
        public string TransactionTimeDisplay => TransactionTime.ToString("yyyy-MM-dd HH:mm:ss");
        public string AmountDisplay => Amount.ToString("F2") + " 元";
        public string QuantityDisplay => Quantity.ToString("F2") + " L";
        
        // 判断是否为有效的加油记录（只保留加油记录）
        public bool IsFuelRecord => BusinessType == "加油";
    }
}