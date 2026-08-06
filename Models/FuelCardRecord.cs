using System;

namespace ExcelToSQLite.Models
{
    /// <summary>
    /// 加油卡交易记录 - 中国石化格式
    /// </summary>
    public class FuelCardRecord
    {
        public int Id { get; set; }
        
        /// <summary>
        /// 卡号
        /// </summary>
        public string? CardNumber { get; set; }
        
        /// <summary>
        /// 交易时间
        /// </summary>
        public DateTime TransactionTime { get; set; }
        
        /// <summary>
        /// 业务类型：圈存、加油
        /// </summary>
        public string? BusinessType { get; set; }
        
        /// <summary>
        /// 油品类型（如：92号车用汽油）
        /// </summary>
        public string? FuelType { get; set; }
        
        /// <summary>
        /// 数量（升）
        /// </summary>
        public decimal Quantity { get; set; }
        
        /// <summary>
        /// 单价（元/升）
        /// </summary>
        public decimal UnitPrice { get; set; }
        
        /// <summary>
        /// 金额（元）- 从分值转换
        /// </summary>
        public decimal Amount { get; set; }
        
        /// <summary>
        /// 奖励分值
        /// </summary>
        public decimal BonusPoints { get; set; }
        
        /// <summary>
        /// 优惠价（元/升）
        /// </summary>
        public decimal DiscountPrice { get; set; }
        
        /// <summary>
        /// 余额（元）
        /// </summary>
        public decimal Balance { get; set; }
        
        /// <summary>
        /// 交易地点（加油站名称）
        /// </summary>
        public string? Location { get; set; }
        
        /// <summary>
        /// 操作员
        /// </summary>
        public string? Operator { get; set; }
        
        /// <summary>
        /// 备注
        /// </summary>
        public string? Remarks { get; set; }
        
        /// <summary>
        /// 客户名称（从表头提取）
        /// </summary>
        public string? CustomerName { get; set; }
        
        /// <summary>
        /// 网点名称（从表头提取）
        /// </summary>
        public string? NetworkName { get; set; }
        
        /// <summary>
        /// 导入时间
        /// </summary>
        public DateTime CreatedAt { get; set; }
        
        // 辅助显示属性
        public string TransactionTimeDisplay => TransactionTime.ToString("yyyy-MM-dd HH:mm:ss");
        public string AmountDisplay => Amount.ToString("F2") + " 元";
        public string QuantityDisplay => Quantity.ToString("F2") + " L";
        
        /// <summary>
        /// 判断是否为有效的交易记录（圈存或加油）
        /// </summary>
        public bool IsValidRecord => BusinessType == "圈存" || BusinessType == "加油";
    }
}