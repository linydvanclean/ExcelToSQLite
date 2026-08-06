using System;

namespace ExcelToSQLite.Models
{
    public class AttendanceRecord
    {
        public int Id { get; set; }
        public string? EmployeeId { get; set; }
        public string? EmployeeName { get; set; }
        public string? Department { get; set; }
        public DateTime CheckTime { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public int DayOfMonth { get; set; }
        
        // 辅助属性
        public string CheckTimeDisplay => CheckTime.ToString("yyyy-MM-dd HH:mm:ss");
        public string DateDisplay => CheckTime.ToString("yyyy-MM-dd");
        public string TimeDisplay => CheckTime.ToString("HH:mm:ss");
    }
}