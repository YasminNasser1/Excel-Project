namespace Excel_Project.Models
{
    public class UserInfo
    {
        public string FullName { get; set; }          // الاسم
        public DateTime Date { get; set; }            // التاريخ
        public string Address { get; set; }           // العنوان
        public string Governorate { get; set; }      // المحافظه
        public string MobileNumber { get; set; }      // رقم موبايل
        public string AdditionalNumber { get; set; }  // رقم اضافي
        public decimal Price { get; set; }           // السعر
        public string ProductCode { get; set; }       // الكود
        public string ProductName { get; set; }       // اسم المنتج
        public int Quantity { get; set; }            // عدد القطع
    }
}