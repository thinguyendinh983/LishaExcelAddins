using ExcelDna.Integration;
using System.Globalization;

namespace Lisha.ExcelAddins.Functions
{
    public static class LishaFunctions
    {
        [ExcelFunction(Description = "Giới thiệu về chương trình")]
        public static string LishaAbout()
        {
            return "Lisha Excel Add-Ins v1.0. (C) " + DateTime.Now.Year + " Nguyen Dinh Thi - Email: thinguyendinh983@gmail.com";
        }

        /// <summary>
        /// Ham chuyen phan nguyen cua so thanh chu
        /// </summary>
        /// <param name="Number"></param>
        /// <param name="suffix"></param>
        /// <returns></returns>
        [ExcelFunction(Description = "Hàm chuyển phần nguyên của số thành chữ")]
        public static string LishaSoSangChu(
            [ExcelArgument(Name ="Number", Description ="Số cần chuyển thành chữ")]
            double Number,
            [ExcelArgument(Name ="suffix", Description ="Hiển thị đơn vị tiền tệ (1 - Có, 0 - Không)")]
            bool suffix = true)
        {
            string[] unitNumbers = new string[] { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
            string[] placeValues = new string[] { "", "nghìn", "triệu", "tỷ" };
            bool isNegative = false;

            // -12345678.3445435 => "-12345678"
            string sNumber = Number.ToString("#");
            double number = Convert.ToDouble(sNumber);
            if (number < 0)
            {
                number = -number;
                sNumber = number.ToString();
                isNegative = true;
            }


            int ones, tens, hundreds;

            int positionDigit = sNumber.Length;   // last -> first

            string result = " ";


            if (positionDigit == 0)
                result = unitNumbers[0] + result;
            else
            {
                // 0:       ###
                // 1: nghìn ###,###
                // 2: triệu ###,###,###
                // 3: tỷ    ###,###,###,###
                int placeValue = 0;

                while (positionDigit > 0)
                {
                    // Check last 3 digits remain ### (hundreds tens ones)
                    tens = hundreds = -1;
                    ones = Convert.ToInt32(sNumber.Substring(positionDigit - 1, 1));
                    positionDigit--;
                    if (positionDigit > 0)
                    {
                        tens = Convert.ToInt32(sNumber.Substring(positionDigit - 1, 1));
                        positionDigit--;
                        if (positionDigit > 0)
                        {
                            hundreds = Convert.ToInt32(sNumber.Substring(positionDigit - 1, 1));
                            positionDigit--;
                        }
                    }

                    if ((ones > 0) || (tens > 0) || (hundreds > 0) || (placeValue == 3))
                        result = placeValues[placeValue] + result;

                    placeValue++;
                    if (placeValue > 3) placeValue = 1;

                    if ((ones == 1) && (tens > 1))
                        result = "một " + result;
                    else
                    {
                        if ((ones == 5) && (tens > 0))
                            result = "lăm " + result;
                        else if (ones > 0)
                            result = unitNumbers[ones] + " " + result;
                    }
                    if (tens < 0)
                        break;
                    else
                    {
                        if ((tens == 0) && (ones > 0)) result = "lẻ " + result;
                        if (tens == 1) result = "mười " + result;
                        if (tens > 1) result = unitNumbers[tens] + " mươi " + result;
                    }
                    if (hundreds < 0) break;
                    else
                    {
                        if ((hundreds > 0) || (tens > 0) || (ones > 0))
                            result = unitNumbers[hundreds] + " trăm " + result;
                    }
                    result = " " + result;
                }
            }
            result = result.Trim();
            if (isNegative) result = "Âm " + result;
            return char.ToUpper(result[0]) + result.Substring(1) + (suffix ? " đồng" : "");
        }

        /// <summary>
        /// Ham dinh dang so tien VND
        /// </summary>
        /// <param name="Number"></param>
        /// <param name="suffix"></param>
        /// <returns></returns>
        [ExcelFunction(Description = "Định dạng số tiền VND")]
        public static string LishaDinhDangSoTienVND(
            [ExcelArgument(Name ="Number", Description ="Số cần chuyển sang định dạng VNĐ")]
            double Number,
            [ExcelArgument(Name ="suffix", Description ="Hiển thị đơn vị tiền tệ (1 - Có, 0 - Không)")]
            bool suffix = true)
        {
            return String.Format(new CultureInfo("vi-VN"), "{0:#,##0}" + (suffix ? " VNĐ" : ""), Number);
        }

        ///// <summary>
        ///// Ham xoa cac khoang trang thua
        ///// </summary>
        ///// <param name="Text"></param>
        ///// <returns></returns>
        //[ExcelFunction(Description = "Hàm xóa các khoảng trắng thừa trong chuỗi ký tự")]
        //public static string LishaRemoveExtraSpaces(
        //    [ExcelArgument(Name ="Text", Description ="Chuỗi cần xóa khoảng trắng thừa")]
        //    string Text)
        //{
        //    return Regex.Replace(Text.Trim(), @"\s+", " ");
        //}
    }
}
