using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ActiveConverter
{
    static class Extensions
    {
        /// <summary>
        /// http://programujte.com/clanek/2006062803-jak-odstranit-diakritiku-z-retezce/
        /// </summary>
        /// <param name="this"></param>
        /// <returns></returns>
        public static string RemoveDiacritics(this string @this)
        {
            // oddělení znaků od modifikátorů (háčků, čárek, atd.)
            @this = @this.Normalize(System.Text.NormalizationForm.FormD);
            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            for (int i = 0; i < @this.Length; i++)
            {
                // do řetězce přidá všechny znaky kromě modifikátorů
                if (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(@this[i]) != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(@this[i]);
                }
            }

            // vrátí řetězec bez diakritiky
            return sb.ToString();
        }
    }
}
