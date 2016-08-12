using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GenerateCV
{
    public static class Extention
    {
        static List<string> symbol;
        static Extention()
        {
            symbol = CultureInfo
           .GetCultures(CultureTypes.AllCultures)
           .Where(c => !c.IsNeutralCulture)
           .Select(culture =>
           {
               try
               {
                   return new RegionInfo(culture.LCID);
               }
               catch
               {
                   return null;
               }
           }).Where(ri => ri != null)
          .Select(ri => ri.CurrencySymbol).Distinct().ToList();
            symbol.AddRange(new List<string> { "/", "+", "-", "*", "!", "?" });
        }

        public static string Value(this string val)
        {            
            if (string.IsNullOrEmpty(val)) return string.Empty;
            var res = new StringBuilder();
            foreach (var cara in val)
            {
                if ((char.IsSymbol(cara) && !symbol.Contains(cara.ToString())) || char.IsControl(cara))
                    continue;
                res.Append(cara);
            }
            return res.ToString();
        }
    }
}
