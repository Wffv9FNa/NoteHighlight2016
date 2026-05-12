using System;
using System.Text.RegularExpressions;

namespace Helper
{
    public static class StringExtension
    {
        public static string TemplateSubstitute(this string input, object data)
        {
            var type = data.GetType();
            return Regex.Replace(input, @"\{(\w+)\}", m =>
            {
                var name = m.Groups[1].Value;
                var prop = type.GetProperty(name);
                if (prop != null)
                {
                    return prop.GetValue(data, null).ToString();
                }
                else
                {
                    return m.Value;
                }
            });
        }
    }
}
