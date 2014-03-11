using System;
using System.IO;

namespace ReportHelper
{
    public static class ImageExtensions
    {
        public static Byte[] FromBase64StringToByte(this string s)
        {
            try
            {
                MemoryStream stream = new MemoryStream(Convert.FromBase64String(s));

                return stream.ToArray();
            }
            catch
            {
                return null;
            }
        }
    }
}   