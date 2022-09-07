using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIExample.Models
{
    public static  class StreamHelpers
    {
        public static byte[] GetBytes(MemoryStream stream)
        {
            
            var bytes = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            stream.ReadAsync(bytes, 0, bytes.Length);
            stream.Dispose();
            stream.Position = 0;


            return bytes;
        }
    }
}
