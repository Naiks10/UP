﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Autodb
{
    class crypt
    {
        public static string getHash(string text)
        {
            byte[] data = new UTF8Encoding().GetBytes(text);
            SHA256 shaM = new SHA256Managed();
            return BitConverter.ToString(shaM.ComputeHash(data)).Replace("-", "").ToLower();
        }
    }
}
