using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HDRPlanningAssistant.Helpers
{
    public static class StringHelper
    {
        public static string cropLine(string line, string cropChar) 
        { 
            return line.Substring(line.IndexOf(cropChar) + 1, line.Length - line.IndexOf(cropChar) - 1); 
        }
    }
}
