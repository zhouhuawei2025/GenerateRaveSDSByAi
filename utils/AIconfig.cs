using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateRaveSDSByAi.utils
{
    public class AIconfig
    {
        public string Url { get; set; } = "http://192.168.8.58:30008/v1";
        public string Model { get; set; } = "openai/gpt-oss-120b";
        public string? ApiKey { get; set; }
        public int batchSize { get; set; }
    }
}
