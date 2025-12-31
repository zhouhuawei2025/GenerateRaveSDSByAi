using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateRaveSDSByAi.dataModel
{
    internal class DataDictionary
    {
        public string OID { get; set; }

        public string Name { get; set; }       

        public List<DataDictionaryEntry> DataDictionaryEntries { get; set; } = new List<DataDictionaryEntry>();
    }
}
