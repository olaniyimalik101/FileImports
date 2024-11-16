using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileColumnReader
{
    public class CreateResponse
    {
        public int rowCount { get; set; }
        public int successCount { get; set; }
        public int failureCount { get; set; }
        public List<string> failedRecordsDetails { get; set; }
    }
}

