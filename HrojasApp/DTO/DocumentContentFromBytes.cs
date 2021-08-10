using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HrojasApp.DTO
{
    public class DocumentContentFromBytes
    {
        public string DocumentId { get; set; }

        public string FileName { get; set; }

        public Func<byte[]> ContentAccessor { get; set; }

        public bool IsEditable { get; set; }

        public string Base64 { get; set; }
        
    }
}
