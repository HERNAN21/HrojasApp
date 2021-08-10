using System;

namespace Alfa.Valores.DTO.ReportModels
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
