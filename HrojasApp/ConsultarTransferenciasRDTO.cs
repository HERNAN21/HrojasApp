using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace HrojasApp
{
    public class ConsultarTransferenciasRDTO
    {
        [Display(Order = 0)]
        public string codEmpresa { get; set; }

        [Display(Order = 1)]
        public string nomEmpresa { get; set; }

        [Display(Order = 2)]
        public string codCartera { get; set; }

        [Display(Order = 3)]
        public string nomCartera { get; set; }

        [Display(Order = 4)]
        public decimal nroActa { get; set; }

        [Display(Order = 5)]
        public decimal nroOperacion { get; set; }

        [Display(Order = 6)]
        public string codAccionista { get; set; }

        [Display(Order = 7)]
        public string desAccionista { get; set; }

        [Display(Order = 8)]
        public string txtOperacion { get; set; }

        [Display(Order = 9)]
        public decimal compra { get; set; }

        [Display(Order = 10)]
        public decimal venta { get; set; }

        [Display(Order = 11)]
        public string codAgenteCompra { get; set; }

        [Display(Order = 12)]
        public string agenteCompra { get; set; }

        [Display(Order = 13)]
        public string codAgenteVende { get; set; }

        [Display(Order = 14)]
        public string agenteVende { get; set; }

        [Display(Order = 15)]
        public string fecOperacion { get; set; }

        [Display(Order = 16)]
        public string codNaturaleza { get; set; }

        [Display(Order = 17)]
        public string desNaturaleza { get; set; }

        [Display(Order = 18)]
        public decimal nroSecuencia { get; set; }

        [Display(Order = 19)]
        public string codCtaAccionista { get; set; }

        [Display(Order = 20)]
        public string nroCertificado { get; set; }

        [Display(Order = 21)]
        public string codSerieCertificado { get; set; }

        [Display(Order = 22)]
        public string idCartera { get; set; }
    }
}
