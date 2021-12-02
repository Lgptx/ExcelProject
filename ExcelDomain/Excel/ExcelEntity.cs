using System;

namespace ExcelDomain
{
    public class ExcelEntity
    {
        public int Id { get; set; }
        public DateTime DataEntrega { get; set; }
        public string NomeDoProduto { get; set; }
        public double Quantidade { get; set; }
        public double ValorUnitario { get; set; }
        public double? ValorTotal { get; set; }
    }
}
