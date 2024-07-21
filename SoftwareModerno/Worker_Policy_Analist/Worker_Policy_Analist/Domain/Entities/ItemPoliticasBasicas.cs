using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Worker_Policy_Analist.Domain.Entities
{
    public class ItemPoliticasBasicas
    {
        public string NoSolicitud { get; set; } = string.Empty;
        public string TipoId { get; set; } = string.Empty;
        public string NumId { get; set; } = string.Empty;
        public string CodRolP { get; set; } = string.Empty;
        public string CodTipoS { get; set; } = string.Empty;
        public string CodCategoria { get; set; } = string.Empty;
    }
}
