using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Camada_Dados_DAL_DAO
{
    public abstract class DAL_DAO : DB_DAO
    {
        public abstract DataTable ConsultarBd(Object objVO_VO);

        public abstract bool InserirBd(Object objVO_VO);

        public abstract bool ExcluirBd(Object objVO_VO);

        public abstract bool AlterarBd(Object objVO_VO);
    }
}
