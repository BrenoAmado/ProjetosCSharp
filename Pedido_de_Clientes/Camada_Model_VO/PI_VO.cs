using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model_VO
{
    public class PI_VO
    {
        private int id;
        private Clientes_VO clinte_id;
        private string descricao;
        private int estado;

        public PI_VO()
        {
        }

        public PI_VO(int intID, Clientes_VO ClClientes_VO, string strDescricao, int intEstado)
        {
            ID = intID;
            Clientes_VO = ClClientes_VO;
            Descricao = strDescricao;
            Estado = intEstado;
        }

        public int ID
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public Clientes_VO Clientes_VO
        {
            get { return this.clinte_id; }
            set { this.clinte_id = value; }
        }

        public string Descricao
        {
            get { return this.descricao; }
            set { this.descricao = value; }
        }

        public int Estado
        {
            get { return this.estado; }
            set { this.estado = value; }
        }
    }
}
