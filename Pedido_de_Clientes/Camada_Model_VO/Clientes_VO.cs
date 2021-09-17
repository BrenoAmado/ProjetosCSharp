using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model_VO
{
    public class Clientes_VO
    {
        private int id;
        private string nome;
        private string descricao;
        private int is_active;

        public Clientes_VO()
        {
        }

        public Clientes_VO(int intID, string strNome, string strDescricao, int intIs_Active)
        {
            ID = intID;
            Nome = strNome;
            Descricao = strDescricao;
            Is_Active = intIs_Active;
        }

        public int ID
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public string Nome
        {
            get { return this.nome; }
            set { this.nome = value; }
        }

        public string Descricao
        {
            get { return this.descricao; }
            set { this.descricao = value; }
        }

        public int Is_Active
        {
            get { return this.is_active; }
            set { this.is_active = value; }
        }
    }
}
