using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model_VO
{
    public class Preferencias_VO
    {
        private int id;
        private string descricao;

        public Preferencias_VO()
        {
        }

        public Preferencias_VO(int intID, string strDescricao)
        {
            SetID(intID);
            SetDescricao(strDescricao);
        }

        public Preferencias_VO(string strDescricao)
        {
            SetDescricao(strDescricao);
        }

        public int GetID()
        {
            return this.id;
        }

        public void SetID(int intID)
        {
            this.id = intID;
        }

        public int ID
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public string GetDescricao()
        {
            return this.descricao;
        }

        public void SetDescricao(string strDescricao)
        {
            this.descricao = strDescricao;
        }

        public string Descricao
        {
            get { return this.descricao; }
            set { this.descricao = value; }
        }

        public List<Preferencias_VO> Preferencias_VOCollection = new List<Preferencias_VO>();
    }
}
