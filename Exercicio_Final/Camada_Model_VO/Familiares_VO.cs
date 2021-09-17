using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model_VO
{
    public class Familiares_VO
    {
        private int cod;
        private string nome;
        private string sexo;
        private int idade;
        private double ganho_mensal_total;
        private double gasto_mensal_total;
        private string observacao;

        public Familiares_VO()
        {
        }

        public Familiares_VO(int intCOD, string strNome, string strSexo)
        {
            SetCOD(intCOD);
            SetNome(strNome);
            SetSexo(strSexo);
        }

        public Familiares_VO(int intCOD, string strNome, string strSexo, int intIdade, double dbGanho, double dbGasto, string strObservacao)
        {
            SetCOD(intCOD);
            SetNome(strNome);
            SetSexo(strSexo);
            SetIdade(intIdade);
            SetGanho(dbGanho);
            SetGasto(dbGasto);
            SetObservacao(strObservacao);
        }

        public int GetCOD()
        {
            return this.cod;
        }

        public void SetCOD(int intCOD)
        {
            this.cod = intCOD;
        }

        public int COD
        {
            get { return this.cod; }
            set { this.cod = value; }
        }

        public string GetNome()
        {
            return this.nome;
        }

        public void SetNome(string strNome)
        {
            this.nome = strNome;
        }

        public string Nome
        {
            get { return this.nome; }
            set { this.nome = value; }
        }

        public string GetSexo()
        {
            return this.sexo;
        }

        public void SetSexo(string strSexo)
        {
            if (strSexo.ToUpper() == "MASCULINO" || strSexo.ToUpper() == "FEMININO" || strSexo.ToUpper() == "OUTRO")
            {
                this.sexo = strSexo.ToUpper();
            }
        }

        public string Sexo
        {
            get { return this.sexo; }
            set
            {
                if (value.ToUpper() == "MASCULINO" || value.ToUpper() == "FEMININO" || value.ToUpper() == "OUTRO")
                {
                    this.sexo = value.ToUpper();
                }
                else
                {
                    throw new Exception("Escolha uma opção válida de Sexo");
                }
            }
        }

        public int GetIdade()
        {
            return this.idade;
        }

        public void SetIdade(int intIdade)
        {
            this.idade = intIdade;
        }

        public int Idade
        {
            get { return this.idade; }
            set { this.idade = value; }
        }

        public double GetGanho()
        {
            return this.ganho_mensal_total;
        }

        public void SetGanho(double dbGanho)
        {
            this.ganho_mensal_total = dbGanho;
        }

        public double Ganho
        {
            get { return this.ganho_mensal_total; }
            set { this.ganho_mensal_total = value; }
        }

        public double GetGasto()
        {
            return this.gasto_mensal_total;
        }

        public void SetGasto(double dbGasto)
        {
            this.gasto_mensal_total = dbGasto;
        }

        public double Gasto
        {
            get { return this.gasto_mensal_total; }
            set { this.gasto_mensal_total = value; }
        }

        public string GetObservacao()
        {
            return this.observacao;
        }

        public void SetObservacao(string strObservacao)
        {
            this.observacao = strObservacao;
        }

        public string Observacao
        {
            get { return this.observacao; }
            set { this.observacao = value; }
        }

        public List<Familiares_VO> Familiares_VOCollection = new List<Familiares_VO>();
    }
}
