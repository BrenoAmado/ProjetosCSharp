using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Camada_Model_VO
{
    public class PrefFam_VO
    {
        private Familiares_VO familiares_VO;
        private Preferencias_VO preferencias_VO;
        private double intensidade;
        private string observacao;

        public PrefFam_VO()
        {
        }

        public PrefFam_VO(Familiares_VO famFamiliares_VO, Preferencias_VO prefPreferencias_VO, double dbIntensidade, string strObservacao)
        {
            SetFamiliares_VO(famFamiliares_VO);
            SetPreferencias_VO(prefPreferencias_VO);
            SetIntensidade(dbIntensidade);
            SetObservacao(strObservacao);
        }

        public Familiares_VO GetFamiliares_VO()
        {
            return this.familiares_VO;
        }

        public void SetFamiliares_VO(Familiares_VO famFamiliares_VO)
        {
            this.familiares_VO = famFamiliares_VO;
        }

        public Familiares_VO Familiares_VO
        {
            get { return this.familiares_VO; }
            set { this.familiares_VO = value; }
        }

        public Preferencias_VO GetPreferencias_VO()
        {
            return this.preferencias_VO;
        }

        public void SetPreferencias_VO(Preferencias_VO Preferencias_VO)
        {
            this.preferencias_VO = Preferencias_VO;
        }

        public Preferencias_VO Preferencias_VO
        {
            get { return this.preferencias_VO; }
            set { this.preferencias_VO = value; }
        }

        public double GetIntensidade()
        {
            return this.intensidade;
        }

        public void SetIntensidade(double dbIntensidade)
        {
            this.intensidade = dbIntensidade;
        }

        public double Intensidade
        {
            get { return this.intensidade; }
            set { this.intensidade = value; }
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

        public List<PrefFam_VO> PrefFam_VOCollection = new List<PrefFam_VO>();
    }
}
