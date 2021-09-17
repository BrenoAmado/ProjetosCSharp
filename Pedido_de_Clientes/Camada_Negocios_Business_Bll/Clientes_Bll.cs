using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Camada_Fachada_Facade_FD;
using Camada_Model_VO;

namespace Camada_Negocios_Business_Bll
{
    public class Clientes_Bll
    {
        StreamReader objLeitor;
        string strLinhaLida;

        Clientes_FD objClientes_FD;

        public List<string> ImportarTxt()
        {
            try
            {
                List<string> Resultado = new List<string>();

                objLeitor = new StreamReader(@"C:\CursoProgramar\Clientes.txt");
                strLinhaLida = objLeitor.ReadLine();

                while (strLinhaLida != null)
                {
                    Resultado.Add(strLinhaLida);
                    strLinhaLida = objLeitor.ReadLine();
                }

                return Resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Erro ao Importar Texto ==>" +ex.Message);
            }
            finally
            {
                objLeitor.Close();
            }
        }

        public List<string> ImportarBdConectado()
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ImportarBdConectado();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public List<string> ImportarBdDesconectado()
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ImportarBdDesconectado();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarBd(Clientes_VO objClientes_VO)
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ConsultarBd(objClientes_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool InserirBd(Clientes_VO objClientes_VO)
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.InserirBd(objClientes_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool ExcluirBd(Clientes_VO objClientes_VO)
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ExcluirBd(objClientes_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public bool AlterarBd(Clientes_VO objClientes_VO)
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.AlterarBd(objClientes_VO);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarPIePEClientes(string strIDClientesSelecionados)
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ConsultarPIePEClientes(strIDClientesSelecionados);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarQuantidadePIClientes(string strIDClientesSelecionados)
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ConsultarQuantidadePIClientes(strIDClientesSelecionados);

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarClientesSemPE()
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ConsultarClientesSemPE();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public DataTable ConsultarQuantidadePedidos()
        {
            try
            {
                objClientes_FD = new Clientes_FD();
                return objClientes_FD.ConsultarQuantidadePedidos();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
