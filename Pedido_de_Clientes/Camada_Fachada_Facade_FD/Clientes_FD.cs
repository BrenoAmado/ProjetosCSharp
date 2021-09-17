using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_Dados_DAL_DAO;
using Camada_Model_VO;

namespace Camada_Fachada_Facade_FD
{
    public class Clientes_FD
    {
        Clientes_DAO objClientes_DAO;

        public List<string> ImportarBdConectado()
        {
            try
            {
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ImportarBdConectado();

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ImportarBdDesconectado();

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ConsultarBd(objClientes_VO);

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.InserirBd(objClientes_VO);

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ExcluirBd(objClientes_VO);

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.AlterarBd(objClientes_VO);

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ConsultarPIePEClientes(strIDClientesSelecionados);

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ConsultarQuantidadePIClientes(strIDClientesSelecionados);

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ConsultarClientesSemPE();

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
                objClientes_DAO = new Clientes_DAO();
                return objClientes_DAO.ConsultarQuantidadePedidos();

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
    }
}
