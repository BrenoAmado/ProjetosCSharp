using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Camada_Negocios_Business_Bll;
using Camada_Model_VO;
using Excel = Microsoft.Office.Interop.Excel;
using Email = Microsoft.Office.Interop.Outlook;
using Microsoft.VisualBasic;

namespace Pedido_de_Clientes
{
    public partial class frmPedidos_de_Clientes : Form
    {
        int intIDSalvoClientes;
        string strNomeSalvoClientes;

        int intIDSalvoPI, intCliente_IDSalvoPI;
        string strNomeSalvoPI;

        int intIDSalvoPE, intCliente_IDSalvoPE;
        string strNomeSalvoPE;

        bool bolAddClientes;
        bool bolAddPI;
        bool bolAddPE;

        Clientes_Bll objClientes_Bll;
        Clientes_VO objClientes_VO;

        PI_Bll objPI_Bll;
        PI_VO objPI_VO;

        PE_Bll objPE_Bll;
        PE_VO objPE_VO;

        Excel.Application objExcelApp;
        Excel.Workbook objExcelWb;
        Excel.Worksheet objExcelWs;
        Excel.Range objExcelCabecalho, objExcelDados;

        Email.Application objEmailApp;
        Email.MailItem objEmailMensagem;
        Email.OlAttachmentType objAnexoTipo;
        string[] objAnexoArquivo = new String[0];
        string objDisplayName;
        long objAnexoPosicao;

        string strIdClientesSelecionados;

        public frmPedidos_de_Clientes()
        {
            InitializeComponent();
        }

        #region Clientes
        private void btnSwitchCase_Click(object sender, EventArgs e)
        {
            switch(MessageBox.Show("Escolha Sim, Não ou Cancelar!", "Estrutura de Escolha", MessageBoxButtons.YesNoCancel))
            {
                case DialogResult.Yes:
                    MessageBox.Show("Você Escolheu Sim!");
                    break;
                case DialogResult.No:
                    MessageBox.Show("Você Escolheu Não!");
                    break;
                case DialogResult.Cancel:
                    MessageBox.Show("Você Escolheu Cancelar!");
                    break;
                default:
                    MessageBox.Show("Escolha Errada! Escolha Sim, Não ou Cancelar!!!");
                    break;
            }
        }

        private void btnImportarTxt_Click(object sender, EventArgs e)
        {
            lstbxClientes.Items.Clear();
            objClientes_Bll = new Clientes_Bll();
            lstbxClientes.Items.AddRange(objClientes_Bll.ImportarTxt().ToArray());
        }

        private void btnImportarBdConectado_Click(object sender, EventArgs e)
        {
            lstbxClientes.Items.Clear();
            objClientes_Bll = new Clientes_Bll();
            lstbxClientes.Items.AddRange(objClientes_Bll.ImportarBdConectado().ToArray());
        }

        private void btnImportarBdDesconectado_Click(object sender, EventArgs e)
        {
            lstbxClientes.Items.Clear();
            objClientes_Bll = new Clientes_Bll();
            lstbxClientes.Items.AddRange(objClientes_Bll.ImportarBdDesconectado().ToArray());
        }

        private void btnConsultarBd_Click(object sender, EventArgs e)
        {
            ConsultarBdClientes();
        }

        public void ConsultarBdClientes(int? intID = null, string strNome = null, string strDescricao = null, int? intIs_Active = null)
        {
            try
            {
                objClientes_Bll = new Clientes_Bll();
                objClientes_VO = new Clientes_VO();
                objClientes_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);
                objClientes_VO.Nome = strNome;
                objClientes_VO.Descricao = strDescricao;
                objClientes_VO.Is_Active = Convert.ToInt32(intIs_Active == null ? 0 : intIs_Active);

                bndSrcClientes.DataSource = objClientes_Bll.ConsultarBd(objClientes_VO);
                dtgdvwClientes.DataSource = bndSrcClientes;

                cmbBxPI.DataSource = null;
                cmbBxPI.Items.Clear();
                cmbBxPI.DisplayMember = "Nome";
                cmbBxPI.ValueMember = "ID";

                cmbBxPI.DataSource = bndSrcClientes.DataSource;
                cmbBxPI.SelectedIndex = Convert.ToInt32(intID > 0 ? intID -1 : 0);

                cmbBxPE.DataSource = null;
                cmbBxPE.Items.Clear();
                cmbBxPE.DisplayMember = "Nome";
                cmbBxPE.ValueMember = "ID";

                cmbBxPE.DataSource = bndSrcClientes.DataSource;
                cmbBxPE.SelectedIndex = Convert.ToInt32(intID > 0 ? intID - 1 : 0);
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        private void btnInserirBd_Click(object sender, EventArgs e)
        {
            InserirBdClientes(dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                              dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                              Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Is_Active"].EditedFormattedValue.ToString()));
            ConsultarBdClientes();
        }

        public void InserirBdClientes(string strNome, string strDescricao, int intIs_Active)
        {
            try
            {
                objClientes_Bll = new Clientes_Bll();
                objClientes_VO = new Clientes_VO();

                objClientes_VO.Nome = strNome;
                objClientes_VO.Descricao = strDescricao;
                objClientes_VO.Is_Active = intIs_Active <= 0 ? 0 : intIs_Active;

                if (objClientes_Bll.InserirBd(objClientes_VO))
                {
                    MessageBox.Show("Inserir OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Inclusão!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        private void btnExcluirBd_Click(object sender, EventArgs e)
        {
            ExcluirBdClientes(intIDSalvoClientes);
            ConsultarBdClientes();
        }

        public void ExcluirBdClientes(int intID)
        {
            try
            {
                objClientes_Bll = new Clientes_Bll();
                objClientes_VO = new Clientes_VO();

                objClientes_VO.ID = intID <= 0 ? 0 : intID;

                if (objClientes_Bll.ExcluirBd(objClientes_VO))
                {
                    MessageBox.Show("Excluir OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Exclusão!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        private void btnAlterarBd_Click(object sender, EventArgs e)
        {
            AlterarBdClientes(intIDSalvoClientes, dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                              dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                              Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Is_Active"].EditedFormattedValue.ToString()));
            ConsultarBdClientes();
        }

        public void AlterarBdClientes(int intID, string strNome, string strDescricao, int intIs_Active)
        {
            try
            {
                objClientes_Bll = new Clientes_Bll();
                objClientes_VO = new Clientes_VO();
                objClientes_VO.ID = intID <= 0 ? 0 : intID;
                objClientes_VO.Nome = strNome;
                objClientes_VO.Descricao = strDescricao;
                objClientes_VO.Is_Active = intIs_Active <= 0 ? 0 : intIs_Active;

                if (objClientes_Bll.InserirBd(objClientes_VO))
                {
                    MessageBox.Show("Alterar OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Alteração!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }


        private void dtgdvwClientes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwClientes.CurrentRow.Cells["ID"].Value.ToString()))
            {
                intIDSalvoClientes = Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["ID"].Value);
            }
            strNomeSalvoClientes = dtgdvwClientes.CurrentRow.Cells["Nome"].Value.ToString();
        }

        private void bndNavBtnPesquisarClientes_Click(object sender, EventArgs e)
        {
            ConsultarBdClientes(null, bndNavTxtClientes.Text);
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            bolAddClientes = true;
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir o Cliente '" + strNomeSalvoClientes + "' na Tabela Clientes? Sim ou Não?", "Exclusão de Clientes", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBdClientes(intIDSalvoClientes);
            }
            ConsultarBdClientes();
        }

        private void bndNavBtnConfirmarClientes_Click(object sender, EventArgs e)
        {
            if (bolAddClientes)
            {
                if (MessageBox.Show("Deseja Inserir o Cliente '" + dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString() + "' na Tabela Clientes? Sim ou Não?", "Inserção de Clientes", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    InserirBdClientes(dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                      dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                      Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Is_Active"].EditedFormattedValue.ToString()));
                    ConsultarBdClientes();
                }
                bolAddClientes = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar o Cliente '" + strNomeSalvoClientes + "' para novo Cliente: '" + dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString() + "' na Tabela Clientes? Sim ou Não?", "Alteração de Clientes", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AlterarBdClientes(intIDSalvoClientes, dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                              dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                              Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Is_Active"].EditedFormattedValue.ToString()));
                }
                else if (MessageBox.Show("Deseja Alterar a Descricao Atual para nova Descricao: '" + dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "' na Tabela Clientes? Sim ou Não?", "Alteração de Descricao", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AlterarBdClientes(intIDSalvoClientes, dtgdvwClientes.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                      dtgdvwClientes.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                      Convert.ToInt32(dtgdvwClientes.CurrentRow.Cells["Is_Active"].EditedFormattedValue.ToString()));
                }
            }
            ConsultarBdClientes();
        }
        #endregion

        #region Pedidos Interior
        public void ConsultarBdPI(int? intID = null, int? intCliente_ID = null, string strNome = null, string strDescricao = null, int? intEstado = null)
        {
            objPI_Bll = new PI_Bll();
            objPI_VO = new PI_VO();
            objPI_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);
            objPI_VO.Clientes_VO = new Clientes_VO();
            objPI_VO.Clientes_VO.ID = Convert.ToInt32(intCliente_ID == null ? 0 : intCliente_ID);
            objPI_VO.Clientes_VO.Nome = strNome;

            bndSrcPI.DataSource = objPI_Bll.ConsultarBd(objPI_VO);

            dtgdvwPI.DataSource = null;
            dtgdvwPI.Columns.Clear();
            dtgdvwPI.AllowUserToAddRows = false;

            dtgdvwPI.Columns.Add("ID", "ID do Pedido");
            dtgdvwPI.Columns["ID"].DataPropertyName = "ID";

            DataGridViewComboBoxColumn objDtgdvwCmbBxClClientes = new DataGridViewComboBoxColumn();
            objDtgdvwCmbBxClClientes.DataSource = bndSrcClientes.DataSource;
            objDtgdvwCmbBxClClientes.Name = "Cliente_ID";
            objDtgdvwCmbBxClClientes.ValueType = typeof(int);
            objDtgdvwCmbBxClClientes.DisplayMember = "Nome";
            objDtgdvwCmbBxClClientes.ValueMember = "ID";
            objDtgdvwCmbBxClClientes.HeaderText = "Nome do Cliente";

            dtgdvwPI.Columns.Add(objDtgdvwCmbBxClClientes);
            dtgdvwPI.Columns["Cliente_ID"].DataPropertyName = "Cliente_ID";

            dtgdvwPI.Columns.Add("Descricao", "Descricao do Pedido Interior");
            dtgdvwPI.Columns["Descricao"].DataPropertyName = "Descricao";

            dtgdvwPI.Columns.Add("Estado", "Estado da Entrega do Pedido");
            dtgdvwPI.Columns["Estado"].DataPropertyName = "Estado";

            dtgdvwPI.DataSource = bndSrcPI;
        }

        public void dtgdvwRefreshPI()
        {
            try
            {
                ConsultarBdPI(null, Convert.ToInt32(cmbBxPI.SelectedValue.ToString()), cmbBxPI.Text);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void cmbBxPI_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbBxPI.SelectedIndex > 0)
            {
                cmbBxPI.Text = cmbBxPI.Text.Trim();
            }
        }

        private void cmbBxPI_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbBxPI.Text))
            {
                strNomeSalvoPI = cmbBxPI.Text;
                dtgdvwRefreshPI();
            }
        }

        public void InserirBdPI(int intCliente_ID, string strNome, string strDescricao, int intEstado)
        {
            try
            {
                objPI_Bll = new PI_Bll();
                objPI_VO = new PI_VO();
                objPI_VO.Clientes_VO = new Clientes_VO();
                objPI_VO.Clientes_VO.ID = intCliente_ID <= 0 ? 0 : intCliente_ID;
                objPI_VO.Clientes_VO.Nome = strNome;
                objPI_VO.Descricao = strDescricao;
                objPI_VO.Estado = intEstado <= 0 ? 0 : intEstado;

                if (objPI_Bll.InserirBd(objPI_VO))
                {
                    MessageBox.Show("Inserir OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Inserção!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" +ex);
            }
        }

        public void ExcluirBdPI(int intID)
        {
            try
            {
                objPI_Bll = new PI_Bll();
                objPI_VO = new PI_VO();
                objPI_VO.ID = intID <= 0 ? 0 : intID;

                if (objPI_Bll.ExcluirBd(objPI_VO))
                {
                    MessageBox.Show("Excluir OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Exclusão!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" +ex);
            }
        }

        public void AlterarBdPI(int intID, int intCliente_ID, string strNome, string strDescricao, int intEstado)
        {
            try
            {
                objPI_Bll = new PI_Bll();
                objPI_VO = new PI_VO();
                objPI_VO.ID = intID <= 0 ? 0 : intID;
                objPI_VO.Clientes_VO = new Clientes_VO();
                objPI_VO.Clientes_VO.ID = intCliente_ID <= 0 ? 0 : intCliente_ID;
                objPI_VO.Clientes_VO.Nome = strNome;
                objPI_VO.Descricao = strDescricao;
                objPI_VO.Estado = intEstado <= 0 ? 0 : intEstado;

                if (objPI_Bll.AlterarBd(objPI_VO))
                {
                    MessageBox.Show("Alterar OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Alteração!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        private void dtgdvwPI_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwPI.CurrentRow.Cells["ID"].Value.ToString()) &&
                !string.IsNullOrEmpty(dtgdvwPI.CurrentRow.Cells["Cliente_ID"].Value.ToString()))
            {
                intIDSalvoPI = Convert.ToInt32(dtgdvwPI.CurrentRow.Cells["ID"].Value.ToString());
                intCliente_IDSalvoPI = Convert.ToInt32(dtgdvwPI.CurrentRow.Cells["Cliente_ID"].Value.ToString());
            }
            else
            {
                dtgdvwPI.CurrentRow.Cells["Cliente_ID"].Value = cmbBxPI.SelectedValue;
                intCliente_IDSalvoPI = Convert.ToInt32(dtgdvwPI.CurrentRow.Cells["Cliente_ID"].Value.ToString());
            }
            dtgdvwPI.CurrentRow.Cells["Cliente_ID"].Selected = false;
            dtgdvwPI.CurrentRow.Cells["Cliente_ID"].ReadOnly = true;
        }

        private void bndNavBtnPesquisarPI_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(bndNavTxtPI.Text))
            {
                ConsultarBdPI(Convert.ToInt32(bndNavTxtPI.Text));
            }
            else
            {
                dtgdvwRefreshPI();
            }
        }

        private void bindingNavigatorAddNewItem1_Click(object sender, EventArgs e)
        {
            bolAddPI = true;
        }

        private void bindingNavigatorDeleteItem1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir o Pedido Interior do Cliente '" + strNomeSalvoPI + "' da Tabela Pedidos Interior? Sim ou Não?", "Exclusão de Pedidos Interior", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBdPI(intIDSalvoPI);
            }
            dtgdvwRefreshPI();
        }

        private void bndNavBtnConfirmarPI_Click(object sender, EventArgs e)
        {
            try
            {
                if (bolAddPI)
                {
                    if (MessageBox.Show("Deseja Inserir o Pedido Interior do Cliente '"+ strNomeSalvoPI +"' na Tabela Pedidos Interior? Sim ou Não?", "Inserção de Pedidos Interior", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        InserirBdPI(intCliente_IDSalvoPI, strNomeSalvoPI, dtgdvwPI.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                    Convert.ToInt32(dtgdvwPI.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                    }
                    bolAddPI = false;
                }
                else
                {
                    if (MessageBox.Show("Deseja Alterar o Pedido Atual para o nov Pedido: '" + dtgdvwPI.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "' na Tabela Pedidos Interior? Sim ou Não?", "Alteração de Pedidos Interior", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        AlterarBdPI(intIDSalvoPI, intCliente_IDSalvoPI, strNomeSalvoPI, dtgdvwPI.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                    Convert.ToInt32(dtgdvwPI.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                    }
                }
                dtgdvwRefreshPI();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Pedidos Exterior
        public void ConsultarBdPE(int? intID = null, int? intCliente_ID = null, string strNome = null, string strDescricao = null, int? intEstado = null)
        {
            objPE_Bll = new PE_Bll();
            objPE_VO = new PE_VO();
            objPE_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);
            objPE_VO.Clientes_VO = new Clientes_VO();
            objPE_VO.Clientes_VO.ID = Convert.ToInt32(intCliente_ID == null ? 0 : intCliente_ID);
            objPE_VO.Clientes_VO.Nome = strNome;

            bndSrcPE.DataSource = objPE_Bll.ConsultarBd(objPE_VO);

            dtgdvwPE.DataSource = null;
            dtgdvwPE.Columns.Clear();
            dtgdvwPE.AllowUserToAddRows = false;

            dtgdvwPE.Columns.Add("ID", "ID do Pedido");
            dtgdvwPE.Columns["ID"].DataPropertyName = "ID";

            DataGridViewComboBoxColumn objDtgdvwCmbBxClClientes = new DataGridViewComboBoxColumn();
            objDtgdvwCmbBxClClientes.DataSource = bndSrcClientes.DataSource;
            objDtgdvwCmbBxClClientes.Name = "Cliente_ID";
            objDtgdvwCmbBxClClientes.ValueType = typeof(int);
            objDtgdvwCmbBxClClientes.DisplayMember = "Nome";
            objDtgdvwCmbBxClClientes.ValueMember = "ID";
            objDtgdvwCmbBxClClientes.HeaderText = "Nome do Cliente";

            dtgdvwPE.Columns.Add(objDtgdvwCmbBxClClientes);
            dtgdvwPE.Columns["Cliente_ID"].DataPropertyName = "Cliente_ID";

            dtgdvwPE.Columns.Add("Descricao", "Descricao do Pedido Interior");
            dtgdvwPE.Columns["Descricao"].DataPropertyName = "Descricao";

            dtgdvwPE.Columns.Add("Estado", "Estado da Entrega do Pedido");
            dtgdvwPE.Columns["Estado"].DataPropertyName = "Estado";

            dtgdvwPE.DataSource = bndSrcPE;
        }

        public void dtgdvwRefreshPE()
        {
            try
            {
                ConsultarBdPE(null, Convert.ToInt32(cmbBxPE.SelectedValue.ToString()), cmbBxPE.Text);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void cmbBxPE_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbBxPE.SelectedIndex > 0)
            {
                cmbBxPE.Text = cmbBxPE.Text.Trim();
            }
        }

        private void cmbBxPE_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbBxPE.Text))
            {
                strNomeSalvoPE = cmbBxPE.Text;
                dtgdvwRefreshPE();
            }
        }

        public void InserirBdPE(int intCliente_ID, string strNome, string strDescricao, int intEstado)
        {
            try
            {
                objPE_Bll = new PE_Bll();
                objPE_VO = new PE_VO();
                objPE_VO.Clientes_VO = new Clientes_VO();
                objPE_VO.Clientes_VO.ID = intCliente_ID <= 0 ? 0 : intCliente_ID;
                objPE_VO.Clientes_VO.Nome = strNome;
                objPE_VO.Descricao = strDescricao;
                objPE_VO.Estado = intEstado <= 0 ? 0 : intEstado;

                if (objPE_Bll.InserirBd(objPE_VO))
                {
                    MessageBox.Show("Inserir OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Inserção!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        public void ExcluirBdPE(int intID)
        {
            try
            {
                objPE_Bll = new PE_Bll();
                objPE_VO = new PE_VO();
                objPE_VO.ID = intID <= 0 ? 0 : intID;

                if (objPE_Bll.ExcluirBd(objPE_VO))
                {
                    MessageBox.Show("Excluir OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Exclusão!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        public void AlterarBdPE(int intID, int intCliente_ID, string strNome, string strDescricao, int intEstado)
        {
            try
            {
                objPE_Bll = new PE_Bll();
                objPE_VO = new PE_VO();
                objPE_VO.ID = intID <= 0 ? 0 : intID;
                objPE_VO.Clientes_VO = new Clientes_VO();
                objPE_VO.Clientes_VO.ID = intCliente_ID <= 0 ? 0 : intCliente_ID;
                objPE_VO.Clientes_VO.Nome = strNome;
                objPE_VO.Descricao = strDescricao;
                objPE_VO.Estado = intEstado <= 0 ? 0 : intEstado;

                if (objPE_Bll.AlterarBd(objPE_VO))
                {
                    MessageBox.Show("Alterar OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Alteração!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" + ex);
            }
        }

        private void dtgdvwPE_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwPE.CurrentRow.Cells["ID"].Value.ToString()) &&
                !string.IsNullOrEmpty(dtgdvwPE.CurrentRow.Cells["Cliente_ID"].Value.ToString()))
            {
                intIDSalvoPE = Convert.ToInt32(dtgdvwPE.CurrentRow.Cells["ID"].Value.ToString());
                intCliente_IDSalvoPE = Convert.ToInt32(dtgdvwPE.CurrentRow.Cells["Cliente_ID"].Value.ToString());
            }
            else
            {
                dtgdvwPE.CurrentRow.Cells["Cliente_ID"].Value = cmbBxPE.SelectedValue;
                intCliente_IDSalvoPE = Convert.ToInt32(dtgdvwPE.CurrentRow.Cells["Cliente_ID"].Value.ToString());
            }
            dtgdvwPE.CurrentRow.Cells["Cliente_ID"].Selected = false;
            dtgdvwPE.CurrentRow.Cells["Cliente_ID"].ReadOnly = true;
        }

        private void bndNavBtnPesquisarPE_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(bndNavTxtPE.Text))
                {
                    ConsultarBdPE(Convert.ToInt32(bndNavTxtPE.Text));
                }
                else
                {
                    dtgdvwRefreshPE();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro!" +ex.Message);
            }
            
        }

        private void bindingNavigatorAddNewItem2_Click(object sender, EventArgs e)
        {
            bolAddPE = true;
        }

        private void bindingNavigatorDeleteItem2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir o Pedido Interior do Cliente '" + strNomeSalvoPE + "' da Tabela Pedidos Interior? Sim ou Não?", "Exclusão de Pedidos Interior", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBdPE(intIDSalvoPE);
            }
            dtgdvwRefreshPE();
        }

        private void bndNavBtnConfirmarPE_Click(object sender, EventArgs e)
        {
            try
            {
                if (bolAddPE)
                {
                    if (MessageBox.Show("Deseja Inserir o Pedido Interior do Cliente '" + strNomeSalvoPE + "' na Tabela Pedidos Interior? Sim ou Não?", "Inserção de Pedidos Interior", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        InserirBdPE(intCliente_IDSalvoPE, strNomeSalvoPE, dtgdvwPE.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                    Convert.ToInt32(dtgdvwPE.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                    }
                    bolAddPE = false;
                }
                else
                {
                    if (MessageBox.Show("Deseja Alterar o Pedido Atual para o nov Pedido: '" + dtgdvwPE.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "' na Tabela Pedidos Interior? Sim ou Não?", "Alteração de Pedidos Interior", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        AlterarBdPE(intIDSalvoPE, intCliente_IDSalvoPE, strNomeSalvoPE, dtgdvwPE.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString(),
                                    Convert.ToInt32(dtgdvwPE.CurrentRow.Cells["Estado"].EditedFormattedValue.ToString()));
                    }
                }
                dtgdvwRefreshPE();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Excel
        private void bndNavBtnExcel_Click(object sender, EventArgs e)
        {
            AutomacaoExcelBd((DataTable)bndSrcClientes.DataSource);
        }

        public void AutomacaoExcelBd(DataTable objTabelaExcel)
        {
            try
            {
                if (objTabelaExcel != null)
                {
                    objExcelApp = new Excel.Application();
                    objExcelApp.Visible = true;
                    objExcelWb = objExcelApp.Workbooks.Add();
                    objExcelWs = objExcelWb.Worksheets[1];
                    int intColuna = 1, intLinha = 2, intLinhaCabecalho = 1;
                    objExcelCabecalho = objExcelWs.Cells[intLinhaCabecalho, intColuna];
                    objExcelDados = objExcelWs.Cells[intLinha, intColuna];

                    foreach (DataRow objLinhaBd in objTabelaExcel.Rows)
                    {
                        foreach (DataColumn objColunaBd in objTabelaExcel.Columns)
                        {
                            if (intLinha <= intLinhaCabecalho + 1)
                            {
                                objExcelCabecalho.set_Value(Type.Missing, objColunaBd.ColumnName);
                            }
                            if (!string.IsNullOrEmpty(objLinhaBd[intColuna - 1].ToString()))
                            {
                                objExcelDados.set_Value(Type.Missing, objLinhaBd[intColuna - 1].ToString());
                            }
                            intColuna++;
                            if (intLinha <= intLinhaCabecalho + 1)
                            {
                                objExcelCabecalho = objExcelWs.Cells[intLinhaCabecalho, intColuna];
                            }
                            objExcelDados = objExcelWs.Cells[intLinha, intColuna];
                        }
                        intLinha++;
                        intColuna = 1;
                        objExcelDados = objExcelWs.Cells[intLinha, intColuna];
                    }
                    sfdSalvarExcel.ShowDialog();
                    objExcelWb.SaveAs(sfdSalvarExcel.FileName.ToString(), Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                          Type.Missing, Excel.XlSaveAsAccessMode.xlShared);
                    MessageBox.Show("Operação Concluída com Sucesso!", "Excel Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    objExcelApp.Quit();
                }
                else
                {
                    MessageBox.Show("Faça a Consulta antes de Gerar Excel!", "Erro Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {

                objExcelApp.Quit();
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Email
        private void bndNavBtnEmail_Click(object sender, EventArgs e)
        {
            AutomacaoEmail();
        }

        public void AutomacaoEmail()
        {
            try
            {
                objEmailApp = new Email.Application();
                objEmailMensagem = objEmailApp.CreateItem(Email.OlItemType.olMailItem);
                objEmailMensagem.SentOnBehalfOfName = "amado.breno13@gmail.com";
                objEmailMensagem.To = "breno.luis@uol.com.br";
                objEmailMensagem.Subject = "Teste automatizado de envio de email por outlook";
                objEmailMensagem.Body = "Boa Tarde!" + Environment.NewLine +
                    "Conforme combinado, segue este texto de email para o envio automático do mesmo pelo C# \t" +
                    "(Esse é um email automático, não responda!)";

                if (MessageBox.Show("Deseja Anexar Arquivo?", "Anexação de Arquivo", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    ofdEscolhaAnexoEmail.Title = "Escolha o arquivo a ser anexado no Email";
                    ofdEscolhaAnexoEmail.InitialDirectory = @"C:\CursoProgramar";
                    ofdEscolhaAnexoEmail.ShowDialog();
                    string srtEnderecoAnexo = ofdEscolhaAnexoEmail.FileName;

                    if (!string.IsNullOrEmpty(srtEnderecoAnexo))
                    {
                        objAnexoArquivo = ofdEscolhaAnexoEmail.FileNames;

                        for (int Z = 0; Z < objAnexoArquivo.Length; Z++)
                        {
                            objAnexoTipo = Email.OlAttachmentType.olByValue;
                            objAnexoPosicao = objEmailMensagem.Body.Length + 1;
                            objDisplayName = objAnexoArquivo[Z].ToString() + "-NovoArquivo-treinoAnexo";
                            objEmailMensagem.Attachments.Add(objAnexoArquivo[Z], objAnexoTipo, objAnexoPosicao, objDisplayName);
                        }
                    }
                }
                if (MessageBox.Show("Enviar Email com Confirmação?", "Pedido de Confirmação", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    objEmailMensagem.Display();
                }
                else
                {
                    objEmailMensagem.Send();
                }
                MessageBox.Show("Email Enviado com Sucesso!", "Envio de Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Consultares
        private void btnConsultarPIePE_Click(object sender, EventArgs e)
        {
            ConsultarPIePE();
        }

        public void ConsultarPIePE()
        {
            objClientes_Bll = new Clientes_Bll();

            strIdClientesSelecionados = Interaction.InputBox("Coloque o(s) ID(s) do(s) Cliente(s) Escolhido(s) para Realizar a Consulta", "ID Cliente Consulta");

            bndSrcConsultarPIePE.DataSource = objClientes_Bll.ConsultarPIePEClientes(strIdClientesSelecionados);
            dtgdvwConsultarPIePE.DataSource = bndSrcConsultarPIePE;
        }

        private void btnConsultarQuantidadePIClientes_Click(object sender, EventArgs e)
        {
            ConsultarQuantidadePIClientes();
        }

        public void ConsultarQuantidadePIClientes()
        {
            objClientes_Bll = new Clientes_Bll();

            strIdClientesSelecionados = Interaction.InputBox("Coloque o(s) ID(s) do(s) Cliente(s) Escolhido(s) para Realizar a Consulta", "ID Cliente Consulta");

            bndSrcConsultarQuantidadePIClientes.DataSource = objClientes_Bll.ConsultarQuantidadePIClientes(strIdClientesSelecionados);
            dtgdvwConsultarQuantidadePIClientes.DataSource = bndSrcConsultarQuantidadePIClientes;
        }

        private void btnConsultarClientesSemPE_Click(object sender, EventArgs e)
        {
            ConsultarClientesSemPE();
        }

        public void ConsultarClientesSemPE()
        {
            objClientes_Bll = new Clientes_Bll();

            bndSrcConsultarClientesSemPE.DataSource = objClientes_Bll.ConsultarClientesSemPE();
            dtgdvwConsultarClientesSemPE.DataSource = bndSrcConsultarClientesSemPE;
        }

        private void btnConsultarQuantidadePedidos_Click(object sender, EventArgs e)
        {
            ConsultarQuantidadePedidos();
        }

        public void ConsultarQuantidadePedidos()
        {
            try
            {
                objClientes_Bll = new Clientes_Bll();

                bndSrcConsultarQuantidadePedidos.DataSource = objClientes_Bll.ConsultarQuantidadePedidos();
                dtgdvwConsultarQuantidadePedidos.DataSource = bndSrcConsultarQuantidadePedidos;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro" +ex.Message);
            }
        }
        #endregion

        private void frmPedidos_de_Clientes_Load(object sender, EventArgs e)
        {
            ConsultarBdClientes();
            dtgdvwRefreshPI();
            dtgdvwRefreshPE();
        }
    }
}
