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

namespace Exercicio_1_09082021
{
    public partial class frmExercicio_1_09082021 : Form
    {
        int intIDSalva;               //preferencias
        string strDescricaoSalva;

        int intCODSalva;              //familiares
        string strNomeSalvo;

        int intIDSalvaPrefFam, intCODSalvaPrefFam;                  //preferencias de familiares     
        string strNomeSalvoPrefFam, strDescricaoSalvaPrefFam;

        bool bolAddPref;

        bool bolAddFam;

        bool bolAddPrefFam;

        Preferencias_VO objPreferencias_VO;
        Preferencias_Bll objPreferencias_Bll;

        Familiares_VO objFamiliares_VO;
        Familiares_Bll objFamiliares_Bll;

        PrefFam_VO objPrefFam_VO;
        PrefFam_Bll objPrefFam_Bll;

        Excel.Application objExcelApp; //Aplicativo Excel
        Excel.Workbook objExcelWorkB; //Arquivo .xls (figura do arquivo Exel , geralmente como default 3 planilhas)
        Excel.Worksheet objExcelWorkS; //É a planilha em si , Composta de Celulas indexadas de Linhas e Colunas 
        Excel.Range objExcelCelCabecalho, objExcelCelDados; //Celulas ou Conjunto de Celulas

        Email.Application objOutlook;
        Email.MailItem objEmailMensagem;
        Email.OlAttachmentType objAnexoTipo;

        string[] objAnexoArquivo = new String[0];
        long objAnexoPosicao;
        string objDisplayName;

        public frmExercicio_1_09082021()
        {
            InitializeComponent();
        }

        #region Preferencias
        private void btnDesvCond_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Escolha Sim ou Não!", "Desvio Condicional", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                MessageBox.Show("Você Escolheu Sim!");
            }
            else
            {
                MessageBox.Show("Você Escolheu Não!");
            }
        }

        private void btnImportarTxt_Click(object sender, EventArgs e)
        {
            lstbxPreferencias.Items.Clear();
            objPreferencias_Bll = new Preferencias_Bll();
            lstbxPreferencias.Items.AddRange(objPreferencias_Bll.ImportarTxt().ToArray());
        }

        private void btnImportarBdConectado_Click(object sender, EventArgs e)
        {
            lstbxPreferencias.Items.Clear();
            objPreferencias_Bll = new Preferencias_Bll();
            lstbxPreferencias.Items.AddRange(objPreferencias_Bll.ImportarBdConectado().ToArray());
        }

        private void btnImportarBdDesconectado_Click(object sender, EventArgs e)
        {
            lstbxPreferencias.Items.Clear();
            objPreferencias_Bll = new Preferencias_Bll();
            lstbxPreferencias.Items.AddRange(objPreferencias_Bll.ImportarBdDesconectado().ToArray());
        }

        private void btnConsultarBd_Click(object sender, EventArgs e)
        {
            ConsultarBd();
        }

        public void ConsultarBd(int? intID = null, string strDescricao = null)
        {
            try
            {
                objPreferencias_Bll = new Preferencias_Bll();
                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);
                objPreferencias_VO.Descricao = strDescricao;

                bndSrcPreferencia.DataSource = objPreferencias_Bll.ConsultarBd(objPreferencias_VO);
                dtgdvwPreferencias.DataSource = bndSrcPreferencia;
            }
            catch (Exception ex)
            {

                MessageBox.Show("Ocorreu um Erro ==>" + ex);
            }
        }

        private void btnInserirBd_Click(object sender, EventArgs e)
        {
            InserirBd(dtgdvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
            ConsultarBd();
        }

        public void InserirBd(string strDescricao)
        {
            try
            {
                objPreferencias_Bll = new Preferencias_Bll();
                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.Descricao = strDescricao;

                if (objPreferencias_Bll.InserirBd(objPreferencias_VO))
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

                MessageBox.Show("Ocorreu um Erro ==>" + ex);
            }
        }

        private void btnExcluirBd_Click(object sender, EventArgs e)
        {
            ExcluirBd(intIDSalva);
            ConsultarBd();
        }

        public void ExcluirBd(int intID)
        {
            try
            {
                objPreferencias_Bll = new Preferencias_Bll();
                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.ID = intID;

                if (objPreferencias_Bll.ExcluirBd(objPreferencias_VO))
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

                MessageBox.Show("Ocorreu um Erro ==>" + ex);
            }
        }

        private void btnAlterarBd_Click(object sender, EventArgs e)
        {
            AlterarBd(intIDSalva, dtgdvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
            ConsultarBd();
        }

        public void AlterarBd(int intIDPreferencia, string strPreferenciaNova)
        {
            try
            {
                objPreferencias_Bll = new Preferencias_Bll();
                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.ID = intIDPreferencia;
                objPreferencias_VO.Descricao = strPreferenciaNova;

                if (objPreferencias_Bll.AlterarBd(objPreferencias_VO))
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

                MessageBox.Show("Ocorreu um Erro ==>" + ex);
            }
        }

        private void dtgdvwPreferencias_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwPreferencias.CurrentRow.Cells["ID"].Value.ToString()))
            {
                intIDSalva = Convert.ToInt32(dtgdvwPreferencias.CurrentRow.Cells["ID"].Value.ToString());
            }
            strDescricaoSalva = dtgdvwPreferencias.CurrentRow.Cells["Descricao"].Value.ToString();
        }

        private void bndNavBtnPesquisar_Click(object sender, EventArgs e)
        {
            ConsultarBd(null, bndNavTxtPreferencia.Text);
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            bolAddPref = true;
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir a Preferencia '" + strDescricaoSalva + "' da Tabela Preferencias? Sim ou Não?", "Exclusão de Preferencias", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBd(intIDSalva);
            }

            ConsultarBd();
        }

        private void bndNavBtnConfirmar_Click(object sender, EventArgs e)
        {
            if (bolAddPref)
            {
                if (MessageBox.Show("Deseja Inserir a Preferencia '" + dtgdvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "' na Tabela Preferencias? Sim ou Não?", "Inserção de Preferencias", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    InserirBd(dtgdvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
                }
                bolAddPref = false;
            }
            else
            {
                if (MessageBox.Show("Deseja Alterar a Preferencia '" + strDescricaoSalva + "' para nova Preferencia: '" + dtgdvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString() + "' na Tabela Preferencias? Sim ou Não?", "Alteração de Preferencias", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    AlterarBd(intIDSalva, dtgdvwPreferencias.CurrentRow.Cells["Descricao"].EditedFormattedValue.ToString());
                }
            }

            ConsultarBd();
        }

        private void bndNavBtnExportarExcelGridPref_Click(object sender, EventArgs e)
        {
            AutomacaoExcelGrid(dtgdvwPreferencias);
        }

        private void bndNavBtnExportarExcelBdPref_Click(object sender, EventArgs e)
        {
            AutomacaoExcelBd((DataTable)bndSrcPreferencia.DataSource);  //casting para o datasource virar uma datatable no caso
        }

        private void bndNavBtnExportarExcelPrefAccess_Click(object sender, EventArgs e)
        {
            AutomacaoExcelAccess();
        }
        #endregion

        #region Familiares

        public void ConsultarBdFam(int? intCOD = null, string strNome = null, string strSexo = null, double? dbGanho = null, double? dbGasto = null, string strObservacao = null)
        {
            try
            {
                objFamiliares_Bll = new Familiares_Bll();
                objFamiliares_VO = new Familiares_VO();

                if (objFamiliares_VO.COD > 0)
                {
                    objFamiliares_VO.COD = Convert.ToInt32(intCOD);
                }

                objFamiliares_VO.Nome = strNome;


                //inicialização das colunas
                dtgdvwFamiliares.Columns.Clear();
                dtgdvwFamiliares.DataSource = null;
                dtgdvwFamiliares.AllowUserToAddRows = false;

                dtgdvwFamiliares.Columns.Add("COD", "Código do Familiar");
                dtgdvwFamiliares.Columns["COD"].DataPropertyName = "COD";

                dtgdvwFamiliares.Columns.Add("Nome", "Nome do Familiar");
                dtgdvwFamiliares.Columns["Nome"].DataPropertyName = "Nome";

                DataGridViewComboBoxColumn objcbbxColumnFamiliaresSexo = new DataGridViewComboBoxColumn();
                objcbbxColumnFamiliaresSexo.Name = "Sexo";
                objcbbxColumnFamiliaresSexo.ValueType = typeof(string);
                objcbbxColumnFamiliaresSexo.HeaderText = "Sexo do Familiar";

                objcbbxColumnFamiliaresSexo.Items.Add("MASCULINO");
                objcbbxColumnFamiliaresSexo.Items.Add("FEMININO");
                objcbbxColumnFamiliaresSexo.Items.Add("OUTRO");
                objcbbxColumnFamiliaresSexo.DataPropertyName = "Sexo";

                dtgdvwFamiliares.Columns.Add(objcbbxColumnFamiliaresSexo);
                dtgdvwFamiliares.Columns["Sexo"].DataPropertyName = "Sexo";

                dtgdvwFamiliares.Columns.Add("Idade", "Idade do Familiar");
                dtgdvwFamiliares.Columns["Idade"].DataPropertyName = "Idade";

                dtgdvwFamiliares.Columns.Add("Ganho_Mensal_Total", "Ganho Mensal do Familiar");
                dtgdvwFamiliares.Columns["Ganho_Mensal_Total"].DataPropertyName = "Ganho_Mensal_Total";

                dtgdvwFamiliares.Columns.Add("Gasto_Mensal_Total", "Gasto Mensal do Familiar");
                dtgdvwFamiliares.Columns["Gasto_Mensal_Total"].DataPropertyName = "Gasto_Mensal_Total";

                dtgdvwFamiliares.Columns.Add("Observacao", "Observacao do Familiar");
                dtgdvwFamiliares.Columns["Observacao"].DataPropertyName = "Observacao";

                bndSrcFamiliares.DataSource = objFamiliares_Bll.ConsultarBd(objFamiliares_VO);
                dtgdvwFamiliares.DataSource = bndSrcFamiliares;

                //Consulta de Preferencia de Familiar será embazada por esse ComboBox 
                cmbbxFamiliares.DataSource = null;
                cmbbxFamiliares.Items.Clear();

                cmbbxFamiliares.DisplayMember = "Nome";
                cmbbxFamiliares.ValueMember = "COD";

                //O lugar original do DataSource é no inicio da config do ComboBox 
                //, mas para funcionamento do codigo ele precisou ser montado apos a config das propriedades display e Value Member
                cmbbxFamiliares.DataSource = bndSrcFamiliares.DataSource;
                cmbbxFamiliares.SelectedIndex = Convert.ToInt32(intCOD > 0 ? intCOD - 1 : 0);

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            
        }

        public void InserirBdFam(string strNome, string strSexo, int intIdade, double dbGanho, double dbGasto, string strObservacao)
        {
            try
            {
                objFamiliares_Bll = new Familiares_Bll();
                objFamiliares_VO = new Familiares_VO();

                objFamiliares_VO.Nome = strNome;

                objFamiliares_VO.Sexo = strSexo;

                if (intIdade > 0)
                {
                    objFamiliares_VO.Idade = Convert.ToInt32(intIdade);
                }

                if (dbGanho > 0)
                {
                    objFamiliares_VO.Ganho = Convert.ToDouble(dbGanho);
                }

                if (dbGasto > 0)
                {
                    objFamiliares_VO.Gasto = Convert.ToDouble(dbGasto);
                }

                objFamiliares_VO.Observacao = strObservacao;

                if (objFamiliares_Bll.InserirBd(objFamiliares_VO))
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

                MessageBox.Show(ex.Message);
            }
        }

        public void ExcluirBdFam(int intCOD)
        {
            try
            {
                objFamiliares_Bll = new Familiares_Bll();
                objFamiliares_VO = new Familiares_VO();

                if (intCOD > 0)
                {
                    objFamiliares_VO.COD = Convert.ToInt32(intCOD);
                }

                if (objFamiliares_Bll.ExcluirBd(objFamiliares_VO))
                {
                    MessageBox.Show("Exclusao OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Exclusão!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        public void AlterarBdFam(int intCOD, string strNome, string strSexo, int intIdade, double dbGanho, double dbGasto, string strObservacao)
        {
            try
            {
                objFamiliares_Bll = new Familiares_Bll();
                objFamiliares_VO = new Familiares_VO();

                if (intCOD > 0)
                {
                    objFamiliares_VO.COD = Convert.ToInt32(intCOD);
                }

                objFamiliares_VO.Nome = strNome;

                objFamiliares_VO.Sexo = strSexo;

                if (intIdade > 0)
                {
                    objFamiliares_VO.SetIdade(Convert.ToInt32(intIdade));
                }

                if (dbGanho > 0)
                {
                    objFamiliares_VO.Ganho = Convert.ToDouble(dbGanho);
                }

                if (dbGasto > 0)
                {
                    objFamiliares_VO.Gasto = Convert.ToDouble(dbGasto);
                }

                objFamiliares_VO.Observacao = strObservacao;

                if (objFamiliares_Bll.AlterarBd(objFamiliares_VO))
                {
                    MessageBox.Show("Alteração OK!");
                }
                else
                {
                    MessageBox.Show("Falha na Alteração!");
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void dtgdvwFamiliares_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrEmpty(dtgdvwFamiliares.CurrentRow.Cells["COD"].Value.ToString()))
            {
                intCODSalva = Convert.ToInt32(dtgdvwFamiliares.CurrentRow.Cells["COD"].Value.ToString());
            }

            strNomeSalvo = dtgdvwFamiliares.CurrentRow.Cells["Nome"].Value.ToString();
        }

        private void bindingNavigatorDeleteItem1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Deseja Excluir o Familiar '" + strNomeSalvo + "' da Tabela Familiares? Sim ou Não?", "Exclusão de Familiares", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ExcluirBdFam(intCODSalva);
            }
            ConsultarBdFam();
        }

        private void bindingNavigatorAddNewItem1_Click(object sender, EventArgs e)
        {
            bolAddFam = true;
        }

        private void bndNavBtnConFam_Click(object sender, EventArgs e)
        {
            try
            {
                if (bolAddFam)
                {
                    if (MessageBox.Show("Deseja Inserir Elementos do Familiar: '" + dtgdvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString() + "' na Tabela Familiares? Sim ou Não?", "Inserção de Familiares", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        InserirBdFam(dtgdvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                     dtgdvwFamiliares.CurrentRow.Cells["Sexo"].EditedFormattedValue.ToString(),
                                     Convert.ToInt32(dtgdvwFamiliares.CurrentRow.Cells["Idade"].EditedFormattedValue.ToString()),
                                     Convert.ToDouble(dtgdvwFamiliares.CurrentRow.Cells["Ganho_Mensal_Total"].EditedFormattedValue.ToString()),
                                     Convert.ToDouble(dtgdvwFamiliares.CurrentRow.Cells["Gasto_Mensal_Total"].EditedFormattedValue.ToString()),
                                     dtgdvwFamiliares.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                    }
                    bolAddFam = false;
                }
                else
                {
                    if (MessageBox.Show("Deseja Alterar Elementos do Familiar: '" + strNomeSalvo + "' para Elementos do Novo Familiar: '" + dtgdvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString() + "' na Tabela Familiares? Sim ou Não?", "Alteração de Familiares", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        AlterarBdFam(intCODSalva, dtgdvwFamiliares.CurrentRow.Cells["Nome"].EditedFormattedValue.ToString(),
                                     dtgdvwFamiliares.CurrentRow.Cells["Sexo"].EditedFormattedValue.ToString(),
                                     Convert.ToInt32(dtgdvwFamiliares.CurrentRow.Cells["Idade"].EditedFormattedValue.ToString()),
                                     Convert.ToDouble(dtgdvwFamiliares.CurrentRow.Cells["Ganho_Mensal_Total"].EditedFormattedValue.ToString()),
                                     Convert.ToDouble(dtgdvwFamiliares.CurrentRow.Cells["Gasto_Mensal_Total"].EditedFormattedValue.ToString()),
                                     dtgdvwFamiliares.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                    }
                }
                ConsultarBdFam();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void bndNavBtnPesquiFam_Click(object sender, EventArgs e)
        {
            ConsultarBdFam(null, bndNavTxtFamiliares.Text);
        }

        private void bndNavBtnExportarExcelGridFam_Click(object sender, EventArgs e)
        {
            AutomacaoExcelGrid(dtgdvwFamiliares);
        }

        private void bndNavBtnExportarExcelBdFam_Click(object sender, EventArgs e)
        {
            AutomacaoExcelBd((DataTable)bndSrcFamiliares.DataSource);
        }

        #endregion

        #region Preferencias de Familiares

        public void ConsultarBdPrefFam(int? intCOD = null, int? intID = null, string strNome = null, string strDescricao = null)
        {
            try
            {
                objPrefFam_Bll = new PrefFam_Bll();
                objPrefFam_VO = new PrefFam_VO();

                objPrefFam_VO.Familiares_VO = new Familiares_VO();
                objPrefFam_VO.Familiares_VO.COD = Convert.ToInt32(intCOD == null ? 0 : intCOD);
                objPrefFam_VO.Familiares_VO.Nome = strNome;

                objPrefFam_VO.Preferencias_VO = new Preferencias_VO();
                objPrefFam_VO.Preferencias_VO.ID = Convert.ToInt32(intID == null ? 0 : intID);
                objPrefFam_VO.Preferencias_VO.Descricao = strDescricao;

                bndSrcPrefFam.DataSource = objPrefFam_Bll.ConsultarBd(objPrefFam_VO);

                dtgdvwPrefFam.DataSource = null;
                dtgdvwPrefFam.Columns.Clear();
                dtgdvwPrefFam.AllowUserToAddRows = false;

                DataGridViewComboBoxColumn objDtgdvwCmbbxClFamLookUp = new DataGridViewComboBoxColumn();
                objDtgdvwCmbbxClFamLookUp.DataSource = bndSrcFamiliares.DataSource;
                objDtgdvwCmbbxClFamLookUp.Name = "COD";
                objDtgdvwCmbbxClFamLookUp.ValueType = typeof(int);
                objDtgdvwCmbbxClFamLookUp.DisplayMember = "Nome";
                objDtgdvwCmbbxClFamLookUp.ValueMember = "COD";
                objDtgdvwCmbbxClFamLookUp.HeaderText = "Identificacao do Familiar";

                dtgdvwPrefFam.Columns.Add(objDtgdvwCmbbxClFamLookUp);
                dtgdvwPrefFam.Columns["COD"].DataPropertyName = "COD";

                Preferencias_Bll preferencias_Bll = new Preferencias_Bll();

                bndSrcPreferenciaLookUp.DataSource = objPreferencias_Bll.ConsultarBd(new Preferencias_VO());

                DataGridViewComboBoxColumn objDtgdvwCmbbxClPrefLookUp = new DataGridViewComboBoxColumn();
                objDtgdvwCmbbxClPrefLookUp.DataSource = bndSrcPreferenciaLookUp.DataSource;
                objDtgdvwCmbbxClPrefLookUp.Name = "ID";
                objDtgdvwCmbbxClPrefLookUp.ValueType = typeof(int);
                objDtgdvwCmbbxClPrefLookUp.DisplayMember = "Descricao";
                objDtgdvwCmbbxClPrefLookUp.ValueMember = "ID";
                objDtgdvwCmbbxClPrefLookUp.HeaderText = "Identificacao da Preferencia";
                objDtgdvwCmbbxClPrefLookUp.DataPropertyName = "ID";

                dtgdvwPrefFam.Columns.Add(objDtgdvwCmbbxClPrefLookUp);
                dtgdvwPrefFam.Columns["ID"].ValueType = typeof(int);
                dtgdvwPrefFam.Columns["ID"].DataPropertyName = "ID";

                dtgdvwPrefFam.Columns.Add("Intensidade", "Intensidade da Preferencia do Familiar");
                dtgdvwPrefFam.Columns["Intensidade"].DataPropertyName = "Intensidade";

                dtgdvwPrefFam.Columns.Add("Observacao", "Observacao da Preferencia do Familiar");
                dtgdvwPrefFam.Columns["Observacao"].DataPropertyName = "Observacao";

                dtgdvwPrefFam.DataSource = bndSrcPrefFam;

                bndNavCmbbxPrefFam.Items.Clear();

                bndNavCmbbxPrefFam.Items.Add("0-");

                foreach (DataRow objPreferenciaLinha in ((DataTable)bndSrcPreferencia.DataSource).Rows)  //casting
                {
                    bndNavCmbbxPrefFam.Items.Add(objPreferenciaLinha["ID"].ToString() + "-" + objPreferenciaLinha["Descricao"].ToString());
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void cmbbxFamiliares_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (cmbbxFamiliares.SelectedIndex >= 0)
            if (((ComboBox)sender).SelectedIndex >= 0)
            {
                ((ComboBox)sender).Text = ((ComboBox)sender).Text.Trim();
            }
        }

        private void cmbbxFamiliares_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbbxFamiliares.Text))
            {
                dtgdvwPrefFamRefresh();
            }
        }

        public void dtgdvwPrefFamRefresh()
        {
            try
            {
                ConsultarBdPrefFam(Convert.ToInt32(cmbbxFamiliares.SelectedValue.ToString()), null, cmbbxFamiliares.Text);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        public void InserirBdPrefFam(int intCOD, int intID, string strNome, string strDescricao, double dbIntensidade, string strObervacao)
        {
            try
            {
                objFamiliares_VO = new Familiares_VO();
                objFamiliares_VO.COD = intCOD <= 0 ? 0 : intCOD;
                objFamiliares_VO.Nome = strNome;

                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.ID = intID <= 0 ? 0 : intID;
                objPreferencias_VO.Descricao = strDescricao;

                objPrefFam_VO = new PrefFam_VO(objFamiliares_VO, objPreferencias_VO, dbIntensidade, strObervacao);

                objPrefFam_Bll = new PrefFam_Bll();

                if (objPrefFam_Bll.InserirBd(objPrefFam_VO))
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

                MessageBox.Show(ex.Message);
            }
        }

        public void AlterarBdPrefFam(int intCOD, int intID, string strNome, string strDescricao, double dbIntensidade, string strObervacao)
        {
            try
            {
                objFamiliares_VO = new Familiares_VO();
                objFamiliares_VO.COD = intCOD <= 0 ? 0 : intCOD;
                objFamiliares_VO.Nome = strNome;

                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.ID = intID <= 0 ? 0 : intID;
                objPreferencias_VO.Descricao = strDescricao;

                objPrefFam_VO = new PrefFam_VO(objFamiliares_VO, objPreferencias_VO, dbIntensidade, strObervacao);

                objPrefFam_Bll = new PrefFam_Bll();

                if (objPrefFam_Bll.AlterarBd(objPrefFam_VO))
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

                MessageBox.Show(ex.Message);
            }
        }

        public void ExcluirBdPrefFam(int intCOD, int intID, string strNome, string strDescricao)
        {
            try
            {
                objFamiliares_VO = new Familiares_VO();
                objFamiliares_VO.COD = intCOD <= 0 ? 0 : intCOD;
                objFamiliares_VO.Nome = strNome;

                objPreferencias_VO = new Preferencias_VO();
                objPreferencias_VO.ID = intID <= 0 ? 0 : intID;
                objPreferencias_VO.Descricao = strDescricao;

                objPrefFam_VO = new PrefFam_VO();
                objPrefFam_VO.Familiares_VO = objFamiliares_VO;
                objPrefFam_VO.Preferencias_VO = objPreferencias_VO;

                objPrefFam_Bll = new PrefFam_Bll();

                if (objPrefFam_Bll.ExcluirBd(objPrefFam_VO))
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

                MessageBox.Show(ex.Message);
            }
        }

        public void ReconfigurarBd(bool bolCODSelected, bool bolCODReadOnly, bool bolIDSelected, bool bolIDReadOnly)
        {
            dtgdvwPrefFam.CurrentRow.Cells["COD"].Selected = bolCODSelected;
            dtgdvwPrefFam.CurrentRow.Cells["COD"].ReadOnly = bolCODReadOnly;
            dtgdvwPrefFam.CurrentRow.Cells["ID"].Selected = bolIDSelected;
            dtgdvwPrefFam.CurrentRow.Cells["ID"].ReadOnly = bolIDReadOnly;
        }

        private void bindingNavigatorAddNewItem2_Click(object sender, EventArgs e)
        {
            bolAddPrefFam = true;
            ReconfigurarBd(false, true, true, false);
        }

        private void bindingNavigatorDeleteItem2_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Deseja Excluir esta Preferencia de Familiar? Sim ou Não?", "Exclusão de Preferencia de Familiar", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
            ExcluirBdPrefFam(intCODSalvaPrefFam, intIDSalvaPrefFam, strNomeSalvoPrefFam, strDescricaoSalvaPrefFam);

                }
                dtgdvwPrefFamRefresh();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void bndNavBtnConfirmarPrefFam_Click(object sender, EventArgs e)
        {
            try
            {
                if (bolAddPrefFam)
                {
                    if (MessageBox.Show("Incluir Nova Preferencia de Familiar", "Incluir Preferencia de Familiar", MessageBoxButtons.YesNoCancel) == System.Windows.Forms.DialogResult.Yes)
                    {
                        InserirBdPrefFam(Convert.ToInt32(dtgdvwPrefFam.CurrentRow.Cells["COD"].Value.ToString()),
                            Convert.ToInt32(dtgdvwPrefFam.CurrentRow.Cells["ID"].Value.ToString()),
                            dtgdvwPrefFam.CurrentRow.Cells["COD"].EditedFormattedValue.ToString(),
                            dtgdvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString(),
                            Convert.ToDouble(dtgdvwPrefFam.CurrentRow.Cells["Intensidade"].EditedFormattedValue),
                            dtgdvwPrefFam.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());
                        bolAddPrefFam = false;
                        ReconfigurarBd(false, true, false, true);
                    }

                }
                else
                {
                    if (MessageBox.Show("Alterar o elemento de Nome: " + strNomeSalvoPrefFam, "Alterar Preferencia De Familiar", MessageBoxButtons.YesNoCancel) == System.Windows.Forms.DialogResult.Yes)
                    {
                        AlterarBdPrefFam(intCODSalvaPrefFam,
                                            intIDSalvaPrefFam,
                                            strNomeSalvoPrefFam,
                                            strDescricaoSalvaPrefFam,
                                            Convert.ToDouble(dtgdvwPrefFam.CurrentRow.Cells["Intensidade"].EditedFormattedValue.ToString()),
                                            dtgdvwPrefFam.CurrentRow.Cells["Observacao"].EditedFormattedValue.ToString());

                    }
                }
                dtgdvwPrefFamRefresh();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
            
        }

        private void dtgdvwPrefFam_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(dtgdvwPrefFam.CurrentRow.Cells["COD"].EditedFormattedValue.ToString()) && !string.IsNullOrEmpty(dtgdvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString()))
                {
                    intIDSalvaPrefFam = Convert.ToInt32(dtgdvwPrefFam.CurrentRow.Cells["ID"].Value.ToString());
                    intCODSalvaPrefFam = Convert.ToInt32(dtgdvwPrefFam.CurrentRow.Cells["COD"].Value.ToString());
                    strDescricaoSalvaPrefFam = dtgdvwPrefFam.CurrentRow.Cells["ID"].EditedFormattedValue.ToString();
                    strNomeSalvoPrefFam = dtgdvwPrefFam.CurrentRow.Cells["COD"].EditedFormattedValue.ToString();

                    ReconfigurarBd(false, true, false, true);
                }
                else
                {
                    dtgdvwPrefFam.CurrentRow.Cells["COD"].Value = cmbbxFamiliares.SelectedValue;
                    ReconfigurarBd(false, true, true, false);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void bndNavBtnPesquisarPrefFam_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(bndNavCmbbxPrefFam.Text.Trim()))
            {
                ConsultarBdPrefFam(Convert.ToInt32(cmbbxFamiliares.SelectedValue),
                                   Convert.ToInt32(bndNavCmbbxPrefFam.Text.Substring(0, bndNavCmbbxPrefFam.Text.IndexOf("-"))),
                                   cmbbxFamiliares.Text,
                                   bndNavCmbbxPrefFam.Text.Substring(bndNavCmbbxPrefFam.Text.IndexOf("-") + 1));
            }
            else
            {
                dtgdvwPrefFamRefresh();
            }
        }

        private void bndNavBtnExportarExcelGridPrefFam_Click(object sender, EventArgs e)
        {
            AutomacaoExcelGrid(dtgdvwPrefFam);
        }

        private void bndNavBtnExportarExcelBdPrefFam_Click(object sender, EventArgs e)
        {
            AutomacaoExcelBd((DataTable)bndSrcPrefFam.DataSource);
        }

        #endregion

        public void AutomacaoExcelGrid(DataGridView dtgdvwModel)
        {
            objExcelApp = new Excel.Application();
            objExcelApp.Visible = true;
            objExcelWorkB = objExcelApp.Workbooks.Add();
            objExcelWorkS = objExcelWorkB.Worksheets[1];

            int intColuna = 1, intLinha = 2, intLinhaCabecalho = 1;

            objExcelCelCabecalho = objExcelWorkS.Cells[intLinhaCabecalho, intColuna];
            objExcelCelDados = objExcelWorkS.Cells[intLinha, intColuna];

            foreach (DataGridViewRow objLinhaGrid in dtgdvwModel.Rows)
            {
                foreach (DataGridViewColumn objColunaGrid in dtgdvwModel.Columns)
                {
                    if (intLinha <= intLinhaCabecalho + 1)
                    {
                        objExcelCelCabecalho.set_Value(Type.Missing, objColunaGrid.HeaderText.ToString());
                    }

                    if (objLinhaGrid.Cells[intColuna - 1].Value != null)
                    {
                        objExcelCelDados.set_Value(Type.Missing, objLinhaGrid.Cells[intColuna - 1].Value);
                    }

                    intColuna++; // como se fosse -> intColuna = intColuna +1

                    if (intLinha <= intLinhaCabecalho + 1)
                    {
                        objExcelCelCabecalho = objExcelWorkS.Cells[intLinhaCabecalho, intColuna];
                    }

                    objExcelCelDados = objExcelWorkS.Cells[intLinha, intColuna];
                }

                intLinha++;
                intColuna = 1;

                objExcelCelDados = objExcelWorkS.Cells[intLinha, intColuna];
            }

            objExcelWorkB.SaveAs(@"C:\CursoProgramar\Excel_Saves\ExcelGrid" + dtgdvwModel.Name.Substring(6) + " " +
                DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Excel.XlSaveAsAccessMode.xlShared);
            objExcelApp.Quit();

            MessageBox.Show("Exportação para ExcelGrid Concluida Bixo " + dtgdvwModel.Name.Substring(6) + " " +
                DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx", "Exportar ExcelGrid", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }
        
        public void AutomacaoExcelBd(DataTable objTabelaBd)
        {
            objExcelApp = new Excel.Application();
            objExcelApp.Visible = true;
            objExcelWorkB = objExcelApp.Workbooks.Add();
            objExcelWorkS = objExcelWorkB.Worksheets[1];

            int intColuna = 1, intLinha = 2, intLinhaCabecalho = 1;

            objExcelCelCabecalho = objExcelWorkS.Cells[intLinhaCabecalho, intColuna];
            objExcelCelDados = objExcelWorkS.Cells[intLinha, intColuna];

            foreach (DataRow objLinhaBd in objTabelaBd.Rows)
            {
                foreach (DataColumn objColunaBd in objTabelaBd.Columns)
                {
                    if (intLinha <= intLinhaCabecalho + 1)
                    {
                        objExcelCelCabecalho.set_Value(Type.Missing, objColunaBd.ColumnName);
                    }

                    if (!string.IsNullOrEmpty(objLinhaBd[intColuna - 1].ToString()))
                    {
                        objExcelCelDados.set_Value(Type.Missing, objLinhaBd[intColuna - 1].ToString());
                    }

                    intColuna++; // como se fosse -> intColuna = intColuna +1

                    if (intLinha <= intLinhaCabecalho + 1)
                    {
                        objExcelCelCabecalho = objExcelWorkS.Cells[intLinhaCabecalho, intColuna];
                    }

                    objExcelCelDados = objExcelWorkS.Cells[intLinha, intColuna];
                }

                intLinha++;
                intColuna = 1;

                objExcelCelDados = objExcelWorkS.Cells[intLinha, intColuna];
            }

            objExcelWorkB.SaveAs(@"C:\CursoProgramar\Excel_Saves\ExcelBd" + objTabelaBd.TableName.ToString() + " " +
                DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Excel.XlSaveAsAccessMode.xlShared);
            objExcelApp.Quit();

            MessageBox.Show("Exportação para ExcelBd Concluida Bixo " + objTabelaBd.TableName.ToString() + " " +
                DateTime.Now.ToString().Replace("/", "-").Replace(":", "-") + ".xlsx", "Exportar ExcelBd", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

        }

        public void AutomacaoExcelAccess()
        {
            try
            {
                //Instacia a Bll para chamar a Geração do Excell por Interop do Access
                objPreferencias_Bll = new Preferencias_Bll();

                //Trabalha com Janela de Diálogo para Salvar Arquivos - Save File Dialog
                //if (sfdPlanilhaInterop.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                //{
                //    //Chamo o método de geração da Planilha da BLL através de Interop do Access
                //    objPreferencia.GerarExcelAccessPorInterOp(sfdPlanilhaInterop.FileName);

                //}
                sfdPlanilhaInterOp.ShowDialog();
                objPreferencias_Bll.GerarExcelAccessPorInterOp(sfdPlanilhaInterOp.FileName);


                MessageBox.Show("Final da Geração do Excel Por Interoperabilidade Do Banco de Dados Access "
                    + objPreferencias_Bll.GetType().Name + " do arruivo "
                    + sfdPlanilhaInterOp.FileName, "Gerar Excel Do Banco De Dados",
                    MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private void bndNavBtnEnviarEmailPrefFam_Click(object sender, EventArgs e)
        {
            enviarEmail();
        }

        private void enviarEmail()
        {
            objOutlook = new Email.Application();
            objEmailMensagem = objOutlook.CreateItem(Email.OlItemType.olMailItem);
            objEmailMensagem.SentOnBehalfOfName = "amado.brenin@outlook.com";
            objEmailMensagem.To = "diegochavesdds@gmail.com";
            objEmailMensagem.CC = "breno.luis@uol.com.br";
            objEmailMensagem.BCC = "adr.sud.cor@gmail.com";
            objEmailMensagem.Subject = "Teste automatizado de envio de email por outlook";
            objEmailMensagem.Body = "Pessoal, Boa Dia!" + Environment.NewLine +
                "Conforme combinado, segue esse texto de e-mail para o envio automático do mesmo pelo C# \t" +
                "(Esse é um email automático não responda)";

            if (MessageBox.Show("Deseja Anexar Arquivos", "Arquivos Anexados", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ofdEscolhaAnexoEmail.Title = "Escolha os arquivos a serem anexados no Email";
                ofdEscolhaAnexoEmail.InitialDirectory = @"C:\CursoProgramar";
                ofdEscolhaAnexoEmail.ShowDialog();
                string strEnderecoAnexo = ofdEscolhaAnexoEmail.FileName;

                if (!string.IsNullOrEmpty(strEnderecoAnexo))
                {
                    objAnexoArquivo = ofdEscolhaAnexoEmail.FileNames;

                    for (int Z = 0; Z < objAnexoArquivo.Length; Z++)
                    {
                        objAnexoTipo = Email.OlAttachmentType.olByValue;
                        objAnexoPosicao = objEmailMensagem.Body.Length + 1;
                        objDisplayName = objAnexoArquivo[Z].ToString() + " - NovoArquivo-treinoAnexo";
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
            MessageBox.Show("Email enviado com Sucesso!", "Envio de Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void frmExercicio_1_09082021_Load(object sender, EventArgs e)
        {
            ConsultarBd();
            ConsultarBdFam();
            dtgdvwPrefFamRefresh();
        }
    }
}
