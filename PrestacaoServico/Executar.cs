using PrestacaoServico.Utils;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestacaoServico
{
    class Executar
    {
        public static SAPbobsCOM.Company Company;
        public static SAPbouiCOM.Application Application;


        public Executar()
        {

            try
            {
                SetApplication();
                CompanyConnection();
                CriarTabelasCampos();

                new Eventos();
                Menu mn = new Menu();
                mn.CriarMenus();
            }
            catch
            {
                System.Environment.Exit(0);
            }
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();

            string sConnectionString = System.Convert.ToString("0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056");

            SboGuiApi.Connect(sConnectionString);
            Application = SboGuiApi.GetApplication(-1);
        }


        private void CompanyConnection()
        {
            try
            {
                Company = (SAPbobsCOM.Company)Application.Company.GetDICompany();
            }
            catch
            {
                Application.StatusBar.SetText(Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void CriarTabelasCampos()
        {
            try
            {


                MigrateTabelaCampos.CriarTabela("ACO_SERVICO", "Serviços", BoUTBTableType.bott_Document);
                MigrateTabelaCampos.CriarTabela("ACO_SERVICO_1", "Serviços 1", BoUTBTableType.bott_DocumentLines);
                MigrateTabelaCampos.CriarTabela("ACO_SERVICO_2", "Serviços 2", BoUTBTableType.bott_DocumentLines);

                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "Descricao", "Descrição", BoFieldTypes.db_Memo, BoFldSubTypes.st_None);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "Cliente", "Cliente", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "Funcionario", "Funcionário", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "Data", "Data", BoFieldTypes.db_Date, BoFldSubTypes.st_None, 8);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "Prioridade", "Prioridade", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "SM", "Saida de Mercadoria", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO", "LCM", "Lançamento Contabil", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 10);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO_1", "Item", "Item", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO_1", "Quantidade", "Quantidade", BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO_1", "Deposito", "Deposito", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO_2", "Descricao", "Descrição", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO_2", "Valor", "Valor", BoFieldTypes.db_Float, BoFldSubTypes.st_Price);
                MigrateTabelaCampos.CriarCampos("@ACO_SERVICO_2", "Tipo", "Tipo", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1);
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro ao criar Tabelas e Campos: {ex.Message}.");
            }

        }
    }
}
