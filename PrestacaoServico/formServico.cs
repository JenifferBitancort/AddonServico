using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestacaoServico
{
    public class formServico
    {
        public void ShowForm()
        {
            try
            {
                //Exibir Form
                FormCreationParams oCreationParams = (FormCreationParams)Executar.Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
                oCreationParams.XmlData = Properties.Resources.formServ;
                oCreationParams.UniqueID = "Sv" + Guid.NewGuid().ToString().Substring(0, 10);
                oCreationParams.FormType = "Sv";
                Form form = Executar.Application.Forms.AddEx(oCreationParams);
                form.Visible = true;
                form.Title = "Serviços";

                //Não exibir coluna
                form.Items.Item("grdMat").AffectsFormMode = true;
                Grid grd = form.Items.Item("grdMat").Specific;
                grd.Columns.Item(0).Visible = false;

                form.Items.Item("Item_9").Click();


                //Grids
                DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                carregarGridTarefas(form);

                Executar.Application.StatusBar.SetText("Tela iniciada", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro ao criar Exibir Form: {ex.Message}.");
            }
        }
        private void carregarGridMateriais(Form form, string DocEntry)
        {
            try
            {
                Recordset ds = (Recordset)Executar.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string consulta = $@"SELECT * FROM ""@ACO_SERVICO_1"" WHERE ""DocEntry"" = {DocEntry}";
                ds.DoQuery(consulta);

                SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("DT_Mat");

                tb.Rows.Clear();

                if (ds.RecordCount > 0)
                {
                    while (!ds.EoF)
                    {

                        tb.Rows.Add();
                        string LineId = ds.Fields.Item("LineId").Value.ToString();
                        tb.SetValue("LineId", tb.Rows.Count - 1, LineId);

                        string Item = ds.Fields.Item("U_Item").Value.ToString();
                        tb.SetValue("U_Item", tb.Rows.Count - 1, Item);

                        string Quantidade = ds.Fields.Item("U_Quantidade").Value.ToString();
                        tb.SetValue("U_Quantidade", tb.Rows.Count - 1, Quantidade);

                        string Deposito = ds.Fields.Item("U_Deposito").Value.ToString();
                        tb.SetValue("U_Deposito", tb.Rows.Count - 1, Deposito);

                        ds.MoveNext();

                    }
                }
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro ao carregar Grid Materiais: {ex.Message}.");
            }
        }
        private void carregarGridTarefas(Form form, string DocEntry = null)
        {
            try
            {

                SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("tbTarefa");
                string consulta = $@"SELECT 
                                    ""U_Descricao"",
                                    ""U_Valor"",
                                    ""U_Tipo""
                                    FROM ""@ACO_SERVICO_2"" WHERE ""DocEntry"" = '{DocEntry}'";
                tb.ExecuteQuery(consulta);


                //Formatar nome colunas
                Grid grd = form.Items.Item("grdTarefa").Specific;
                grd.Columns.Item("U_Descricao").TitleObject.Caption = "Descrição";
                grd.Columns.Item("U_Valor").TitleObject.Caption = "Valor";
                grd.Columns.Item("U_Tipo").TitleObject.Caption = "Tipo";

                //Valores ComboBox
                grd.Columns.Item("U_Tipo").Type = BoGridColumnType.gct_ComboBox;
                ComboBoxColumn oComboColumn = (ComboBoxColumn)grd.Columns.Item("U_Tipo");
                oComboColumn.DisplayType = BoComboDisplayType.cdt_Description;
                oComboColumn.ValidValues.Add("P", "Planejado");
                oComboColumn.ValidValues.Add("I", "Imprevisto");
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro ao carregar Grid Tarefas: {ex.Message}.");
            }
        }
        private void SalvarGridMateriais(Form form)
        {
            try
            {
                DBDataSource db1 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO_1");
                SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("DT_Mat");

                db1.Clear();

                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    db1.InsertRecord(db1.Size);
                    db1.SetValue("U_Item", i, tb.GetValue("U_Item", i));
                    db1.SetValue("U_Quantidade", i, tb.GetValue("U_Quantidade", i));
                    db1.SetValue("U_Deposito", i, tb.GetValue("U_Deposito", i));
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void SalvarGridTarefas(Form form)
        {
            try
            {

                DBDataSource db2 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO_2");
                SAPbouiCOM.DataTable tb2 = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("tbTarefa");

                db2.Clear();

                for (int i = 0; i < tb2.Rows.Count; i++)
                {
                    db2.InsertRecord(db2.Size);
                    db2.SetValue("U_Descricao", i, tb2.GetValue("U_Descricao", i));
                    db2.SetValue("U_Valor", i, tb2.GetValue("U_Valor", i));
                    db2.SetValue("U_Tipo", i, tb2.GetValue("U_Tipo", i));
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private bool Validacao(Form form)
        {
            Recordset ds = (Recordset)Executar.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                string DocEntry = db0.GetValue("DocEntry", 0);

                string consultaMateriais = $@"SELECT * FROM ""@ACO_SERVICO_1"" WHERE ""DocEntry"" = '{DocEntry}'";
                ds.DoQuery(consultaMateriais);
                if (ds.RecordCount > 0)
                {
                    for (int i = 1; i <= ds.RecordCount; i++)
                    {
                        string Item = ds.Fields.Item("U_Item").Value.ToString();
                        double Quantidade = ds.Fields.Item("U_Quantidade").Value;
                        string Deposito = ds.Fields.Item("U_Deposito").Value.ToString();
                        if (string.IsNullOrEmpty(Item))
                        {
                            Executar.Application.MessageBox($"O campo Item está vazio na linha {i}!");
                            return false;
                        }
                        if (string.IsNullOrEmpty(Deposito))
                        {
                            Executar.Application.MessageBox($"O campo Depósito está vazio na linha {i}!");
                            return false;
                        }
                        if (Quantidade <= 0)
                        {
                            Executar.Application.MessageBox($"A quantidade precisa ser maior que zero na linha {i}!");
                            return false;
                        }
                        ds.MoveNext();
                    }
                }

                string consultaTarefas = $@"SELECT * FROM ""@ACO_SERVICO_2"" WHERE ""DocEntry"" = '{DocEntry}'";
                ds.DoQuery(consultaTarefas);
                double CustoTotal = 0; //Custo 0 quando não ha tarefa
                while (!ds.EoF)
                {
                    CustoTotal += Convert.ToDouble(ds.Fields.Item("U_Valor").Value.ToString());
                    ds.MoveNext();
                }
                if (CustoTotal <= 0)
                {
                    Executar.Application.MessageBox($"O custo das tarefas precisa ser maior que zero!");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao realizar validação: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ds);
            }
        }
        private void EncerramentoDoServico(Form form)
        {
            Recordset ds = (Recordset)Executar.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                string DocEntry = db0.GetValue("DocEntry", 0);

                #region Realizar baixa dos materiais
                string consulta = $@"SELECT * FROM ""@ACO_SERVICO_1"" WHERE ""DocEntry"" = '{DocEntry}'";
                ds.DoQuery(consulta);

                if (ds.RecordCount > 0)
                {
                    Executar.Application.StatusBar.SetText("Realizando Saida de Mercadoria...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                    Documents saida = Executar.Company.GetBusinessObject(BoObjectTypes.oInventoryGenExit);
                    saida.Comments = $"Baseado na Prestação de Serviço nº {DocEntry}";
                    saida.DocDate = DateTime.Today;

                    while (!ds.EoF)
                    {
                        saida.Lines.ItemCode = ds.Fields.Item("U_Item").Value.ToString();
                        saida.Lines.Quantity = ds.Fields.Item("U_Quantidade").Value;
                        saida.Lines.WarehouseCode = ds.Fields.Item("U_Deposito").Value.ToString();
                        saida.Lines.Add();
                        ds.MoveNext();
                    }

                    int resp = saida.Add();
                    string Erro = "";

                    if (resp != 0)
                    {
                        Executar.Company.GetLastError(out resp, out Erro);
                        Executar.Application.StatusBar.SetText("Erro ao realizar Saida de Mercadoria!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        throw new Exception("Erro ao realizar Saida de Mercadoria: " + Erro);
                    }
                    string DocEntrySaida = Executar.Company.GetNewObjectKey();


                    string updateSM = $@"UPDATE ""@ACO_SERVICO""
                                        SET U_SM = '{DocEntrySaida}'
                                        WHERE DocEntry = '{DocEntry}'";
                    ds.DoQuery(updateSM);

                    Executar.Application.StatusBar.SetText("Saida de Mercadoria Relizada com Sucesso.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    Executar.Application.StatusBar.SetText("Não existe materiais para esse serviço.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                #endregion

                #region Criar LCM

                Executar.Application.StatusBar.SetText("Realizando Lançamento Contábil...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                string consultaPN = $@"SELECT * FROM ""@ACO_SERVICO"" WHERE ""DocEntry"" = '{DocEntry}'";
                ds.DoQuery(consultaPN);
                string Pn = ds.Fields.Item("U_Cliente").Value.ToString();

                string ContaCredito = Pn; //PN
                string ContaDebito = "3.01.01.01.08"; //Conta serviço

                string consultaTarefas = $@"SELECT * FROM ""@ACO_SERVICO_2"" WHERE ""DocEntry"" = '{DocEntry}'";
                ds.DoQuery(consultaTarefas);

                double CustoTotal = 0;
                while (!ds.EoF)
                {
                    CustoTotal += Convert.ToDouble(ds.Fields.Item("U_Valor").Value.ToString());

                    ds.MoveNext();
                }

                JournalEntries jEventos = (JournalEntries)Executar.Company.GetBusinessObject(BoObjectTypes.oJournalEntries);
                jEventos.TaxDate = DateTime.Now;
                jEventos.DueDate = DateTime.Now;
                jEventos.ReferenceDate = DateTime.Now;
                jEventos.Memo = $"Lançamento de Custo da Prestação de Serviço {DocEntry}.";


                #region Lançamento de Credito
                jEventos.Lines.ShortName = ContaCredito;
                //jEventos.Lines.AccountCode = ContaCredito;
                jEventos.Lines.Credit = CustoTotal;
                jEventos.Lines.Debit = 0;
                jEventos.Lines.TaxDate = DateTime.Now;
                jEventos.Lines.DueDate = DateTime.Now;
                jEventos.Lines.ReferenceDate1 = DateTime.Now;
                jEventos.Lines.Add();
                #endregion

                #region Lançamento de Debito
                jEventos.Lines.ShortName = ContaDebito;
                //jEventos.Lines.AccountCode = ContaDebito;
                jEventos.Lines.Credit = 0;
                jEventos.Lines.Debit = CustoTotal;
                jEventos.Lines.TaxDate = DateTime.Now;
                jEventos.Lines.DueDate = DateTime.Now;
                jEventos.Lines.ReferenceDate1 = DateTime.Now;
                jEventos.Lines.Add();
                #endregion

                int resp2 = jEventos.Add();
                if (resp2 != 0)
                {
                    string Erro = "";
                    Executar.Company.GetLastError(out resp2, out Erro);
                    Executar.Application.StatusBar.SetText("Erro ao realizar Lnaçamento Contábil!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    throw new Exception("Erro ao realizar Lnaçamento Contábil: " + Erro);
                }
                Executar.Application.StatusBar.SetText("Lançamento Contábil Realizado com Sucesso.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                string DocEntryLCM = Executar.Company.GetNewObjectKey();

                string updateLCM = $@"UPDATE ""@ACO_SERVICO""
                                        SET U_LCM = '{DocEntryLCM}'
                                        WHERE DocEntry = '{DocEntry}'";
                ds.DoQuery(updateLCM);


                #endregion

                #region Fechar o serviço
                CompanyService oCompanyService = (CompanyService)Executar.Company.GetCompanyService();
                GeneralService oGeneralService = (GeneralService)oCompanyService.GetGeneralService("ACO_SERVICO");
                GeneralDataParams oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);

                oGeneralParams.SetProperty("DocEntry", DocEntry);
                oGeneralService.Close(oGeneralParams);
                #endregion
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ds);
                GC.Collect();
            }
        }
        private void ModoFormulario(Form form)
        {
            try
            {
                DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                string Status = db0.GetValue("Status", 0);
                string Canceled = db0.GetValue("Canceled", 0);
                if ("C".Equals(Status) || "Y".Equals(Canceled))
                {
                    form.Items.Item("Item_1").Enabled = false;
                    form.Items.Item("Item_3").Enabled = false;
                    form.Items.Item("Item_7").Enabled = false;
                    form.Items.Item("Item_12").Enabled = false;
                    form.Items.Item("Item_21").Enabled = false;
                    form.Items.Item("Item_22").Enabled = false;
                    form.Items.Item("Item_16").Enabled = false;
                    form.Items.Item("grdMat").Enabled = false;
                    form.Items.Item("grdTarefa").Enabled = false;
                }
                else
                {
                    form.Items.Item("Item_1").Enabled = true;
                    form.Items.Item("Item_3").Enabled = true;
                    form.Items.Item("Item_7").Enabled = true;
                    form.Items.Item("Item_12").Enabled = true;
                    form.Items.Item("Item_21").Enabled = true;
                    form.Items.Item("Item_22").Enabled = true;
                    form.Items.Item("Item_16").Enabled = true;
                    form.Items.Item("grdMat").Enabled = true;
                    form.Items.Item("grdTarefa").Enabled = true;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        #region Eventos de Formulario
        public void RightClickEventSv(SAPbouiCOM.ContextMenuInfo info, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                MenuItem oMenuItem;
                Menus oMenus;
                MenuCreationParams oCreationPackage;
                oMenuItem = Executar.Application.Menus.Item("1280");
                oMenus = oMenuItem.SubMenus;
                Form form = Executar.Application.Forms.Item(info.FormUID);

                if (info.BeforeAction)
                {

                    DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                    string Status = db0.GetValue("Status", 0);
                    string Canceled = db0.GetValue("Canceled", 0);
                    if ("O".Equals(Status) && "N".Equals(Canceled))
                    {
                        if (!oMenuItem.SubMenus.Exists("mnEs") && form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            oCreationPackage = (MenuCreationParams)Executar.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "mnEs";
                            oCreationPackage.String = "Encerrar Serviço";
                            oCreationPackage.Position = 1;
                            oCreationPackage.Enabled = true;
                            oMenus.AddEx(oCreationPackage);
                        }

                        if (info.ItemUID == "grdMat" || info.ItemUID == "grdTarefa")
                        {
                            if (!oMenuItem.SubMenus.Exists("mnuAdd"))
                            {
                                oCreationPackage = (MenuCreationParams)Executar.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                                oCreationPackage.UniqueID = "mnuAdd";
                                oCreationPackage.String = "Nova Linha";
                                oCreationPackage.Position = 1;
                                oCreationPackage.Enabled = true;
                                oMenus.AddEx(oCreationPackage);
                            }
                            if (!oMenuItem.SubMenus.Exists("mnuRemove") && info.Row >= 0)
                            {
                                oCreationPackage = (MenuCreationParams)Executar.Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                                oCreationPackage.UniqueID = "mnuRemove";
                                oCreationPackage.String = "Remover Linha";
                                oCreationPackage.Position = 2;
                                oCreationPackage.Enabled = true;
                                oMenus.AddEx(oCreationPackage);
                            }
                        }
                    }
                }
                else
                {
                    if (oMenuItem.SubMenus.Exists("mnuAdd"))
                    {
                        oMenuItem.SubMenus.RemoveEx("mnuAdd");
                    }
                    if (oMenuItem.SubMenus.Exists("mnuRemove"))
                    {
                        oMenuItem.SubMenus.RemoveEx("mnuRemove");
                    }
                    if (oMenuItem.SubMenus.Exists("mnEs"))
                    {
                        oMenuItem.SubMenus.RemoveEx("mnEs");
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void formDataEventSv(SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                Form form = Executar.Application.Forms.Item(pVal.FormUID);
                DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                string DocEntry = db0.GetValue("DocEntry", 0);

                ModoFormulario(form);
                carregarGridMateriais(form, DocEntry);
                carregarGridTarefas(form, DocEntry);

            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro formDataEvent do Form: {ex.Message}.");
            }
        }

        public void itemEventSv(ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "1" && pVal.BeforeAction)
                {
                    Form form = Executar.Application.Forms.Item(pVal.FormUID);
                    SalvarGridMateriais(form);
                    SalvarGridTarefas(form);
                }

                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                {
                    Form form = Executar.Application.Forms.Item(pVal.FormUID);
                    IChooseFromListEvent ecfl = (IChooseFromListEvent)pVal;
                    DBDataSource db0 = (DBDataSource)form.DataSources.DBDataSources.Item("@ACO_SERVICO");
                    SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("DT_Mat");

                    if (ecfl.SelectedObjects != null)
                    {
                        if (pVal.ItemUID == "Item_1")
                        {
                            db0.SetValue("U_Cliente", 0, ecfl.SelectedObjects.GetValue("CardCode", 0).ToString());
                        }

                        if (pVal.ItemUID == "Item_3")
                        {
                            db0.SetValue("U_Funcionario", 0, ecfl.SelectedObjects.GetValue("Code", 0).ToString());
                        }

                        if (pVal.ItemUID == "grdMat")
                        {
                            if (pVal.ColUID == "U_Item")
                            {
                                tb.SetValue("U_Item", pVal.Row, ecfl.SelectedObjects.GetValue("ItemCode", 0).ToString());
                            }

                            if (pVal.ColUID == "U_Deposito")
                            {
                                tb.SetValue("U_Deposito", pVal.Row, ecfl.SelectedObjects.GetValue("WhsCode", 0).ToString());
                            }

                            if (form.Mode != BoFormMode.fm_ADD_MODE)
                            {
                                form.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro itemEvent do Form: {ex.Message}.");
            }
        }

        public void MenuEvents(SAPbouiCOM.MenuEvent menuEvent, string selectedForm, string selectedMatrix, int selectedRow, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (!menuEvent.BeforeAction && menuEvent.MenuUID == "mnEs")
                {
                    Form form = Executar.Application.Forms.Item(selectedForm);
                    if (Executar.Application.MessageBox("Deseja realmente continuar?\nEsse processo é irreversível!", 2, "Sim", "Não") == 1)
                    {
                        if (Validacao(form))
                        {
                            EncerramentoDoServico(form);
                        }
                    }
                }

                if (selectedMatrix == "grdMat")
                {
                    Form form = Executar.Application.Forms.Item(selectedForm);
                    if (!menuEvent.BeforeAction && menuEvent.MenuUID == "mnuAdd")
                    {
                        SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("DT_Mat");
                        tb.Rows.Add();
                    }
                    if (!menuEvent.BeforeAction && menuEvent.MenuUID == "mnuRemove")
                    {
                        SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("DT_Mat");
                        tb.Rows.Remove(selectedRow);
                    }
                }

                if (selectedMatrix == "grdTarefa")
                {
                    Form form = Executar.Application.Forms.Item(selectedForm);
                    if (!menuEvent.BeforeAction && menuEvent.MenuUID == "mnuAdd")
                    {
                        SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("tbTarefa");
                        tb.Rows.Add();

                        tb.SetValue("U_Tipo", tb.Rows.Count - 1, "P");
                    }
                    if (!menuEvent.BeforeAction && menuEvent.MenuUID == "mnuRemove")
                    {
                        SAPbouiCOM.DataTable tb = (SAPbouiCOM.DataTable)form.DataSources.DataTables.Item("tbTarefa");
                        tb.Rows.Remove(selectedRow);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        #endregion

    }
}