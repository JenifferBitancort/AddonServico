using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestacaoServico
{
    class Eventos
    {
        formServico Sv;

        public SAPbouiCOM.Application Application;
        public SAPbobsCOM.Company Company;
        private string selectedItem = "";
        private int selectedRow = -1;
        private string selectedForm = "";

        public Eventos()
        {
            try
            {
                Application = Executar.Application;
                Company = Executar.Company;


                Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
                Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
                Application.PrintEvent += new SAPbouiCOM._IApplicationEvents_PrintEventEventHandler(SBO_Application_PrintEvent);

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


        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent menuEvent, out bool BubbleEvent)

        {
            BubbleEvent = true;
            string FormTyp = Executar.Application.Forms.ActiveForm.TypeEx;

            try
            {
                if (!menuEvent.BeforeAction)
                {
                    switch (menuEvent.MenuUID)
                    {

                        case "mn_Sv":
                            Sv = new formServico();
                            Sv.ShowForm();
                            break;

                        case "mn_Cf":
                            //est = new formEst();
                            //est.ShowForm();
                            break;
                    }
                }

                switch (FormTyp)
                {
                    case "Sv":
                        Sv.MenuEvents(menuEvent, selectedForm, selectedItem, selectedRow, out BubbleEvent);
                        break;

                }

            }
            catch (Exception e)
            {
                Application.StatusBar.SetText($"Erro UIEvents: " + e.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {

                if (pVal.FormTypeEx == "Sv")
                {
                    Sv.itemEventSv(pVal, out BubbleEvent);
                }

                if (pVal.FormTypeEx == "est")
                {
                    //est.itemEventEstoq(pVal, out BubbleEvent);
                }

            }
            catch (Exception)
            {

            }

        }


        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo info, out bool BubbleEvent)
        {

            BubbleEvent = true;
            string FormTyp = Executar.Application.Forms.Item(info.FormUID).TypeEx;

            selectedRow = info.Row;
            selectedItem = info.ItemUID;
            selectedForm = info.FormUID;

            if (FormTyp == "Sv")
            {
                Sv.RightClickEventSv(info, out BubbleEvent);
            }
        }


        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

        }


        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx == "Sv" && !pVal.BeforeAction)
            {
                Sv.formDataEventSv(pVal, out BubbleEvent);
            }

        }


        private void SBO_Application_PrintEvent(ref SAPbouiCOM.PrintEventInfo info, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }


    }
}
