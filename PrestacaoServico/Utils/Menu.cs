using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrestacaoServico.Utils
{
    public class Menu
    {
        public SAPbouiCOM.Application Application;
        public SAPbobsCOM.Company Company;

        public Menu()
        {
            Application = Executar.Application;
            Company = Executar.Company;
        }
        public void CriarMenus()
        {
            try
            {
                if (Application.Menus.Exists("mn_PS"))
                    Application.Menus.Item("43520").SubMenus.Remove(Application.Menus.Item("mn_PS"));


                if (!Application.Menus.Exists("mn_PS"))
                {
                    SAPbouiCOM.MenuItem oMenuItem = Application.Menus.Item("43520");
                    SAPbouiCOM.Menus oMenus = oMenuItem.SubMenus;
                    SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                    oCreationPackage.UniqueID = "mn_PS";
                    oCreationPackage.String = "Prestação de Serviço";
                    oCreationPackage.Position = 17;
                    oMenus.AddEx(oCreationPackage);
                }

                if (!Application.Menus.Exists("mn_Sv"))
                {
                    SAPbouiCOM.MenuItem oMenuItem = Application.Menus.Item("mn_PS");
                    SAPbouiCOM.Menus oMenus = oMenuItem.SubMenus;
                    SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "mn_Sv";
                    oCreationPackage.String = "Serviços";
                    oCreationPackage.Position = 1;
                    oMenus.AddEx(oCreationPackage);
                }

                if (!Application.Menus.Exists("mn_Cf"))
                {
                    SAPbouiCOM.MenuItem oMenuItem = Application.Menus.Item("mn_PS");
                    SAPbouiCOM.Menus oMenus = oMenuItem.SubMenus;
                    SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "mn_Cf";
                    oCreationPackage.String = "Configurações";
                    oCreationPackage.Position = 2;
                    oMenus.AddEx(oCreationPackage);
                }


            }
            catch (Exception ex)
            {
                Application.MessageBox($"Erro ao criar Menu: {ex.Message}.");
            }
        }
    }
}
