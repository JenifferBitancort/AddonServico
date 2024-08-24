using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.CompilerServices.RuntimeHelpers;

namespace PrestacaoServico
{
    class MigrateTabelaCampos
    {

        public static void CriarTabela(string NomeTabela, string Descricao, BoUTBTableType Tipo)
        {
            try
            {
                int ErrCode;
                string ErrMsg;
                SAPbobsCOM.UserTablesMD oUserTable = (SAPbobsCOM.UserTablesMD)Executar.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!oUserTable.GetByKey(NomeTabela))
                {
                    oUserTable.TableName = NomeTabela;
                    oUserTable.TableDescription = Descricao;
                    oUserTable.TableType = Tipo;

                    int RetVal = oUserTable.Add();
                    if (RetVal != 0)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                        Executar.Company.GetLastError(out ErrCode, out ErrMsg);
                        Executar.Application.MessageBox($"Erro ao criar a tabela {NomeTabela} no SAP: {ErrMsg}");
                    }
                }
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro ao criar Tabela: {ex}.");
            }
        }
        public static void CriarCampos(string Tabela, string NomeCampo, string Descricao, BoFieldTypes Tipo, BoFldSubTypes SubTipo, int Tamanho = 0, string ValorPadrao = null, List<ValidValuesMD> fieldValues = null)
        {
            try
            {
                Recordset ds = (Recordset)Executar.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string queryCampo = $@"SELECT *
                                FROM ""{Executar.Company.CompanyDB}""..""CUFD"" 
                                WHERE ""TableID"" = '{Tabela}' AND ""AliasID"" = '{NomeCampo}'";
                ds.DoQuery(queryCampo);

                SAPbobsCOM.UserFieldsMD userFields = null;

                if (ds.RecordCount == 0)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ds); //Limpeza do objeto 
                    GC.Collect();

                    userFields = (SAPbobsCOM.UserFieldsMD)Executar.Company.GetBusinessObject(BoObjectTypes.oUserFields);

                    userFields.TableName = Tabela;
                    userFields.Name = NomeCampo;
                    userFields.Description = Descricao;
                    userFields.Type = Tipo;
                    userFields.Size = Tamanho;
                    userFields.SubType = SubTipo;

                    if (ValorPadrao != null)
                    {
                        userFields.DefaultValue = ValorPadrao;
                    }

                    if (fieldValues != null)
                    {
                        foreach (var field in fieldValues)
                        {
                            userFields.ValidValues.Value = field.Value;
                            userFields.ValidValues.Description = field.Description;
                            userFields.ValidValues.Add();
                        }
                    }

                    int resp = userFields.Add();
                    if (resp != 0)
                    {
                        string err = $"Erro ao tentar criar campo: {Executar.Company.GetLastErrorDescription()}.";
                        Executar.Application.MessageBox(err);
                    }
                }
            }
            catch (Exception ex)
            {
                Executar.Application.MessageBox($"Erro ao criar Campo: {ex.Message}.");
            }

        }

        public class ValidValuesMD
        {
            public string Value { get; set; }
            public string Description { get; set; }
        }

    }
}
