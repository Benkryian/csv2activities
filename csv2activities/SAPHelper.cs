using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace csv2activities
{
    public static class SAPHelper
    {

        public enum tipoServer { hana,mssql2012,mssql2014}

        public static Company getSocieta(string server , string porta, tipoServer serverType, string utenteDB, string passDB, string utenteSAP, string passSAP, string DB )
        {

            Company myCompany = new Company();
            myCompany.Server = $"{server}:{porta}";

            switch (serverType)
            {
                case tipoServer.hana:
                    myCompany.DbServerType = BoDataServerTypes.dst_HANADB;
                    break;
                case tipoServer.mssql2012:
                    myCompany.DbServerType = BoDataServerTypes.dst_MSSQL2012;
                    break;
                case tipoServer.mssql2014:
                    myCompany.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                    break;
                default:
                    break;
            }

            myCompany.DbUserName = utenteDB;
            myCompany.DbPassword = passDB;
            myCompany.UserName = utenteSAP;
            myCompany.Password = passSAP;
            myCompany.CompanyDB = DB;

            //la connessione restituisce 0 se ha successo e un numero negativo se da errore. Il numero è il codice di errore.
            int i = myCompany.Connect();

            if (i == 0) {
                Console.WriteLine("Company connessa");
                return myCompany;
            }
            else
            {
                int errorCode = myCompany.GetLastErrorCode();
                string errorDesc = myCompany.GetLastErrorDescription();

                Console.WriteLine("Errore: cod"+errorCode+" Desc"+errorDesc);
                return null;

            }
        
        }

        public static void getAct(Company myCompany)
        {
            //metodo precedente all'introduzione del CRM in sap
            ActivitiesService acts = myCompany.GetCompanyService().GetBusinessService(ServiceTypes.ActivitiesService) as ActivitiesService;
            Activity myAct = acts.GetDataInterface(ActivitiesServiceDataInterfaces.asActivity);

            //get dei dati
            Recordset RecSet = myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            RecSet.DoQuery("SELECT * FROM OCLG");
            var o = RecSet.RecordCount;
            while (!RecSet.EoF)
            {
                var CardCode = RecSet.Fields.Item("CardCode");
                RecSet.MoveNext();
            }

            //le attività sono di tipo oContacts. Bestemmie ma è così.
            Contacts i = myCompany.GetBusinessObject(BoObjectTypes.oContacts);
            //ADD a new Contact
            i.Notes = "pippo";
            i.UserFields.Fields.Item("U_MATRICOLA").Value = "2134325236";
            int p = i.Add();

            destroyCom(myCompany);

        }
        //destroy dell'oggetto company
        public static void destroyCom(Company myCompany) {
            try
            {
                myCompany.Disconnect();
            }
            catch (Exception e)
            {

                Console.WriteLine(""+e);
            }

            try
            {
                Marshal.ReleaseComObject(myCompany);
                Marshal.FinalReleaseComObject(myCompany);
            }
            catch (Exception e)
            {

                Console.WriteLine("" + e);
            }
         

        }


    }
}
