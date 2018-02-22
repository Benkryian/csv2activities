using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace csv2activities
{
    public static class SAPHelper
    {

        public enum tipoServer { hana, mssql2012, mssql2014 }

        public static Company getSocieta(string server, string porta, tipoServer serverType, string utenteDB, string passDB, string utenteSAP, string passSAP, string DB)
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

            if (i == 0)
            {
                Console.WriteLine("Company connessa");
                return myCompany;
            }
            else
            {
                int errorCode = myCompany.GetLastErrorCode();
                string errorDesc = myCompany.GetLastErrorDescription();

                Console.WriteLine("Errore: cod" + errorCode + " Desc" + errorDesc);
                return null;

            }

        }

        public static bool insActFromCSV(Company myCompany, string file, char splitter)
        {

            //se uso "@"+file mi fa un escape dei caratteri come la barra. Non mi serve perchè file viene già generato con la dopiia \ per l'escape
            using (var reader = new StreamReader(file))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(splitter);

                    //listA.Add(values[0]);
                    //listB.Add(values[1]);
                    Contacts i = myCompany.GetBusinessObject(BoObjectTypes.oContacts);
                    //ADD a new Contact

                    try
                    {
                        bool convertible = false;
                        int actType = 0;
                        convertible = int.TryParse(values[3].Replace('"', ' ').Trim(), out actType);

                        bool convertible2 = false;
                        int act = 0;
                        convertible2 = int.TryParse(values[2].Replace('"', ' ').Trim(), out act);

                        DateTime start = new DateTime();
                        DateTime end = new DateTime();
                        start = funzioniComuni.sapToDatetime(values[5].Replace('"', ' ').Trim());
                        end = funzioniComuni.sapToDatetime(values[6].Replace('"', ' ').Trim());

                        i.CardCode = values[0].Replace('"', ' ').Trim();
                        i.Details = values[1].Replace('"', ' ').Trim();
                        i.Activity = (BoActivities)act;
                        i.ActivityType = actType;
                        i.Notes = values[4].Replace('"', ' ').Trim();
                        i.StartTime = start;
                        i.EndTime = end;

                        //gestione custom fields
                        //i.UserFields.Fields.Item("U_MATRICOLA").Value = "2134325236";
                        int p = i.Add();

                        if (p < 0)
                        {
                            return false;
                        }
                    }
                    catch (Exception e)
                    {

                        Console.WriteLine(""+e);
                    }

                }
            }

            return true;

        }

        public static void getActToXML(Company myCompany)
        {
            //metodo precedente all'introduzione del CRM in sap
            ActivitiesService acts = myCompany.GetCompanyService().GetBusinessService(ServiceTypes.ActivitiesService) as ActivitiesService;
            Activity myAct = acts.GetDataInterface(ActivitiesServiceDataInterfaces.asActivity);

            //get dei dati
            string TodayDateQuery = DateTime.Now.ToString("yyyy-MM-dd");
            Recordset RecSet = myCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            RecSet.DoQuery("select * from \"OCLG\",\"DUMMY\" WHERE \"CntctDate\"= TO_TIMESTAMP ('"+TodayDateQuery+" 00:00:00', 'YYYY-MM-DD HH24:MI:SS')");
            

            string pathToXml = Path.GetDirectoryName(Application.ExecutablePath);
            string TodayDate = DateTime.Now.ToString("dd-MM-yyyy");

            RecSet.SaveXML(TodayDate+"_activities.xml");
            MessageBox.Show("File saved in '"+ pathToXml+"' ", "Export to XML", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            /*
            var o = RecSet.RecordCount;
            while (!RecSet.EoF)
            {
                var CardCode = RecSet.Fields.Item("CardCode");
                RecSet.MoveNext();
            }
            */
        }
        //destroy dell'oggetto company
        public static void destroyCom(Company myCompany)
        {
            try
            {
                myCompany.Disconnect();
            }
            catch (Exception e)
            {

                Console.WriteLine("" + e);
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
