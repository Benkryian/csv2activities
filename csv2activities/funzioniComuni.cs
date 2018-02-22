using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace csv2activities
{
    public static class funzioniComuni
    {
       public static DateTime sapToDatetime(string sapHour)
        {
            sapHour = sapHour.PadLeft(4, '0');
            string sapH = sapHour.Substring(0, 2);
            string sapM = sapHour.Substring(2, 2);

            bool convertible = false;
            int Hour = 0;
            convertible = int.TryParse(sapH, out Hour);
            int Minute = 0;
            convertible = int.TryParse(sapM, out Minute);

            DateTime Today = DateTime.Now;
            DateTime dtHour = new DateTime(Today.Year, Today.Month, Today.Day, Hour, Minute, 0);
            return dtHour;
        }
       
    }


    //classe che definisce l'oggetto attività basato sulla tabella OCLG

    public class myActivity
    {

        enum BoActivities
        {
            telefonata, Riunione, Task, Nota, Compagna, Altro
        }

        tipoAttivita activityType;



    }


    //classe che definisce l'oggetto tipo attività. Necessario perchè in attività il tipo è salvato tramite intero che è chiave esterna della tabella tipo che contiene la descrizione.
    //SELECT *  FROM OCLT T0
    public class tipoAttivita
    {

        public int chiave;
        public string descr;

        public tipoAttivita(int id, string descrizione)
        {
            chiave = id;
            descr = descrizione;

        }

    }


}
