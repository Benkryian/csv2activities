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
