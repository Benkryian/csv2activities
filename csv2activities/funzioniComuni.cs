using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace csv2activities
{
    public static class funzioniComuni
    {

        public static void fetchCSV(string file,char splitter){

            //se uso "@"+file mi fa un escape dei caratteri come la barra. Non mi serve perchè file viene già generato con la dopiia \ per l'escape
            using (var reader = new StreamReader(file))
            {
                List<string> listA = new List<string>();
                List<string> listB = new List<string>();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(splitter);

                    //listA.Add(values[0]);
                    //listB.Add(values[1]);

                }
            }

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
