using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tüfteltruhe
{
    public class Bereich
    {
        public int id;
        public string bezeichnung;
        public string[] gewöhnlich;
        public string[] ungewöhnlich;
        public string[] selten;
        public string[] sehrselten;

        public Bereich(int id, string bezeichnung, string[] gewöhnlich, string[] ungewöhnlich, string[] selten, string[] sehrselten)
        {
            this.id = id;
            this.bezeichnung = bezeichnung;
            this.gewöhnlich = gewöhnlich;
            this.ungewöhnlich = ungewöhnlich;
            this.selten = selten;
            this.sehrselten = sehrselten;
        }
    }
}
