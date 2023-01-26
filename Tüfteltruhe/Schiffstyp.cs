using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tüfteltruhe
{
    public class Schiffstyp
    {
        public int id;
        public string schiffsbezeichnung;
        public int grundgeschwindigkeit;
        public int rudersoll;
        public int segelsoll;

        public Schiffstyp(int id, string schiffsbezeichnung, int grundgeschwindigkeit, int rudersoll, int segelsoll)
        {
            this.id = id;
            this.schiffsbezeichnung = schiffsbezeichnung;
            this.grundgeschwindigkeit = grundgeschwindigkeit;
            this.rudersoll = rudersoll;
            this.segelsoll = segelsoll;
        }
    }
}
