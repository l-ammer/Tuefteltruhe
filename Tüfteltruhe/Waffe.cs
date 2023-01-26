using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tüfteltruhe
{
    public class Waffe
    {
        public string name;
        public string waffentyp;
        public int reichweiteNah;
        public int reichweiteFern;
        public double gewicht;
        public int stichmod;
        public int wuchtmod;
        public int schnittmod;
        public int parierenmod;
        public int blockenmod;
        public int ausschaltenmod;
        public int werfenmod;
        public int oeffnenmod;
        public int waffenenergie;
        public string waffenschaden;
        public string wurfschaden;
        public int instanzindex;
        public int schirmwert;

        public Waffe(string name, string waffentyp, double gewicht, int stichmod, int wuchtmod, int schnittmod, int parierenmod, int blockenmod, int ausschaltenmod, int werfenmod, int oeffnenmod, string waffenschaden, string wurfschaden, int instanzindex, int reichweiteNah, int reichweiteFern, int waffenenergie, int schirmwert)
        {
            this.name = name;
            this.waffentyp = waffentyp;
            this.gewicht = gewicht;
            this.stichmod = stichmod;
            this.wuchtmod = wuchtmod;
            this.schnittmod = schnittmod;
            this.parierenmod = parierenmod;
            this.blockenmod = blockenmod;
            this.ausschaltenmod = ausschaltenmod;
            this.werfenmod = werfenmod;
            this.oeffnenmod = oeffnenmod;
            this.waffenschaden = waffenschaden;
            this.wurfschaden = wurfschaden;
            this.instanzindex = instanzindex;
            this.reichweiteNah = reichweiteNah;
            this.reichweiteFern = reichweiteFern;
            this.waffenenergie = waffenenergie;
            this.schirmwert = schirmwert;
        }
    }       
}
