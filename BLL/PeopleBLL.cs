using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IServer;

namespace BLL
{
    public class PeopleBLL:IPeople
    {
        IPressWator p = null;
        public PeopleBLL(IPressWator pw)
        {
            p = pw;
        }
        public void toDoSomething()
        {
            IWatorTools wt = p.GetWatorTools();
            wt.GetWator();
        }
    }
}
