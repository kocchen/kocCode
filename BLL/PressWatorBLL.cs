using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IServer;

namespace BLL
{
    public class PressWatorBLL:IPressWator
    {
        public IWatorTools GetWatorTools()
        {
            return new WatorToolsBLL();
        }
    }
}
