using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace MacolaSynch
{
    class Program
    {

        static void Main(string[] args)
        {
            bool debugMode = false;

            foreach (Object iterator in args)
            {
                if (iterator.ToString().Equals("debug")) 
                {
                    debugMode = true;
                }
            }

            Application app = new Application(debugMode);
        }
        
    }
}
