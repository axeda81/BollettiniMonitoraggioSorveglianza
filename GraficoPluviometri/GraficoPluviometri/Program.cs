using System;

namespace BollettiniMonitoraggio
{
    class Program
    {
        static void Main(string[] args)
        {
            new All2(args[0], args[1], args[2]);
            //new All2("allegato2.csv", "07/06/2016", "03:00");
        }
    }
}
