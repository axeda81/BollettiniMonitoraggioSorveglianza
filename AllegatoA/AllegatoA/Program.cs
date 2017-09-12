using System;

namespace BollettiniMonitoraggio
{
    class Program
    {
        static void Main(string[] args)
        {
            //new AllA("./allegato1a.csv", "./allegato3a.csv", "001", "01", "15.09.2016", "15.09.2016", "18:00", "17.09.2016", "12:00");
            new AllA(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8]);

        }
    }
}


