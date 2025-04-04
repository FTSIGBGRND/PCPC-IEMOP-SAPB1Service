using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace SAPB1iService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
#if DEBUG
            SAPB1Service myService = new SAPB1Service();
            myService.OnDebug();
            System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite);
#else        
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new SAPB1Service()
            };
            ServiceBase.Run(ServicesToRun);
#endif
        }
    }
}
