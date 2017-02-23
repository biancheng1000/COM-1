using System;
using System.Collections.Generic;
using System.EnterpriseServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComPlusExample
{
    [EventTrackingEnabled(true)]
    [Description("Serviced Component")]
    [assembly: ApplicationActivation(ActivationOption.Server)]
    [assembly: ApplicationAccessControl(false)]
    public class ComPlusClass : ServicedComponent, iInterface
    {
        public int PerformAddition(int a, int b)
        {
            //throw new NotImplementedException();
            try
            {
                return a + b;
            }
            catch
            {
                return 0;
            }
            finally
            {
                Dispose();
            }
        }

        public int PerformSubtraction(int a, int b)
        {
            //throw new NotImplementedException();
            try
            {
                return a - b;
            }
            catch
            {
                return 0;
            }
            finally
            {
                Dispose();
            }
        }

    }
}
