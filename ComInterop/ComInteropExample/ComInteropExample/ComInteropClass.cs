using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;



namespace ComInteropExample
{
    [Guid("BF81A5B3-A923-46DF-98FF-EEE33A632A77")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ComInteropClass")]
    public class ComInteropClass:iInterface
    {

        public int PerformAddition(int a, int b)
        {
            try
            {
                return a + b;
            }

            catch
            {
                return 0;
            }
        }

        public int PerformDeletion(int a, int b)
        {

            try
            {
                return a - b;
            }
            catch
            {
                return 0;
            }
        }

    }
}
