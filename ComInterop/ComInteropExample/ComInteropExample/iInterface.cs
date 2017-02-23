using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;



namespace ComInteropExample
{
    [Guid("B8E0BED6-72F6-45BF-B114-B5C5EFBA34C4")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface iInterface
    {
        int PerformAddition(int a, int b);
        int PerformDeletion(int a, int b);
    }
}
