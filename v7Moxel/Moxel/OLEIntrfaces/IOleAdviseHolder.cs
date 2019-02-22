using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Linq;

namespace Ole
{
    [Guid("00000111-0000-0000-C000-000000000046")]
    [InterfaceType(1)]
    public interface IOleAdviseHolder
    {
        void Advise(IAdviseSink pAdvise, out uint pdwConnection);
        void EnumAdvise(out IEnumStatData ppenumAdvise);
        void SendOnClose();
        void SendOnRename(IMoniker pmk);
        void SendOnSave(IMoniker pmk);
        void Unadvise(uint dwConnection);
    }
}
