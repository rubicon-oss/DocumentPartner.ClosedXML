using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public interface IXLDrawingProtection
    {
        bool Locked { get; set; }
        bool LockText { get; set; }

        IXLDrawingStyle SetLocked(); IXLDrawingStyle SetLocked(bool value);
        IXLDrawingStyle SetLockText(); IXLDrawingStyle SetLockText(bool value);

    }
}
