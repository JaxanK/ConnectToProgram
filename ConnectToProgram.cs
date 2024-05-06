using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Runtime.Versioning;

using interop.ICApiIronCAD;
using Autodesk.AutoCAD.Interop;
using System.Collections.Generic;

namespace ConnectToProgram
{
    public static class GetActiveCOMConnection
    {
        public static dynamic? GetActiveConnection(string progID)
        {
            try
            {
                return (dynamic) Marshal2.GetActiveObject(progID);
            }
            catch { }
            return null;
        }
        public static dynamic? GetNewConnection(string progID)
        {
            try { 
                System.Type? appType = System.Type.GetTypeFromProgID(progID);
                if(appType == null ) { return null; }
                return (dynamic?) System.Activator.CreateInstance(appType);

            }
            catch { }
            return null;
        }
    }


    /// <summary>
    /// This class is a custom marshal function that will also work for .Net 5. Since the normal .Net 4.72 functions do not exist with the new .Net system.
    /// </summary>
    public static class Marshal2
    {
        //Finally found a solution!! https://stackoverflow.com/questions/58010510/no-definition-found-for-getactiveobject-from-system-runtime-interopservices-mars/65496277#65496277?newreg=1ddb0d5031d847128c4005f75df3de1c
        internal const String OLEAUT32 = "oleaut32.dll";
        internal const String OLE32 = "ole32.dll";

        [System.Security.SecurityCritical]  // auto-generated_required
        public static Object GetActiveObject(String progID)
        {
            Object obj = null;
            Guid clsid;

            // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
            // CLSIDFromProgIDEx doesn't exist.
            try
            {
                CLSIDFromProgIDEx(progID, out clsid);
            }
            //            catch
            catch (Exception)
            {
                CLSIDFromProgID(progID, out clsid);
            }

            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
        [DllImport(OLEAUT32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

    }
}
