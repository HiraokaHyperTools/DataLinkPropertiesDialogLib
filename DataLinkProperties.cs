using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace DataLinkPropertiesDialogLib
{
    public class DataLinkProperties
    {
        public static string PromptNew()
        {
            var dataLinks = new DataLinksClass();
            try
            {
                var connection = dataLinks.PromptNew() as _Connection;
                if (connection != null)
                {
                    try
                    {
                        return connection.ConnectionString;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(connection);
                    }
                }
                return null;
            }
            finally
            {
                Marshal.ReleaseComObject(dataLinks);
            }
        }

        public static string PromptEdit(string connectionString)
        {
            var inputConnection = new ConnectionClass();
            try
            {
                inputConnection.ConnectionString = connectionString;
                var dataLinks = new DataLinksClass();
                try
                {
                    object connectionRef = inputConnection;
                    if (dataLinks.PromptEdit(ref connectionRef))
                    {
                        _Connection outputConnection = (_Connection)connectionRef;
                        try
                        {
                            return outputConnection.ConnectionString;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(connectionRef);
                        }
                    }
                    return null;
                }
                finally
                {
                    Marshal.ReleaseComObject(dataLinks);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(inputConnection);
            }
        }

        [ClassInterface(ClassInterfaceType.None), Guid("2206CDB2-19C1-11D1-89E0-00C04FD7A829")]
        [ComImport]
        public class DataLinksClass : IDataSourceLocator
        {
            [DispId(1610743808)]
            public virtual extern int hWnd
            {
                [DispId(1610743808)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                get;
                [DispId(1610743808)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                set;
            }

            [DispId(1610743810)]
            [MethodImpl(MethodImplOptions.InternalCall)]
            [return: MarshalAs(UnmanagedType.IDispatch)]
            public virtual extern object PromptNew();

            [DispId(1610743811)]
            [MethodImpl(MethodImplOptions.InternalCall)]
            public virtual extern bool PromptEdit([MarshalAs(UnmanagedType.IDispatch)] [In] [Out] ref object ppADOConnection);
        }

        [Guid("2206CCB2-19C1-11D1-89E0-00C04FD7A829")]
        [ComImport]
        public interface IDataSourceLocator
        {
            [DispId(1610743808)]
            int hWnd
            {
                [DispId(1610743808)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                get;
                [DispId(1610743808)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                set;
            }

            [DispId(1610743810)]
            [MethodImpl(MethodImplOptions.InternalCall)]
            [return: MarshalAs(UnmanagedType.IDispatch)]
            object PromptNew();

            [DispId(1610743811)]
            [MethodImpl(MethodImplOptions.InternalCall)]
            bool PromptEdit([MarshalAs(UnmanagedType.IDispatch)] [In] [Out] ref object ppADOConnection);
        }

        [DefaultMember("ConnectionString"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces("ADODB.ConnectionEvents\0"), Guid("00000514-0000-0010-8000-00AA006D2EA4")]
        [ComImport]
        public class ConnectionClass : _Connection
        {
            [DispId(500)]
            public virtual extern object Properties
            {
                [DispId(500)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                [return: MarshalAs(UnmanagedType.Interface)]
                get;
            }

            [DispId(0)]
            public virtual extern string ConnectionString
            {
                [DispId(0)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                [return: MarshalAs(UnmanagedType.BStr)]
                get;
                [DispId(0)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                [param: MarshalAs(UnmanagedType.BStr)]
                set;
            }
        }

        [DefaultMember("ConnectionString"), Guid("00000550-0000-0010-8000-00AA006D2EA4"), TypeLibType(TypeLibTypeFlags.FDual | TypeLibTypeFlags.FDispatchable)]
        [ComImport]
        public interface _Connection
        {
            [DispId(500)]
            object Properties
            {
                [DispId(500)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                [return: MarshalAs(UnmanagedType.Interface)]
                get;
            }

            [DispId(0)]
            string ConnectionString
            {
                [DispId(0)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                [return: MarshalAs(UnmanagedType.BStr)]
                get;
                [DispId(0)]
                [MethodImpl(MethodImplOptions.InternalCall)]
                [param: MarshalAs(UnmanagedType.BStr)]
                set;
            }
        }
    }
}
