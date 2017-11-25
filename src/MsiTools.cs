using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32.Interop;


namespace MsiTools
{
    /// <summary>
    /// Properties recognized by MsiGetProductInfo WIN API.
    /// </summary>
    public sealed class InstallProperty
    {
        private readonly String name;
        private readonly int value;

        private static int count = 0;

        public static readonly InstallProperty INSTALLPROPERTY_PACKAGENAME = new InstallProperty("PackageName");
        public static readonly InstallProperty INSTALLPROPERTY_TRANSFORMS = new InstallProperty("Transforms");
        public static readonly InstallProperty INSTALLPROPERTY_LANGUAGE = new InstallProperty("Language");
        public static readonly InstallProperty INSTALLPROPERTY_PRODUCTNAME = new InstallProperty("ProductName");
        public static readonly InstallProperty INSTALLPROPERTY_ASSIGNMENTTYPE = new InstallProperty("AssignmentType");
        public static readonly InstallProperty INSTALLPROPERTY_INSTANCETYPE = new InstallProperty("InstanceType");
        public static readonly InstallProperty INSTALLPROPERTY_AUTHORIZED_LUA_APP = new InstallProperty("AuthorizedLUAApp");
        public static readonly InstallProperty INSTALLPROPERTY_PACKAGECODE = new InstallProperty("PackageCode");
        public static readonly InstallProperty INSTALLPROPERTY_VERSION = new InstallProperty("Version");
        public static readonly InstallProperty INSTALLPROPERTY_PRODUCTICON = new InstallProperty("ProductIcon");
        public static readonly InstallProperty INSTALLPROPERTY_INSTALLEDPRODUCTNAME = new InstallProperty("InstalledProductName");
        public static readonly InstallProperty INSTALLPROPERTY_VERSIONSTRING = new InstallProperty("VersionString");
        public static readonly InstallProperty INSTALLPROPERTY_HELPLINK = new InstallProperty("HelpLink");
        public static readonly InstallProperty INSTALLPROPERTY_HELPTELEPHONE = new InstallProperty("HelpTelephone");
        public static readonly InstallProperty INSTALLPROPERTY_INSTALLLOCATION = new InstallProperty("InstallLocation");
        public static readonly InstallProperty INSTALLPROPERTY_INSTALLSOURCE = new InstallProperty("InstallSource");
        public static readonly InstallProperty INSTALLPROPERTY_INSTALLDATE = new InstallProperty("InstallDate");
        public static readonly InstallProperty INSTALLPROPERTY_PUBLISHER = new InstallProperty("Publisher");
        public static readonly InstallProperty INSTALLPROPERTY_LOCALPACKAGE = new InstallProperty("LocalPackage");
        public static readonly InstallProperty INSTALLPROPERTY_URLINFOABOUT = new InstallProperty("URLInfoAbout");
        public static readonly InstallProperty INSTALLPROPERTY_URLUPDATEINFO = new InstallProperty("URLUpdateInfo");
        public static readonly InstallProperty INSTALLPROPERTY_VERSIONMINOR = new InstallProperty("VersionMinor");
        public static readonly InstallProperty INSTALLPROPERTY_VERSIONMAJOR = new InstallProperty("VersionMajor");
        public static readonly InstallProperty INSTALLPROPERTY_PRODUCTID = new InstallProperty("ProductID");
        public static readonly InstallProperty INSTALLPROPERTY_REGCOMPANY = new InstallProperty("RegCompany");
        public static readonly InstallProperty INSTALLPROPERTY_REGOWNER = new InstallProperty("RegOwner");
        public static readonly InstallProperty INSTALLPROPERTY_UNINSTALLABLE = new InstallProperty("Uninstallable");
        public static readonly InstallProperty INSTALLPROPERTY_PRODUCTSTATE = new InstallProperty("State");
        public static readonly InstallProperty INSTALLPROPERTY_PATCHSTATE = new InstallProperty("State");
        public static readonly InstallProperty INSTALLPROPERTY_PATCHTYPE = new InstallProperty("PatchType");
        public static readonly InstallProperty INSTALLPROPERTY_LUAENABLED = new InstallProperty("LUAEnabled");
        public static readonly InstallProperty INSTALLPROPERTY_DISPLAYNAME = new InstallProperty("DisplayName");
        public static readonly InstallProperty INSTALLPROPERTY_MOREINFOURL = new InstallProperty("MoreInfoURL");
        public static readonly InstallProperty INSTALLPROPERTY_LASTUSEDSOURCE = new InstallProperty("LastUsedSource");
        public static readonly InstallProperty INSTALLPROPERTY_LASTUSEDTYPE = new InstallProperty("LastUsedType");
        public static readonly InstallProperty INSTALLPROPERTY_MEDIAPACKAGEPATH = new InstallProperty("MediaPackagePath");
        public static readonly InstallProperty INSTALLPROPERTY_DISKPROMPT = new InstallProperty("DiskPrompt");

        private InstallProperty(String name)
        {
            this.name = name;
            this.value = count++;
        }

        public override String ToString()
        {
            return name;
        }

        public int Index
        {
            get
            {
                return value;
            }
        }

        public static int Count
        {
            get
            {
                return count;
            }
        }
    }

    /// <summary>
    /// Installed product details.
    /// </summary>
    public class InstalledProduct
    {
        private Guid productCode;
        private string[] productInfo;

        public Guid ProductCode
        {
            get
            {
                return productCode;
            }
        }

        public InstalledProduct()
        {
        }

        public InstalledProduct(Guid productCode)
        {
            this.productCode = productCode;
        }

        public string GetLastUsedSourcePath()
        {
            string source = GetProductInfo(InstallProperty.INSTALLPROPERTY_LASTUSEDSOURCE);
            string msiName = GetProductInfo(InstallProperty.INSTALLPROPERTY_PACKAGENAME);

            if (msiName == String.Empty || source == String.Empty)
            {
                return null;
            }

            int substrBegin = source.LastIndexOf(';');
            if (substrBegin == -1)
            {
                return null;
            }

            string reply = Path.Combine(source.Substring(substrBegin + 1), msiName);
            return reply;
        }

        public string GetProductInfo(InstallProperty property)
        {
            if (productInfo == null)
            {
                productInfo = new string[InstallProperty.Count];
            }
            
            if (productInfo[property.Index] == null)
            {
                productInfo[property.Index] = GetProductInfo(property.ToString());
            }

            return productInfo[property.Index];
        }
        
        /// <summary>
        /// High level Wrapper for MsiGetProductInfo WIN API
        /// </summary>
        /// <param name="property">Install property name. See InstallProperty</param>
        /// <param name="buf">Intermediate buffer where to store property value. Can be null</param>
        /// <returns></returns>
        private string GetProductInfo(string property, StringBuilder buf = null)
        {
            if (buf == null) 
            {
                buf = new StringBuilder(10);
            }
            Int32 len = buf.Capacity;
            UInt32 status = MsiGetProductInfo(productCode.ToString("B"), property, buf, ref len);
            if (status == ResultWin32.ERROR_SUCCESS)
            {
                return buf.ToString();
            }

            if (status == ResultWin32.ERROR_MORE_DATA)
            {
                // " + 1" for trailing null character.
                buf.EnsureCapacity(len + 1);
                return GetProductInfo(property, buf);
            }

            Helper.ReportWinAPICallFailure("MsiGetProductInfo(" + productCode + ", " + property + ", " + len + ")", status);
            return null;
        }
        
        [DllImport("msi.dll", CharSet=CharSet.Unicode)]
        static extern UInt32 MsiGetProductInfo(string product, string property, [Out] StringBuilder valueBuf, ref Int32 len); 
    }

    internal static class Helper
    {
        public static void ReportWinAPICallFailure(string funcCall, uint status)
        {
            StringBuilder errMsg = new StringBuilder("Error calling " + funcCall + " WIN API occured. Error: " + status);

            string errName = ResultWin32.GetErrorName((int)status);
            if (errName != String.Empty)
            {
                errMsg.Append(" (" + errName + ")");
            }

            throw new SystemException(errMsg.ToString());
        }
    }

    /// <summary>
    /// Enumerator for installed products.
    /// </summary>
    public class ProductInfoEnum : IEnumerator<InstalledProduct>
    {
        public ProductInfoEnum()
        {
        }

        public InstalledProduct Current
        {
            get
            {
                if (null == currentValue)
                {
                    throw new InvalidOperationException();
                }
                return currentValue;
            }
        }

        object System.Collections.IEnumerator.Current
        {
            get
            {
                return Current;
            }
        }

        void IDisposable.Dispose()
        {
        }

        bool System.Collections.IEnumerator.MoveNext()
        {
            UInt32 status = MsiEnumProducts(position++, curentProductCode);
            switch ((int)status)
            {
                case ResultWin32.ERROR_NO_MORE_ITEMS:
                    return false;

                case ResultWin32.ERROR_SUCCESS:
                    currentValue = new InstalledProduct(Guid.Parse(curentProductCode.ToString()));
                    return true;

                default:
                    Helper.ReportWinAPICallFailure("MsiEnumProducts(" + (position - 1) + ")", status);
                    return false;
            }
        }

        void System.Collections.IEnumerator.Reset()
        {
            position = 0;
        }

        private StringBuilder curentProductCode = new StringBuilder(39 /* GUID string length with terminating '\0' */);
        private UInt32 position = 0;
        private InstalledProduct currentValue;

        [DllImport("msi.dll", CharSet = CharSet.Unicode)]
        private static extern UInt32 MsiEnumProducts(UInt32 iProductIndex, [Out] StringBuilder lpProductBuf);
    }

    public class Products : IEnumerable<InstalledProduct>
    {
        public Products()
        {
        }

        public IEnumerator<InstalledProduct> GetEnumerator()
        {
            return new ProductInfoEnum();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
