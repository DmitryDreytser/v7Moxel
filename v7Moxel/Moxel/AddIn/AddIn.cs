using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using AddIn;
using System.Reflection;
using System.Collections;
using System.Windows.Forms;
using Moxel;
using System.ComponentModel;
using System.IO;
//using Moxel;

namespace Moxel
{
    [ComVisible(false)]
    internal interface IConverter
    {
        void Attach(dynamic Table);
        string Save(string filename, SaveFormat format);
    }


    [ComVisible(true), Guid("bc631c98-2f0b-49b9-b722-b7e223e46059")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Description("Конвертер MOXEL")]
    [ProgId("AddIn.Moxel.Converter")]
    public class Converter : IInitDone, ILanguageExtender, IConverter
    {

        /// <summary>ProgID COM-объекта компоненты</summary>
        string AddInName = "Moxel.Converter";

        /// <summary>Указатель на IDispatch</summary>
        protected object connect1c;

        /// <summary>Вызов событий 1С</summary>
        protected IAsyncEvent asyncEvent;

        /// <summary>Статусная строка 1С</summary>
        protected IStatusLine statusLine;

        /// <summary>Сообщения об ошибках 1С</summary>
        protected IErrorLog errorLog;

        private Type[] allInterfaceTypes;  // Коллекция интерфейсов
        private MethodInfo[] allMethodInfo;  // Коллекция методов
        private PropertyInfo[] allPropertyInfo; // Коллекция свойств

        private Hashtable nameToNumber; // метод - идентификатор
        private Hashtable numberToName; // идентификатор - метод
        private Hashtable numberToParams; // количество параметров метода
        private Hashtable numberToRetVal; // имеет возвращаемое значение (является функцией)
        private Hashtable propertyNameToNumber; // свойство - идентификатор
        private Hashtable propertyNumberToName; // идентификатор - свойство
        private Hashtable numberToMethodInfoIdx; // номер метода - индекс в массиве методов
        private Hashtable propertyNumberToPropertyInfoIdx; // номер свойства - индекс в массиве свойств

        Moxel mxl;

        protected void PostException(Exception ex)
        {
            if (errorLog == null)
                return;

            var info = new System.Runtime.InteropServices.ComTypes.EXCEPINFO
            {
                wCode = 1006,
                bstrDescription = $"{AddInName}: ошибка {ex.GetType()} : {ex.Message}", //  {ex.StackTrace}
                bstrSource = AddInName, 
                scode = 1
            };

            errorLog.AddError("", ref info);
        }


        public void Attach(object Table)
        {
            string tempfile = Path.GetTempFileName();
            File.Delete(tempfile);
            tempfile += ".mxl";
            object[] param = { tempfile, "mxl" };

            var tt = Table.GetType().InvokeMember("Записать", BindingFlags.InvokeMethod, null, Table, param);

            while (Marshal.ReleaseComObject(Table) > 0) { } ;
            Marshal.FinalReleaseComObject(Table);
            
            if (File.Exists(tempfile))
                mxl = new Moxel(tempfile);

            File.Delete(tempfile);


        }

        public string Save(string filename, SaveFormat format)
        {
            mxl.SaveAs(filename, format);
            return filename;
        }


        #region IInitDone
        public HRESULT Init([MarshalAs(UnmanagedType.IDispatch)] object connection)
        {
            connect1c = connection;
            statusLine = (IStatusLine)connection;
            asyncEvent = (IAsyncEvent)connection;
            errorLog = (IErrorLog)connection;
            return HRESULT.S_OK;
        }

        public HRESULT GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = 2000;
            return HRESULT.S_OK;
        }

        public HRESULT Done()
        {
            
            if (connect1c != null)
            {
                while (Marshal.ReleaseComObject(asyncEvent) > 0) { };
                Marshal.FinalReleaseComObject(asyncEvent);
                asyncEvent = null;

                while (Marshal.ReleaseComObject(statusLine) > 0) { };
                Marshal.FinalReleaseComObject(statusLine);
                statusLine = null;

                while (Marshal.ReleaseComObject(errorLog) > 0) { };
                Marshal.FinalReleaseComObject(errorLog);
                statusLine = null;

                while (Marshal.ReleaseComObject(connect1c) > 0) { };
                Marshal.FinalReleaseComObject(connect1c);
                connect1c = null;
            }
            Marshal.CleanupUnusedObjectsInCurrentContext();

            mxl = null;

            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return HRESULT.S_OK;
        }
        #endregion


#region ILAnguageExtender
        public HRESULT CallAsFunc(int methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
                
            }
            catch (Exception e)
            {
                PostException(e.InnerException);
            }
            return HRESULT.S_OK;
        }

        public HRESULT CallAsProc(int methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref dynamic[] pParams)
        {
            try
            {
                allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
                
            }
            catch (Exception e)
            {
                PostException(e.InnerException);
            }
            return HRESULT.S_OK;
        }


        public Int32 FindMethod([MarshalAs(UnmanagedType.BStr)] string methodName)
        {
            if (nameToNumber.ContainsKey(methodName))
                return (Int32)nameToNumber[methodName];
            else
            return -1;
            
        }

        public HRESULT FindProp([MarshalAs(UnmanagedType.BStr)] string propName, ref Int32 propNum)
        {
            if (propertyNameToNumber.ContainsKey(propName))
            {
                propNum = (Int32)propertyNameToNumber[propName];
                return HRESULT.S_OK;
            }

            propNum = -1;
            return HRESULT.S_FALSE;
        }



        public HRESULT GetMethodName(int methodNum, int methodAlias, [MarshalAs(UnmanagedType.BStr)] ref string methodName)
        {
            if (numberToName.ContainsKey(methodNum))
            {
                methodName = (String)numberToName[methodNum];
                return HRESULT.S_OK;
            }
            return HRESULT.S_FALSE;

        }

        public HRESULT GetNMethods(ref Int32 pMethods)
        {
            pMethods = allMethodInfo.Length;
            return HRESULT.S_OK;
        }

        public HRESULT GetNParams(int methodNum, ref int pParams)
        {
            if (numberToParams.ContainsKey(methodNum))
            {
                pParams = (Int32)numberToParams[methodNum];
                return HRESULT.S_OK;
            }

            pParams = -1;
            return HRESULT.S_FALSE;
        }

        public HRESULT GetNProps(ref int props)
        {
            props = (Int32)propertyNameToNumber.Count;
            return HRESULT.S_OK;
        }

        public HRESULT GetParamDefValue(int methodNum, int paramNum, [MarshalAs(UnmanagedType.Struct)] ref object paramDefValue)
        {
            return HRESULT.S_OK;
        }

        public HRESULT GetPropName(int propNum, int propAlias, [MarshalAs(UnmanagedType.BStr)] ref string propName)
        {
            if (propertyNumberToName.ContainsKey(propNum))
            {
                propName = (String)propertyNumberToName[propNum];
                return HRESULT.S_OK;
            }
            return HRESULT.S_FALSE;
        }

        public HRESULT GetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            propVal = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].GetValue(this, null);
            return HRESULT.S_OK;
        }

        public HRESULT HasRetVal(int methodNum, ref bool retValue)
        {
            if (numberToRetVal.ContainsKey(methodNum))
            {
                retValue = (bool)numberToRetVal[methodNum];
                return HRESULT.S_OK;
            }

            return HRESULT.S_FALSE;
        }


        public HRESULT IsPropReadable(int propNum, ref bool propRead)
        {
            propRead = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanRead;
            return HRESULT.S_OK;
        }

        public HRESULT IsPropWritable(int propNum, ref bool propWrite)
        {
            propWrite = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanWrite;
            return HRESULT.S_OK;
        }

        public HRESULT RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref String extensionName)
        {
            try
            {
                Type type = this.GetType();

                allInterfaceTypes = type.GetInterfaces();
                allMethodInfo = type.GetMethods();
                allPropertyInfo = type.GetProperties();

                // Хэш-таблицы с именами методов и свойств компоненты
                nameToNumber = new Hashtable();
                numberToName = new Hashtable();
                numberToParams = new Hashtable();
                numberToRetVal = new Hashtable();
                propertyNameToNumber = new Hashtable();
                propertyNumberToName = new Hashtable();
                numberToMethodInfoIdx = new Hashtable();
                propertyNumberToPropertyInfoIdx = new Hashtable();

                int Identifier = 0;

                foreach (Type interfaceType in allInterfaceTypes)
                {
                    // Интересуют только методы в пользовательских интерфейсах, стандартные пропускаем
                    if ( interfaceType.Name.Equals("IDisposable")
                      || interfaceType.Name.Equals("IManagedObject")
                      || interfaceType.Name.Equals("IRemoteDispatch")
                      || interfaceType.Name.Equals("IServicedComponentInfo")
                      || interfaceType.Name.Equals("IInitDone")
                      || interfaceType.Name.Equals("ILanguageExtender"))
                    {
                        continue;
                    };

                    // Обработка методов интерфейса
                    MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                    foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                    {
                        nameToNumber.Add(interfaceMethodInfo.Name, Identifier);
                        numberToName.Add(Identifier, interfaceMethodInfo.Name);
                        numberToParams.Add(Identifier, interfaceMethodInfo.GetParameters().Length);
                        numberToRetVal.Add(Identifier, (interfaceMethodInfo.ReturnType != typeof(void)));
                        Identifier++;
                    }

                    // Обработка свойств интерфейса
                    PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                    foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                    {
                        propertyNameToNumber.Add(interfacePropertyInfo.Name, Identifier);
                        propertyNumberToName.Add(Identifier, interfacePropertyInfo.Name);
                        Identifier++;
                    }
                }

                // Отображение номера метода на индекс в массиве
                foreach (DictionaryEntry entry in numberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allMethodInfo.Length; ii++)
                    {
                        if (allMethodInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            numberToMethodInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Метод " + entry.Value.ToString() + " не реализован");
                }

                // Отображение номера свойства на индекс в массиве
                foreach (DictionaryEntry entry in propertyNumberToName)
                {
                    bool found = false;
                    for (int ii = 0; ii < allPropertyInfo.Length; ii++)
                    {
                        if (allPropertyInfo[ii].Name.Equals(entry.Value.ToString()))
                        {
                            propertyNumberToPropertyInfoIdx.Add(entry.Key, ii);
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                        throw new COMException("Свойство " + entry.Value.ToString() + " не реализовано");
                }

                // Компонент инициализирован успешно. Возвращаем имя компонента.
                extensionName = AddInName;
            }
            catch (Exception e)
            {
                PostException(e);
            }

            return HRESULT.S_OK;
            
        }

        public HRESULT SetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal, null);
            return HRESULT.S_OK;
        }
        #endregion
    }
}
