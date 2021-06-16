using System;
using System.Runtime.InteropServices;
using AddIn;
using System.Reflection;
using System.Collections;
using System.Threading.Tasks;
using System.Windows.Forms;
using v7Moxel.Moxel.ExcelWriter;

namespace Moxel
{
    public sealed class AliasAttribute : Attribute
    {
        public string RussianName { get; set; }
        public AliasAttribute(string alias)
        {
            RussianName = alias;
        }
    }

    public abstract class AddIn : IInitDone, ILanguageExtender
    {
        /// <summary>ProgID COM-объекта компоненты</summary>
        string AddInName
        {
            get
            {
                return this.GetType().GetCustomAttribute<ProgIdAttribute>().Value.Replace("AddIn.", "");
            }
        }
        

        /// <summary>Указатель на IDispatch</summary>
        static protected dynamic connect1c;

        /// <summary>Вызов событий 1С</summary>
        static protected IAsyncEvent asyncEvent;

        /// <summary>Статусная строка 1С</summary>
        static protected IStatusLine statusLine;

        /// <summary>Сообщения об ошибках 1С</summary>
        static protected IErrorLog errorLog;
        /// <summary>
        /// Коллекция интерфейсов
        /// </summary>
        private Type[] allInterfaceTypes;  
        /// <summary>
        /// Коллекция методов
        /// </summary>
        private MethodInfo[] allMethodInfo; 
        /// <summary>
        /// Коллекция свойств
        /// </summary>
        private PropertyInfo[] allPropertyInfo;
        /// <summary>
        /// метод - идентификатор
        /// </summary>
        private Hashtable nameToNumber;
        /// <summary>
        /// идентификатор - метод
        /// </summary>
        private Hashtable numberToName;  
        /// <summary>
        /// количество параметров метода
        /// </summary>
        private Hashtable numberToParams;
        /// <summary>
        /// имеет возвращаемое значение (является функцией)
        /// </summary>
        private Hashtable numberToRetVal;
        /// <summary>
        /// свойство - идентификатор
        /// </summary>
        private Hashtable propertyNameToNumber;
        /// <summary>
        /// идентификатор - свойство
        /// </summary>
        private Hashtable propertyNumberToName;
        /// <summary>
        /// номер метода - индекс в массиве методов
        /// </summary>
        private Hashtable numberToMethodInfoIdx;
        /// <summary>
        /// номер свойства - индекс в массиве свойств
        /// </summary>
        private Hashtable propertyNumberToPropertyInfoIdx; 

        public delegate IntPtr dGetMainFrame();
        
        /// <summary>
        /// GetMainFrame
        /// 
        /// </summary>
        public static dGetMainFrame GetMainFrame = WinApi.GetDelegate<dGetMainFrame>("frame.dll", "?GetMainFrame@@YAPAVCMDIFrameWnd@@XZ");
       
        /// <summary>
        /// AfxGetMainWnd
        /// </summary>
        public static dGetMainFrame GetMainWindow = MFCNative.GetDelegate<dGetMainFrame>(6575);


        protected string ErrorDescription = null;
        protected string ErrorStackTrace = null;

        protected abstract HRESULT OnRegister();
        protected abstract void OnInit();
        protected abstract void OnDone();

        static int refCount = 0;
        /// <summary>
        /// Получает главное окно 1С
        /// </summary>
        /// <returns>IntPtr hWnd</returns>
        public static IntPtr Get1CWindow()
        {
            if (GetMainFrame == null)
                return IntPtr.Zero;

            return Marshal.ReadIntPtr(GetMainFrame(), 32);
        }

        /// <summary>
        /// Возвращает MDI контейнер 1С
        /// </summary>
        /// <returns></returns>
        public static IntPtr Get1CMDIWindow()
        {
            if (GetMainFrame == null)
                return IntPtr.Zero;

            return Marshal.ReadIntPtr(GetMainFrame(), 192);
        }

        static IntPtr MainWindow = IntPtr.Zero;

        /// <summary>
        /// Устанавливает строку статуса. 
        /// </summary>
        /// <param name="message"></param>
        public static void StatusLine(string message)
        {
            if (MainWindow == IntPtr.Zero)
                MainWindow = Get1CWindow();

            if (MainWindow != IntPtr.Zero)
                Task.Run(() => WinApi.SendMessage(Get1CWindow(), 866, 0, message));
        }

        /// <summary>
        /// Событие изменения прогресса
        /// </summary>
        /// <param name="progress"></param>
        private static void Writer_onProgress(int progress)
        {
            StatusLine($"{progress:D2}%");
            Application.DoEvents();
        }

        public void PostException(Exception ex)
        {

            if (connect1c == null)
                return;

            object[] param = { 1006, AddInName, ex.Message, 1};

            var tt = connect1c.GetType().InvokeMember("AddError", BindingFlags.InvokeMethod, null, connect1c, param);
        }

        #region IInitDone
        HRESULT IInitDone.Init([MarshalAs(UnmanagedType.IDispatch)] dynamic connection)
        {
            try
            {
                connect1c = connection;
                if (statusLine == null)
                {
                    statusLine = (IStatusLine)connection;
                    ExcelWriter.OnProgress += Writer_onProgress;
                    HtmlWriter.onProgress += Writer_onProgress;
                    PDFWriter.onProgress += Writer_onProgress;
                }

                asyncEvent = (IAsyncEvent)connection;

                if(errorLog == null)
                    errorLog = (IErrorLog)connection;

                OnInit();
                refCount++;
                return HRESULT.S_OK;
            }
            catch
            {
                return HRESULT.E_FAIL;
            }

        }

        HRESULT IInitDone.GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = 2000;
            return HRESULT.S_OK;
        }

        HRESULT IInitDone.Done()
        {

            OnDone();

            if (--refCount == 0)
            {
                try
                {
                    if (statusLine != null)
                    {
                        //while (Marshal.ReleaseComObject(statusLine) > 0) { };
                        Marshal.FinalReleaseComObject(statusLine);
                        statusLine = null;
                    }

                    if (errorLog != null)
                    {
                        //while (Marshal.ReleaseComObject(errorLog) > 0) { };
                        Marshal.FinalReleaseComObject(errorLog);
                        errorLog = null;
                    }

                    if (asyncEvent != null)
                    {
                        //while (Marshal.ReleaseComObject(asyncEvent) > 0) { };
                        Marshal.FinalReleaseComObject(asyncEvent);
                        asyncEvent = null;
                    }
                    if (connect1c != null)
                    {
                        //while (Marshal.ReleaseComObject(connect1c) > 0) { };
                        Marshal.FinalReleaseComObject(connect1c);
                        connect1c = null;
                    }
                }
                catch { }
            }

            GC.Collect();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.CleanupUnusedObjectsInCurrentContext();
            return HRESULT.S_OK;
        }
        #endregion

        #region ILAnguageExtender
        HRESULT ILanguageExtender.CallAsFunc(int methodNum, [MarshalAs(UnmanagedType.Struct)] ref object retValue, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            retValue = "";
            try
            {
                retValue = allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
                return HRESULT.S_OK;
            }

            catch (Exception e)
            {
                while (e.InnerException != null)
                    e = e.InnerException;

                ErrorDescription = e.Message;
                ErrorStackTrace = e.StackTrace;

                Converter.RaiseExtRuntimeError(e.Message);
                pParams = null;
                return HRESULT.E_FAIL;
            }
            
        }

        HRESULT ILanguageExtender.CallAsProc(int methodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {
            try
            {
                allMethodInfo[(int)numberToMethodInfoIdx[methodNum]].Invoke(this, pParams);
            }
            catch (Exception e)
            {
                ErrorDescription = e.InnerException.Message;
                pParams = null;
                return HRESULT.E_FAIL;
            }
            return HRESULT.S_OK;
        }


        HRESULT ILanguageExtender.FindMethod([MarshalAs(UnmanagedType.BStr)] string methodName, ref int methodNUm)
        {
            if (nameToNumber.ContainsKey(methodName.ToUpper()))
            {
                methodNUm = (int)nameToNumber[methodName.ToUpper()];
                return HRESULT.S_OK;
            }

            methodNUm = -1;
            return HRESULT.S_FALSE;
        }

        HRESULT ILanguageExtender.FindProp([MarshalAs(UnmanagedType.BStr)] string propName, ref Int32 propNum)
        {
            if (propertyNameToNumber.ContainsKey(propName.ToUpper()))
            {
                propNum = (Int32)propertyNameToNumber[propName.ToUpper()];
                return HRESULT.S_OK;
            }

            propNum = -1;
            return HRESULT.S_FALSE;
        }



        HRESULT ILanguageExtender.GetMethodName(int methodNum, int methodAlias, [MarshalAs(UnmanagedType.BStr)] ref string methodName)
        {
            if (numberToName.ContainsKey(methodNum))
            {
                methodName = (String)numberToName[methodNum];
                return HRESULT.S_OK;
            }
            return HRESULT.S_FALSE;

        }

        HRESULT ILanguageExtender.GetNMethods(ref Int32 pMethods)
        {
            pMethods = allMethodInfo.Length;
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.GetNParams(int methodNum, ref int pParams)
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
            propVal = allPropertyInfo[propNum].GetValue(this, null);
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.HasRetVal(int methodNum, ref bool retValue)
        {
            if (numberToRetVal.ContainsKey(methodNum))
            {
                retValue = (bool)numberToRetVal[methodNum];
                return HRESULT.S_OK;
            }

            return HRESULT.S_FALSE;
        }


        HRESULT ILanguageExtender.IsPropReadable(int propNum, ref bool propRead)
        {
            propRead = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanRead;
            return HRESULT.S_OK;
        }

        public HRESULT IsPropWritable(int propNum, ref bool propWrite)
        {
            propWrite = allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].CanWrite;
            return HRESULT.S_OK;
        }

        HRESULT ILanguageExtender.RegisterExtensionAs([In, Out, MarshalAs(UnmanagedType.BStr)] ref string extensionName)
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
                      || interfaceType.Name.Equals(nameof(IInitDone))
                      || interfaceType.Name.Equals(nameof(ILanguageExtender)))
                    {
                        continue;
                    };

                    // Обработка методов интерфейса
                    MethodInfo[] interfaceMethods = interfaceType.GetMethods();
                    foreach (MethodInfo interfaceMethodInfo in interfaceMethods)
                    {
                        if (interfaceMethodInfo.IsSpecialName)
                            continue;
                        string alias = ((AliasAttribute)Attribute.GetCustomAttribute(interfaceMethodInfo, typeof(AliasAttribute))).RussianName.ToUpper();

                        nameToNumber.Add(interfaceMethodInfo.Name.ToUpper(), Identifier);
                        numberToName.Add(Identifier, interfaceMethodInfo.Name);
                        numberToParams.Add(Identifier, interfaceMethodInfo.GetParameters().Length);
                        numberToRetVal.Add(Identifier, (interfaceMethodInfo.ReturnType != typeof(void)));
                        if (!string.IsNullOrWhiteSpace(alias))
                        {
                            nameToNumber.Add(alias, Identifier);
                        }
                        Identifier++;
                    }

                    Identifier = 0;
                    // Обработка свойств интерфейса
                    PropertyInfo[] interfaceProperties = interfaceType.GetProperties();
                    foreach (PropertyInfo interfacePropertyInfo in interfaceProperties)
                    {
                        string alias = ((AliasAttribute)Attribute.GetCustomAttribute(interfacePropertyInfo, typeof(AliasAttribute))).RussianName;

                        propertyNameToNumber.Add(interfacePropertyInfo.Name.ToUpper(), Identifier);

                        if (!string.IsNullOrWhiteSpace(alias))
                            propertyNameToNumber.Add(alias.ToUpper(), Identifier);

                        if (!string.IsNullOrWhiteSpace(alias))
                            propertyNumberToName.Add(Identifier, alias);
                        else
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

                        if (propertyNameToNumber.ContainsKey(entry.Value.ToString().ToUpper()))
                        {
                            propertyNumberToPropertyInfoIdx.Add(entry.Key, propertyNameToNumber[entry.Value.ToString().ToUpper()]);
                            found = true;
                            break;
                        };
                    }
                    if (!found && !propertyNameToNumber.ContainsKey(entry.Value.ToString().ToUpper()))
                        throw new COMException("Свойство " + entry.Value.ToString() + " не реализовано");
                }

                // Компонент инициализирован успешно. Возвращаем имя компонента.
                extensionName = AddInName;
            }
            catch (Exception e)
            {
                return HRESULT.S_FALSE;
            }

            return OnRegister();
            
        }

        public HRESULT SetPropVal(int propNum, [MarshalAs(UnmanagedType.Struct)] ref object propVal)
        {
            allPropertyInfo[(int)propertyNumberToPropertyInfoIdx[propNum]].SetValue(this, propVal, null);
            return HRESULT.S_OK;
        }
        #endregion
    }
}
