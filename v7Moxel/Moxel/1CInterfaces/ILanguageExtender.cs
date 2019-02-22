using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BMP1C.Net._1CInterfaces
{
    /// <summary>
    /// При вызове функции из 1С производится следующая последовательность вызовов:
    ///    1. FindMethod
    ///    2. GetNParams
    ///    3. HasRetVal
    /// </summary>
    /// 
    [ComVisible(true)]
    [Guid("AB634003-F13D-11d0-A459-004095E1DAEA")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILanguageExtender
    {
        /// <summary>
        /// Регистрация компонента в 1C
        /// </summary>
        /// <param name="extensionName"></param>
        /// <remarks>
        /// <prototype>
        /// [helpstring("method RegisterExtensionAs")]
        /// HRESULT RegisterExtensionAs([in,out]BSTR *bstrExtensionName);
        /// </prototype>
        /// </remarks>
        void RegisterExtensionAs(
          [MarshalAs(UnmanagedType.BStr)]
    ref String extensionName);

        /// <summary>
        /// Возвращается количество свойств
        /// </summary>
        /// <param name="props">Количество свойств </param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetNProps([in,out]long *plProps);
        /// </prototype>
        /// </remarks>
        void GetNProps(ref Int32 props);

        /// <summary>
        /// Возвращает целочисленный идентификатор свойства, соответствующий 
        /// переданному имени
        /// </summary>
        /// <param name="propName">Имя свойства</param>
        /// <param name="propNum">Идентификатор свойства</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT FindProp([in]BSTR bstrPropName,[in,out]long *plPropNum);
        /// </prototype>
        /// </remarks>
        void FindProp(
          [MarshalAs(UnmanagedType.BStr)]
    String propName,
          ref Int32 propNum);

        /// <summary>
        /// Возвращает имя свойства, соответствующее 
        /// переданному целочисленному идентификатору
        /// </summary>
        /// <param name="propNum">Идентификатор свойства</param>
        /// <param name="propAlias"></param>
        /// <param name="propName">Имя свойства</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetPropName([in]long lPropNum,[in]long lPropAlias,[in,out]BSTR *pbstrPropName);
        /// </prototype>
        /// </remarks>
        void GetPropName(
          Int32 propNum,
          Int32 propAlias,
          [MarshalAs(UnmanagedType.BStr)]
    ref String propName);

        /// <summary>
        /// Возвращает значение свойства.
        /// </summary>
        /// <param name="propNum">Идентификатор свойства </param>
        /// <param name="StringpropVal">Значение свойства</param>
        /// <remarks>
        /// <prototype>
        /// </prototype>
        /// </remarks>
        void GetPropVal(
          Int32 propNum,
          [MarshalAs(UnmanagedType.Struct)]
    ref object propVal);

        /// <summary>
        /// Устанавливает значение свойства.
        /// </summary>
        /// <param name="propName">Имя свойства</param>
        /// <param name="propVal">Значение свойства</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT SetPropVal([in]long lPropNum,[in]VARIANT *varPropVal);
        /// </prototype>
        /// </remarks>
        void SetPropVal(
          Int32 propNum,
          [MarshalAs(UnmanagedType.Struct)]
    ref object propVal);

        /// <summary>
        /// Определяет, можно ли читать значение свойства
        /// </summary>
        /// <param name="propNum"> Идентификатор свойства </param>
        /// <param name="propRead">Флаг читаемости</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT IsPropReadable([in]long lPropNum,[in,out]BOOL *pboolPropRead);
        /// </prototype>
        /// </remarks>
        void IsPropReadable(Int32 propNum, ref bool propRead);

        /// <summary>
        /// Определяет, можно ли изменять значение свойства
        /// </summary>
        /// <param name="propNum">Идентификатор свойства</param>
        /// <param name="propRead">Флаг записи</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT IsPropWritable([in]long lPropNum,[in,out]BOOL *pboolPropWrite);
        /// </prototype>
        /// </remarks>
        void IsPropWritable(Int32 propNum, ref Boolean propWrite);

        /// <summary>
        /// Возвращает количество методов
        /// </summary>
        /// <param name="pMethods">Количество методов</param>
        /// <remarks>
        /// <prototype>
        /// [helpstring("method GetNMethods")]
        /// HRESULT GetNMethods([in,out]long *plMethods);
        /// </prototype>
        /// </remarks>
        void GetNMethods(ref Int32 pMethods);

        /// <summary>
        /// Возвращает идентификатор метода по его имени
        /// </summary>
        /// <param name="methodName">Имя метода</param>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT FindMethod(BSTR bstrMethodName,[in,out]long *plMethodNum);
        /// </prototype>
        /// </remarks>
        void FindMethod(
          [MarshalAs(UnmanagedType.BStr)]
    String methodName,
          ref Int32 methodNum);

        /// <summary>
        /// Возвращает имя метода по его идентификатору
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="methodAlias"></param>
        /// <param name="methodName">Имя метода</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetMethodName([in]long lMethodNum,[in]long lMethodAlias,[in,out]BSTR *pbstrMethodName);
        /// </prototype>
        /// </remarks>
        void GetMethodName(Int32 methodNum,
          Int32 methodAlias,
          [MarshalAs(UnmanagedType.BStr)]
    ref String methodName);

        /// <summary>
        /// Возвращает количество параметров метода по его идентификатору
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Количество параметров</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT GetNParams([in]long lMethodNum, [in,out]long *plParams);
        /// </prototype>
        /// </remarks>
        void GetNParams(Int32 methodNum, ref Int32 pParams);

        void GetParamDefValue(
          Int32 methodNum,
          Int32 paramNum,
          [MarshalAs(UnmanagedType.Struct)]
    ref object paramDefValue);

        /// <summary>
        /// Указывает, что у метода есть возвращаемое значение
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Наличие возвращаемого значения</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT HasRetVal([in]long lMethodNum,[in,out]BOOL *pboolRetValue);
        /// </prototype>
        /// </remarks>
        void HasRetVal(Int32 methodNum, ref Boolean retValue);

        /// <summary>
        /// Вызов метода как процедуры с использованием идентификатора
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="pParams">Параметры</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT CallAsProc([in]long lMethodNum,[in] SAFEARRAY (VARIANT) *paParams);
        /// </prototype>
        /// </remarks>
        void CallAsProc(
          Int32 methodNum,
          [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
    ref object[] pParams);

        /// <summary>
        /// Вызов метода как функции с использованием идентификатора
        /// </summary>
        /// <param name="methodNum">Идентификатор метода</param>
        /// <param name="retValue">Возвращаемое значение</param>
        /// <param name="pParams">Параметры</param>
        /// <remarks>
        /// <prototype>
        /// HRESULT CallAsFunc([in]long lMethodNum,[in,out] VARIANT *pvarRetValue,
        ///         [in] SAFEARRAY (VARIANT) *paParams);
        /// </prototype>
        /// </remarks>
        void CallAsFunc(
          Int32 methodNum,
          [MarshalAs(UnmanagedType.Struct)]
    ref object retValue,
          [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)]
    ref object[] pParams);
    }
}
