using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    [ComVisible(false)]
    public static class BinaryReaderExtention
    {
        public class RequireStruct<T> where T : struct { }
        public class RequireClass<T> where T : class { }


        public unsafe static T DynamicCast<T>(this IntPtr pointer)
        {
            return __refvalue(__makeref(pointer), T);
        }

        /// <summary>
        /// Быстрое чтение массива байт в структуру
        /// </summary>
        /// <param name="b">Массив байт</param>
        /// <param name="offset">Смещение начала струкутуры от начала массива байт</param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        static unsafe T ByteArrayToStructure<T>(byte[] b, int offset = 0) where T : struct
        {
            fixed (byte* pb = &b[offset])
            {
                //T val = new T();
                //TypedReference tr = __makeref(val);
                //// Первое поле - указатель в структуре TypedReference - это 
                //// адрес объекта, поэтому мы записываем в него 
                //// указатель на нужный элемент в массиве с данными
                //*(IntPtr*)&tr = (IntPtr)pb;
                //// __refvalue копирует указатель из TypedReference в 'value'
                //val = __refvalue(tr, T);
                //return val;
                return (T)Marshal.PtrToStructure((IntPtr)pb, typeof(T));
            }
        }

        /// <summary>
        /// Чтение структуры из массива байт
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="br">BinaryReader</param>
        /// <param name="parent">Родительский объект</param>
        /// <param name="ignore">типизатор для синтаксис-контроля. Нужен чтобы этото метод использоватлся для чтения структур</param>
        /// <returns></returns>
        public static T Read<T>(this BinaryReader br, object parent = null, RequireStruct<T> ignore = null) where T : struct
        {
            var Length = Marshal.SizeOf(typeof(T));

            byte[] bytes = br.ReadBytes(Length);
            
            return ByteArrayToStructure<T>(bytes);

        }
        
        /// <summary>
        /// Чтение класса из массива байт
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="br">BinaryReader</param>
        /// <param name="parent">Родительский объект</param>
        /// <param name="ignore">типизатор для синтаксис-контроля. Нужен чтобы этото метод использоватлся для чтения класса</param>
        /// <returns></returns>
        public static T Read<T>(this BinaryReader br, object parent = null, RequireClass<T> ignore = null) where T : class
        {
            if(parent == null)
                return (T)Activator.CreateInstance(typeof(T), br);
            else
                return (T)Activator.CreateInstance(typeof(T), br, parent);
        }

        /// <summary>
        /// Чтение словаря структур из массива байт
        /// </summary>
        /// <typeparam name="T">Тип структуры</typeparam>
        /// <param name="br">BinaryReader</param>
        /// <param name="parent">Родительский объект</param>
        /// <param name="ignore">типизатор для синтаксис-контроля. Нужен чтобы этото метод использоватлся для чтения структур</param>
        /// <returns></returns>
        public static Dictionary<int, T> ReadDictionary<T>(this BinaryReader br, object parent = null,RequireStruct<T> ignore = null) where T : struct
        {
            Dictionary<int, T> result = new Dictionary<int, T>();
            int[] numbers = br.ReadIntArray();
            int length = br.ReadCount();
            foreach (int num in numbers)
                result.Add(num, br.Read<T>());

            numbers = null;
            return result;
        }

        /// <summary>
        /// Чтение словаря классов из массива байт
        /// </summary>
        /// <typeparam name="T">Тип структуры</typeparam>
        /// <param name="br">BinaryReader</param>
        /// <param name="parent">Родительский объект</param>
        /// <param name="ignore">типизатор для синтаксис-контроля. Нужен чтобы этото метод использоватлся для чтения класса</param>
        /// <returns></returns>
        public static Dictionary<int, T> ReadDictionary<T>(this BinaryReader br, object parent = null, RequireClass<T> ignore = null) where T : class
        {
            Dictionary<int, T> result = new Dictionary<int, T>();
            int[] numbers = br.ReadIntArray();
            int length = br.ReadCount();
            foreach (int num in numbers)
            {
                if (parent == null)
                    result.Add(num, br.Read<T>());
                else
                    result.Add(num, br.Read<T>(parent));
            }

            numbers = null;
            return result;
        }

        /// <summary>
        /// Чтение массива структур, сериализованного в формате MFC
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="br"></param>
        /// <param name="parent"></param>
        /// <param name="ignore"></param>
        /// <returns></returns>
        public static List<T> ReadList<T>(this BinaryReader br, object parent = null, RequireStruct<T> ignore = null) where T : struct
        {
            List<T> result = new List<T>();
            int length = br.ReadCount();
            for (int num = 0; num < length; num++)
            {
                if(parent == null)
                    result.Add(br.Read<T>());
                else
                    result.Add(br.Read<T>(parent));
            }
            return result;
        }

        /// <summary>
        /// /// Чтение массива классов, сериализованного в формате MFC
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="br"></param>
        /// <param name="parent"></param>
        /// <param name="ignore"></param>
        /// <returns></returns>
        public static List<T> ReadList<T>(this BinaryReader br, object parent = null, RequireClass<T> ignore = null) where T : class
        {
            List<T> result = new List<T>();
            int length = br.ReadCount();
            for (int num = 0; num < length; num++)
            {
                if (parent == null)
                    result.Add(br.Read<T>());
                else
                    result.Add(br.Read<T>(parent));
            }
            return result;
        }

        /// <summary>
        /// Чтение строки, сериализованной в формате MFC
        /// </summary>
        /// <param name="br"></param>
        /// <returns></returns>
        public static string ReadCString(this BinaryReader br)
        {
            int stringLength = br.ReadByte();
            if (stringLength == 0xFF)
                stringLength = br.ReadUInt16();
            if (stringLength == 0xFFFF)
                stringLength = br.ReadInt32();

            return Encoding.GetEncoding(1251).GetString(br.ReadBytes(stringLength));
        }

        /// <summary>
        /// Чтение размера массива в формате MFC
        /// </summary>
        /// <param name="br"></param>
        /// <returns></returns>
        public static int ReadCount(this BinaryReader br)
        {
            int Count = br.ReadUInt16();

            if (Count > 65534)
                Count = br.ReadInt32();

            return Count;
        }

        /// <summary>
        /// Чтение массива Int32, сериализованного в формате MFC ()
        /// </summary>
        /// <param name="br"></param>
        /// <returns></returns>
        public static int[] ReadIntArray(this BinaryReader br)
        {

            int[] result = new int[br.ReadCount()];

            for (int i = 0; i < result.Length; i++)
                try
                {
                    result[i] = br.ReadInt32();
                }
                catch(Exception ex)
                {
                    System.Diagnostics.Debugger.Break();
                }

            return result;
        }
    }

}
