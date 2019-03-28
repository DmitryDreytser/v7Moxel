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

        /// <summary>
        /// Быстрое чтение массива байт в структуру
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="b"></param>
        /// <returns></returns>
        static unsafe T ByteArrayToStructure<T>(byte[] b) where T : struct
        {
            fixed (byte* pb = &b[0])
                return (T)Marshal.PtrToStructure((IntPtr)pb, typeof(T));
        }

        public static T Read<T>(this BinaryReader br, object parent = null, RequireStruct<T> ignore = null) where T : struct
        {
            var Length = Marshal.SizeOf(typeof(T));

            byte[] bytes = br.ReadBytes(Length);
            
            return ByteArrayToStructure<T>(bytes);

        }

        public static T Read<T>(this BinaryReader br, object parent = null, RequireClass<T> ignore = null) where T : class
        {
            if(parent == null)
                return (T)Activator.CreateInstance(typeof(T), br);
            else
                return (T)Activator.CreateInstance(typeof(T), br, parent);
        }

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

        public static string ReadCString(this BinaryReader br)
        {
            int stringLength = br.ReadByte();
            if (stringLength == 0xFF)
                stringLength = br.ReadInt16();
            if (stringLength == 0xFFFF)
                stringLength = br.ReadInt32();

            return Encoding.GetEncoding(1251).GetString(br.ReadBytes(stringLength));
        }

        public static int ReadCount(this BinaryReader br)
        {
            int Count = br.ReadInt16();

            if (Count == -1)
                Count = br.ReadInt32();

            return Count;
        }

        public static int[] ReadIntArray(this BinaryReader br)
        {

            int[] result = new int[br.ReadCount()];

            for (int i = 0; i < result.Length; i++)
                result[i] = br.ReadInt32();

            return result;
        }
    }

}
