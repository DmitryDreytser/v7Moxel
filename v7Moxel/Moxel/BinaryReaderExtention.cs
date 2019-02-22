using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Moxel
{
    public static class BinaryReaderExtention
    {
        public class RequireStruct<T> where T : struct { }
        public class RequireClass<T> where T : class { }

        public static T Read<T>(this BinaryReader br, RequireStruct<T> ignore = null) where T : struct
        {
            var Length = Marshal.SizeOf(typeof(T));

            byte[] bytes = br.ReadBytes(Length);

            var ptr = Marshal.AllocHGlobal(bytes.Length);

            try
            {
                Marshal.Copy(bytes, 0, ptr, bytes.Length);
                //#pragma warning disable CS0618 // Type or member is obsolete
                T result = (T)Marshal.PtrToStructure(ptr, typeof(T));
                //#pragma warning restore CS0618 // Type or member is obsolete
                return result;
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }

        public static T Read<T>(this BinaryReader br, RequireClass<T> ignore = null) where T : class
        {
            return (T)Activator.CreateInstance(typeof(T), br);
        }

        public static Dictionary<int, T> ReadDictionary<T>(this BinaryReader br, RequireStruct<T> ignore = null) where T : struct
        {
            Dictionary<int, T> result = new Dictionary<int, T>();
            int[] numbers = br.ReadIntArray();
            int length = br.ReadCount();
            foreach (int num in numbers)
            {
                result.Add(num, br.Read<T>());
            }
            return result;
        }

        public static Dictionary<int, T> ReadDictionary<T>(this BinaryReader br, RequireClass<T> ignore = null) where T : class
        {
            Dictionary<int, T> result = new Dictionary<int, T>();
            int[] numbers = br.ReadIntArray();
            int length = br.ReadCount();
            foreach (int num in numbers)
            {
                result.Add(num, br.Read<T>());
            }

            return result;
        }

        public static List<T> ReadList<T>(this BinaryReader br, RequireStruct<T> ignore = null) where T : struct
        {
            List<T> result = new List<T>();
            int length = br.ReadCount();
            for (int num = 0; num < length; num++)
            {
                result.Add(br.Read<T>());
            }
            return result;
        }

        public static List<T> ReadList<T>(this BinaryReader br, RequireClass<T> ignore = null) where T : class
        {
            List<T> result = new List<T>();
            int length = br.ReadCount();
            for (int num = 0; num < length; num++)
            {
                result.Add(br.Read<T>());
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

            if (Count == 0xFFFF)
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
