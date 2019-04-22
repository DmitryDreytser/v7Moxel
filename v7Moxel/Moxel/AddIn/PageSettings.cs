using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Moxel.MemoryReader;

namespace Moxel
{



    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct CArray<T> where T : new()
    {
        public IntPtr m_Pobj;
        public IntPtr m_pEntrys;
        public int Count;
        public int Capacity;
        public int XZ1;
        public T[] Entrys
        {
            get
            {
                IntPtr[] _refs = new IntPtr[Count];
                T[] result = new T[Count];
                int index = 0;

                unsafe
                {
                    fixed (IntPtr* pDst = _refs)
                    {
                        byte* pSrc = (byte*)m_pEntrys.ToPointer();
                        var lenght = _refs.Length * sizeof(IntPtr);
                        Buffer.MemoryCopy(pSrc, pDst, lenght, lenght);

                        foreach (IntPtr p in _refs)
                        {
                            if (typeof(T) == typeof(IntPtr))
                            {
                                result[index++] = p.DynamicCast<T>();
                            }
                            else
                            {
                                result[index++] = Marshal.PtrToStructure<T>(p);
                            }
                        }

                    }
                }
                return result;
            }
        }
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public class PageSettings //CProfile7
    {
        public enum OptionType
        {
            Paper = 0,
            Orient = 1,
            Scale = 2,
            Collate = 3,
            Copyes = 4,
            PerPage = 5,
            Top = 6,
            Left = 7,
            Bottom = 8,
            Right = 9,
            Header = 10,
            Footer = 11,
            RepeatRowFrom = 12,
            RepeatRowTo = 13,
            RepeatColFrom = 14,
            RepeatColTo = 15,
            RangeTop = 16,
            RangeBottom = 17,
            RangeLeft = 18,
            RangeRight = 19,
            Protection = 20,
            FitToPage = 21,
            BlackAndWhite = 22,
            DefaultPrinter = 23,
            NextMode = 24,
            PaperSource = 25
        }

        IntPtr m_pObj;
        [MarshalAs(UnmanagedType.LPStr)]
        public string Name;          // 04h + 04h
        
        //[MarshalAs(UnmanagedType.LPStruct, MarshalType = nameof(CProfile7))]
        IntPtr m_pParentProfile;  // 08h + 04h

        /*CPtrList m_SubProfileList;  */// 0Ch + 1Ch
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 7)]
        IntPtr[] m_SubProfileList;

        /*CItemList*/
        IntPtr m_pItemList;       // 28h + 04h

        [MarshalAs(UnmanagedType.LPStr)]
        public string Path;         // 2Ch + 04h

        /*CPtrArray m_PropArray;       */// 30h + 14h
        [MarshalAs(UnmanagedType.Struct)]
        CArray<IntPtr> m_PropArray;

        /*CProfileEntryArr m_Entrys;    */// 44h + 14h
        [MarshalAs(UnmanagedType.Struct)]
        CArray<CProfileEntry7> m_Entrys;

        public PageSettings ParentProfile { get { return PageSettings.FromIntPtr(m_pParentProfile); } }

        public CProfileEntry7[] PropertyList { get { return m_Entrys.Entrys; } }
        public IntPtr[] PropertyValues { get { return m_PropArray.Entrys; } }

        public object Get(OptionType opt)
        {
            int index = (int)opt;
            if (index > m_Entrys.Count)
                return null;

            CProfileEntry7 valueType = PropertyList[index];
            IntPtr PropValue = PropertyValues[index];

            if (PropValue == IntPtr.Zero)
                return valueType.DefaultValue;

            switch (valueType.type)
            {
                case global::Moxel.OptionType.Int:
                    return PropValue.ToInt32();
                case global::Moxel.OptionType.Long:
                    return (long)PropValue.ToInt32();
                case global::Moxel.OptionType.Double:
                    return (double)PropValue.ToInt32();
                case global::Moxel.OptionType.String:
                    return Marshal.PtrToStringAnsi(PropValue);
                default:
                    return PropValue;
            }

        }

        public static PageSettings FromIntPtr(IntPtr m_pObj)
        {
            return Marshal.PtrToStructure<PageSettings>(m_pObj);
        }
    }

    public enum OptionType : uint
    {
        Int = 1,
        Long = 2,
        Double = 3,
        String = 4,
        Date = 5,
        Numeric = 6,
        Color = 7,
    }

    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public class CProfileEntry7
    {
        public OptionType type;
        uint data1;
        [MarshalAs(UnmanagedType.LPStr)]
        string str;
        [MarshalAs(UnmanagedType.LPStr)]
        string Name;
        [MarshalAs(UnmanagedType.LPStr)]
        string value;

        public override string ToString()
        {
            return $"{Name} : {value}";
        }

        public object DefaultValue
        {
            get
            {
                switch (type)
                {
                    case OptionType.Int:
                        return int.Parse(value);
                    case OptionType.Long:
                        return long.Parse(value);
                    case OptionType.Double:
                        return double.Parse(value);
                    case OptionType.String:
                        return value;
                    default:
                        return value;
                }
            }
        }
    }
}
