using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using static Moxel.MemoryReader;

namespace Moxel
{
    public class PE
    {
        #region defines
        [Flags]
        public enum DataSectionFlags : uint
        {
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            TypeReg = 0x00000000,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            TypeDsect = 0x00000001,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            TypeNoLoad = 0x00000002,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            TypeGroup = 0x00000004,
            /// <summary>
            /// The section should not be padded to the next boundary. This flag is obsolete and is replaced by IMAGE_SCN_ALIGN_1BYTES. This is valid only for object files.
            /// </summary>
            TypeNoPadded = 0x00000008,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            TypeCopy = 0x00000010,
            /// <summary>
            /// The section contains executable code.
            /// </summary>
            ContentCode = 0x00000020,
            /// <summary>
            /// The section contains initialized data.
            /// </summary>
            ContentInitializedData = 0x00000040,
            /// <summary>
            /// The section contains uninitialized data.
            /// </summary>
            ContentUninitializedData = 0x00000080,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            LinkOther = 0x00000100,
            /// <summary>
            /// The section contains comments or other information. The .drectve section has this type. This is valid for object files only.
            /// </summary>
            LinkInfo = 0x00000200,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            TypeOver = 0x00000400,
            /// <summary>
            /// The section will not become part of the image. This is valid only for object files.
            /// </summary>
            LinkRemove = 0x00000800,
            /// <summary>
            /// The section contains COMDAT data. For more information, see section 5.5.6, COMDAT Sections (Object Only). This is valid only for object files.
            /// </summary>
            LinkComDat = 0x00001000,
            /// <summary>
            /// Reset speculative exceptions handling bits in the TLB entries for this section.
            /// </summary>
            NoDeferSpecExceptions = 0x00004000,
            /// <summary>
            /// The section contains data referenced through the global pointer (GP).
            /// </summary>
            RelativeGP = 0x00008000,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            MemPurgeable = 0x00020000,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            Memory16Bit = 0x00020000,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            MemoryLocked = 0x00040000,
            /// <summary>
            /// Reserved for future use.
            /// </summary>
            MemoryPreload = 0x00080000,
            /// <summary>
            /// Align data on a 1-byte boundary. Valid only for object files.
            /// </summary>
            Align1Bytes = 0x00100000,
            /// <summary>
            /// Align data on a 2-byte boundary. Valid only for object files.
            /// </summary>
            Align2Bytes = 0x00200000,
            /// <summary>
            /// Align data on a 4-byte boundary. Valid only for object files.
            /// </summary>
            Align4Bytes = 0x00300000,
            /// <summary>
            /// Align data on an 8-byte boundary. Valid only for object files.
            /// </summary>
            Align8Bytes = 0x00400000,
            /// <summary>
            /// Align data on a 16-byte boundary. Valid only for object files.
            /// </summary>
            Align16Bytes = 0x00500000,
            /// <summary>
            /// Align data on a 32-byte boundary. Valid only for object files.
            /// </summary>
            Align32Bytes = 0x00600000,
            /// <summary>
            /// Align data on a 64-byte boundary. Valid only for object files.
            /// </summary>
            Align64Bytes = 0x00700000,
            /// <summary>
            /// Align data on a 128-byte boundary. Valid only for object files.
            /// </summary>
            Align128Bytes = 0x00800000,
            /// <summary>
            /// Align data on a 256-byte boundary. Valid only for object files.
            /// </summary>
            Align256Bytes = 0x00900000,
            /// <summary>
            /// Align data on a 512-byte boundary. Valid only for object files.
            /// </summary>
            Align512Bytes = 0x00A00000,
            /// <summary>
            /// Align data on a 1024-byte boundary. Valid only for object files.
            /// </summary>
            Align1024Bytes = 0x00B00000,
            /// <summary>
            /// Align data on a 2048-byte boundary. Valid only for object files.
            /// </summary>
            Align2048Bytes = 0x00C00000,
            /// <summary>
            /// Align data on a 4096-byte boundary. Valid only for object files.
            /// </summary>
            Align4096Bytes = 0x00D00000,
            /// <summary>
            /// Align data on an 8192-byte boundary. Valid only for object files.
            /// </summary>
            Align8192Bytes = 0x00E00000,
            /// <summary>
            /// The section contains extended relocations.
            /// </summary>
            LinkExtendedRelocationOverflow = 0x01000000,
            /// <summary>
            /// The section can be discarded as needed.
            /// </summary>
            MemoryDiscardable = 0x02000000,
            /// <summary>
            /// The section cannot be cached.
            /// </summary>
            MemoryNotCached = 0x04000000,
            /// <summary>
            /// The section is not pageable.
            /// </summary>
            MemoryNotPaged = 0x08000000,
            /// <summary>
            /// The section can be shared in memory.
            /// </summary>
            MemoryShared = 0x10000000,
            /// <summary>
            /// The section can be executed as code.
            /// </summary>
            MemoryExecute = 0x20000000,
            /// <summary>
            /// The section can be read.
            /// </summary>
            MemoryRead = 0x40000000,
            /// <summary>
            /// The section can be written to.
            /// </summary>
            MemoryWrite = 0x80000000
        }

        public enum MachineType : ushort
        {
            Native = 0,
            I386 = 0x014c,
            Itanium = 0x0200,
            x64 = 0x8664
        }

        public enum MagicType : ushort
        {
            IMAGE_NT_OPTIONAL_HDR32_MAGIC = 0x10b,
            IMAGE_NT_OPTIONAL_HDR64_MAGIC = 0x20b
        }

        public enum SubSystemType : ushort
        {
            IMAGE_SUBSYSTEM_UNKNOWN = 0,
            IMAGE_SUBSYSTEM_NATIVE = 1,
            IMAGE_SUBSYSTEM_WINDOWS_GUI = 2,
            IMAGE_SUBSYSTEM_WINDOWS_CUI = 3,
            IMAGE_SUBSYSTEM_POSIX_CUI = 7,
            IMAGE_SUBSYSTEM_WINDOWS_CE_GUI = 9,
            IMAGE_SUBSYSTEM_EFI_APPLICATION = 10,
            IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER = 11,
            IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER = 12,
            IMAGE_SUBSYSTEM_EFI_ROM = 13,
            IMAGE_SUBSYSTEM_XBOX = 14

        }

        public enum DllCharacteristicsType : ushort
        {
            RES_0 = 0x0001,
            RES_1 = 0x0002,
            RES_2 = 0x0004,
            RES_3 = 0x0008,
            IMAGE_DLL_CHARACTERISTICS_DYNAMIC_BASE = 0x0040,
            IMAGE_DLL_CHARACTERISTICS_FORCE_INTEGRITY = 0x0080,
            IMAGE_DLL_CHARACTERISTICS_NX_COMPAT = 0x0100,
            IMAGE_DLLCHARACTERISTICS_NO_ISOLATION = 0x0200,
            IMAGE_DLLCHARACTERISTICS_NO_SEH = 0x0400,
            IMAGE_DLLCHARACTERISTICS_NO_BIND = 0x0800,
            RES_4 = 0x1000,
            IMAGE_DLLCHARACTERISTICS_WDM_DRIVER = 0x2000,
            IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE = 0x8000
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct IMAGE_SECTION_HEADER
        {
            [FieldOffset(0)]
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public char[] Name;

            [FieldOffset(8)]
            public UInt32 VirtualSize;

            [FieldOffset(12)]
            public UInt32 VirtualAddress;

            [FieldOffset(16)]
            public UInt32 SizeOfRawData;

            [FieldOffset(20)]
            public UInt32 PointerToRawData;

            [FieldOffset(24)]
            public UInt32 PointerToRelocations;

            [FieldOffset(28)]
            public UInt32 PointerToLinenumbers;

            [FieldOffset(32)]
            public UInt16 NumberOfRelocations;

            [FieldOffset(34)]
            public UInt16 NumberOfLinenumbers;

            [FieldOffset(36)]
            public DataSectionFlags Characteristics;

            public string Section
            {
                get { return new string(Name); }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IMAGE_DATA_DIRECTORY
        {
            public UInt32 VirtualAddress;
            public UInt32 Size;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct IMAGE_OPTIONAL_HEADER32
        {
            [FieldOffset(0)]
            public MagicType Magic;

            [FieldOffset(2)]
            public byte MajorLinkerVersion;

            [FieldOffset(3)]
            public byte MinorLinkerVersion;

            [FieldOffset(4)]
            public uint SizeOfCode;

            [FieldOffset(8)]
            public uint SizeOfInitializedData;

            [FieldOffset(12)]
            public uint SizeOfUninitializedData;

            [FieldOffset(16)]
            public uint AddressOfEntryPoint;

            [FieldOffset(20)]
            public uint BaseOfCode;

            // PE32 contains this additional field
            [FieldOffset(24)]
            public uint BaseOfData;

            [FieldOffset(28)]
            public uint ImageBase;

            [FieldOffset(32)]
            public uint SectionAlignment;

            [FieldOffset(36)]
            public uint FileAlignment;

            [FieldOffset(40)]
            public ushort MajorOperatingSystemVersion;

            [FieldOffset(42)]
            public ushort MinorOperatingSystemVersion;

            [FieldOffset(44)]
            public ushort MajorImageVersion;

            [FieldOffset(46)]
            public ushort MinorImageVersion;

            [FieldOffset(48)]
            public ushort MajorSubsystemVersion;

            [FieldOffset(50)]
            public ushort MinorSubsystemVersion;

            [FieldOffset(52)]
            public uint Win32VersionValue;

            [FieldOffset(56)]
            public uint SizeOfImage;

            [FieldOffset(60)]
            public uint SizeOfHeaders;

            [FieldOffset(64)]
            public uint CheckSum;

            [FieldOffset(68)]
            public SubSystemType Subsystem;

            [FieldOffset(70)]
            public DllCharacteristicsType DllCharacteristics;

            [FieldOffset(72)]
            public uint SizeOfStackReserve;

            [FieldOffset(76)]
            public uint SizeOfStackCommit;

            [FieldOffset(80)]
            public uint SizeOfHeapReserve;

            [FieldOffset(84)]
            public uint SizeOfHeapCommit;

            [FieldOffset(88)]
            public uint LoaderFlags;

            [FieldOffset(92)]
            public uint NumberOfRvaAndSizes;

            [FieldOffset(96)]
            public IMAGE_DATA_DIRECTORY ExportTable;

            [FieldOffset(104)]
            public IMAGE_DATA_DIRECTORY ImportTable;

            [FieldOffset(112)]
            public IMAGE_DATA_DIRECTORY ResourceTable;

            [FieldOffset(120)]
            public IMAGE_DATA_DIRECTORY ExceptionTable;

            [FieldOffset(128)]
            public IMAGE_DATA_DIRECTORY CertificateTable;

            [FieldOffset(136)]
            public IMAGE_DATA_DIRECTORY BaseRelocationTable;

            [FieldOffset(144)]
            public IMAGE_DATA_DIRECTORY Debug;

            [FieldOffset(152)]
            public IMAGE_DATA_DIRECTORY Architecture;

            [FieldOffset(160)]
            public IMAGE_DATA_DIRECTORY GlobalPtr;

            [FieldOffset(168)]
            public IMAGE_DATA_DIRECTORY TLSTable;

            [FieldOffset(176)]
            public IMAGE_DATA_DIRECTORY LoadConfigTable;

            [FieldOffset(184)]
            public IMAGE_DATA_DIRECTORY BoundImport;

            [FieldOffset(192)]
            public IMAGE_DATA_DIRECTORY IAT;

            [FieldOffset(200)]
            public IMAGE_DATA_DIRECTORY DelayImportDescriptor;

            [FieldOffset(208)]
            public IMAGE_DATA_DIRECTORY CLRRuntimeHeader;

            [FieldOffset(216)]
            public IMAGE_DATA_DIRECTORY Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IMAGE_DOS_HEADER
        {
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
            public char[] e_magic;       // Magic number
            public short e_cblp;    // Bytes on last page of file
            public short e_cp;      // Pages in file
            public short e_crlc;    // Relocations
            public short e_cparhdr;     // Size of header in paragraphs
            public short e_minalloc;    // Minimum extra paragraphs needed
            public short e_maxalloc;    // Maximum extra paragraphs needed
            public short e_ss;      // Initial (relative) SS value
            public short e_sp;      // Initial SP value
            public short e_csum;    // Checksum
            public short e_ip;      // Initial IP value
            public short e_cs;      // Initial (relative) CS value
            public short e_lfarlc;      // File address of relocation table
            public short e_ovno;    // Overlay number
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public short[] e_res1;    // Reserved words
            public short e_oemid;       // OEM identifier (for e_oeminfo)
            public short e_oeminfo;     // OEM information; e_oemid specific
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
            public short[] e_res2;    // Reserved words
            public int e_lfanew;      // File address of new exe header

            private string _e_magic
            {
                get { return new string(e_magic); }
            }

            public bool isValid
            {
                get { return _e_magic == "MZ"; }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IMAGE_FILE_HEADER
        {
            public UInt16 Machine;
            public UInt16 NumberOfSections;
            public UInt32 TimeDateStamp;
            public UInt32 PointerToSymbolTable;
            public UInt32 NumberOfSymbols;
            public UInt16 SizeOfOptionalHeader;
            public UInt16 Characteristics;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct IMAGE_NT_HEADERS32
        {
            [FieldOffset(0)]
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
            public char[] Signature;

            [FieldOffset(4)]
            public IMAGE_FILE_HEADER FileHeader;

            [FieldOffset(24)]
            public IMAGE_OPTIONAL_HEADER32 OptionalHeader;

            private string _Signature
            {
                get { return new string(Signature); }
            }

            public bool isValid
            {
                get { return _Signature == "PE\0\0" && (OptionalHeader.Magic == MagicType.IMAGE_NT_OPTIONAL_HDR32_MAGIC || OptionalHeader.Magic == MagicType.IMAGE_NT_OPTIONAL_HDR64_MAGIC); }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IMAGE_EXPORT_DIRECTORY
        {
            public uint Characteristics;
            public uint TimeDateStamp;
            public ushort MajorVersion;
            public ushort MinorVersion;
            public uint Name;
            public uint Base;
            public uint NumberOfFunctions;
            public uint NumberOfNames;
            public uint AddressOfFunctions;     // RVA from base of image
            public uint AddressOfNames;     // RVA from base of image
            public uint AddressOfNameOrdinals;  // RVA from base of image

        }
        #endregion

        public IntPtr ImageBase = IntPtr.Zero;
        public Int64 dw_ImageBase;
        public IMAGE_DOS_HEADER DOSHeader;
        public IMAGE_NT_HEADERS32 NTHeader;
        public IMAGE_EXPORT_DIRECTORY ExportDirectory;

        public PE(string ModuleName)
        {
            ImageBase = WinApi.GetModuleHandle(ModuleName);
            dw_ImageBase = ImageBase.ToInt64();
            DOSHeader = Marshal.PtrToStructure<IMAGE_DOS_HEADER>(ImageBase);
            NTHeader = Marshal.PtrToStructure<IMAGE_NT_HEADERS32>(new IntPtr(dw_ImageBase + DOSHeader.e_lfanew));
            ExportDirectory = Marshal.PtrToStructure<IMAGE_EXPORT_DIRECTORY>(new IntPtr(dw_ImageBase + NTHeader.OptionalHeader.ExportTable.VirtualAddress));

           
        }

        public IntPtr GetProcExportAddress(string ProcName)
        {
            IntPtr pProc = WinApi.GetProcAddress(ImageBase, ProcName);

            short[] Ordinals = new short[ExportDirectory.NumberOfNames];

            Marshal.Copy(new IntPtr(dw_ImageBase + ExportDirectory.AddressOfNameOrdinals), Ordinals, 0, (int)ExportDirectory.NumberOfNames);
            foreach(int Ordinal in Ordinals)
            {
                IntPtr pCurProc = new IntPtr(dw_ImageBase + ExportDirectory.AddressOfFunctions);
                pCurProc = new IntPtr(dw_ImageBase + Marshal.ReadInt32(pCurProc + Ordinal * 4));

                if(pCurProc == pProc)
                {
                    return new IntPtr(dw_ImageBase + ExportDirectory.AddressOfFunctions + Ordinal * 4);
                }
            }

            return IntPtr.Zero;
        }



    }

    public static class SaveWrapper
    {
        [UnmanagedFunctionPointer(CallingConvention.Cdecl, CharSet = CharSet.Ansi, ThrowOnUnmappableChar = true)]
        public delegate int dSaveToExcel(IntPtr SheetDoc, IntPtr SheetGDI, string FileName);

        static dSaveToExcel WrapperProc = new dSaveToExcel(SaveToExcel);
        static GCHandle wrapper = GCHandle.Alloc(WrapperProc);
        static IntPtr pWrapperProc = Marshal.GetFunctionPointerForDelegate<dSaveToExcel>(WrapperProc);

        static bool isWraped = false;
        static IntPtr ProcRVA = IntPtr.Zero;

        static int SaveToExcel(IntPtr pSheetDoc, IntPtr SheetGDI, string FileName)
        {
            CFile f = CFile.FromHFile(IntPtr.Zero);
            int result = 0;
            try
            {
                CSheetDoc SheetDoc = new CSheetDoc(pSheetDoc);
                
                CArchive Arch = new CArchive(f, SheetDoc);
                SheetDoc.Serialize(Arch);
                Arch.Flush();
                Arch = null;
                byte[] buffer = f.GetBufer();
                f = null;

                Moxel mxl = new Moxel(ref buffer);

                if (FileName.EndsWith(".xlsx"))
                    mxl.SaveAs(FileName, SaveFormat.Excel);
                else
                {
                    if (FileName.EndsWith(".xls"))
                    {
                        mxl.SaveAs(FileName + "x", SaveFormat.Excel);
                        File.Move(FileName + "x", FileName);
                    }
                    else
                        mxl.SaveAs(FileName + ".xlsx", SaveFormat.Excel);
                }
                result = 1;
            }
            catch(Exception ex)
            {
                f.unpatch(); // Обязательно снять перехват
                f = null;
                result = 0;
            }

            return result;
        }

        static int hMoxel = 0;
        static IntPtr hProcess;// = Process.GetCurrentProcess().Handle;

        static byte[] OriginalBytes = new byte[6];

        static byte[] oldRes = new byte[288];

        static IntPtr FileSaveFilterResource;

        public static int Wrap(bool DoWrap)
        {

            if (hMoxel == 0)
            {
                hMoxel = WinApi.GetModuleHandle("Moxel.dll").ToInt32();
                ProcRVA = new IntPtr(hMoxel + 0x5B420);
                Marshal.Copy(ProcRVA, OriginalBytes, 0, 6);
            }

            uint Protection = 0;
            try
            {
                if (DoWrap)
                {
                    if (isWraped)
                        return 1;
                    hProcess = Process.GetCurrentProcess().Handle;

                    IntPtr RetAddress = Marshal.GetFunctionPointerForDelegate<dSaveToExcel>(WrapperProc);

                    if (WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), 0x40, out Protection))
                    {
                        Debug.WriteLine($"Патч по адресу {ProcRVA.ToInt32():X8}: PUSH {RetAddress.ToInt32():X8}; RET;");

                        Marshal.WriteByte(ProcRVA, 0x68); //PUSH
                        Marshal.WriteIntPtr(new IntPtr(ProcRVA.ToInt32() + 1), RetAddress);
                        Marshal.WriteByte(new IntPtr(ProcRVA.ToInt32() + 5), 0xC3); //RET

                        WinApi.FlushInstructionCache(hProcess, ProcRVA, new IntPtr(6));

                        WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), Protection, out Protection);
                        

                        //Заменим фильтр диалога сохранения. Заменим *.xls на *.xlsx

                        IntPtr hRCRus = WinApi.GetModuleHandle("1crcrus.dll");

                        IntPtr Res = WinApi.FindResource(hRCRus, WinApi.MakeIntResource(0x770), 6);

                        int len = WinApi.SizeofResource(hRCRus, Res);
                        Array.Resize(ref oldRes, len);

                        Res = WinApi.LoadResource(hRCRus, Res);

                        FileSaveFilterResource = WinApi.LockResource(Res);

                        Debug.WriteLine($"Ресурс найден по адресу {FileSaveFilterResource.ToInt32():X8}");

                        Marshal.Copy(FileSaveFilterResource, oldRes, 0, len);


                        Char[] newRes = new Char[len / 2];
                        Array.Clear(newRes, 0, newRes.Length);

                        string[] DialogFilter = { "Таблица Excel 2007 (*.xlsx)|*.xlsx", "HTML Документ (*.html)|*.html", "Текстовый файл (*.txt)|*.txt" };

                        int charindex = 0;
                        foreach (string Filter in DialogFilter)
                        {
                            newRes[charindex++] = (Char)Filter.Length;

                            foreach (Char chr in Filter)
                            {
                                newRes[charindex++] = chr;
                            }
                        }

                        byte[] buffer = new byte[charindex * 2];

                        Array.Resize(ref newRes, charindex);

                        Buffer.BlockCopy(newRes, 0, buffer, 0, newRes.Length * 2);

                        if (WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(buffer.Length), 0x40, out Protection))
                        {

                            Marshal.Copy(buffer, 0, FileSaveFilterResource, buffer.Length);

                            WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(buffer.Length), Protection, out Protection);
                        }

                        isWraped = true;
                    }

                }
                else
                {
                    if (!isWraped)
                        return 1;

                    hProcess = Process.GetCurrentProcess().Handle;

                    if (WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), 0x40, out Protection));
                    {
                        Marshal.Copy(OriginalBytes, 0, ProcRVA, 6);

                        Debug.WriteLine($"Снят патч по адресу {ProcRVA.ToInt32():X8}");

                        WinApi.FlushInstructionCache(hProcess, ProcRVA, new IntPtr(6));

                        WinApi.VirtualProtectEx(hProcess, ProcRVA, new IntPtr(6), Protection, out Protection);


                        //Вернем фильтр диалога сохранения таблиц на место

                        if(WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), 0x40, out Protection));
                        {

                            Marshal.Copy(oldRes, 0, FileSaveFilterResource, oldRes.Length);

                            WinApi.VirtualProtectEx(hProcess, FileSaveFilterResource, new IntPtr(oldRes.Length), Protection, out Protection);
                        }

                        isWraped = false;
                    }
                }

                if (isWraped == DoWrap)
                    return 1;
                else
                    return 0;
            }
            catch
            {
                return 0;
            }

        }


    }
}
