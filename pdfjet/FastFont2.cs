/**
 *  FastFont2.cs
 *
Copyright (c) 2018, Innovatics Inc.
All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright notice,
      this list of conditions and the following disclaimer.

    * Redistributions in binary form must reproduce the above copyright notice,
      this list of conditions and the following disclaimer in the documentation
      and / or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR
CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

using System;
using System.IO;
using System.Text;
using System.Collections.Generic;


namespace PDFjet.NET {
class FastFont2 {

    internal static void Register(
            SortedDictionary<Int32, PDFobj> objects,
            Font font,
            Stream inputStream) {

        int len = inputStream.ReadByte();
        byte[] fontName = new byte[len];
        inputStream.Read(fontName, 0, len);
        font.name = System.Text.Encoding.UTF8.GetString(fontName, 0, len);
        // Console.WriteLine(font.name);

        len = GetInt24(inputStream);
        byte[] fontInfo = new byte[len];
        inputStream.Read(fontInfo, 0, len);
        font.info = System.Text.Encoding.UTF8.GetString(fontInfo, 0, len);
        // Console.WriteLine(font.info);

        byte[] buf = new byte[GetInt32(inputStream)];
        inputStream.Read(buf, 0, buf.Length);
        Decompressor decompressor = new Decompressor(buf);
        MemoryStream stream =
                new MemoryStream(decompressor.GetDecompressedData());

        font.unitsPerEm = GetInt32(stream);
        font.bBoxLLx = GetInt32(stream);
        font.bBoxLLy = GetInt32(stream);
        font.bBoxURx = GetInt32(stream);
        font.bBoxURy = GetInt32(stream);
        font.ascent = GetInt32(stream);
        font.descent = GetInt32(stream);
        font.firstChar = GetInt32(stream);
        font.lastChar = GetInt32(stream);
        font.capHeight = GetInt32(stream);
        font.underlinePosition = GetInt32(stream);
        font.underlineThickness = GetInt32(stream);

        len = GetInt32(stream);
        font.advanceWidth = new int[len];
        for (int i = 0; i < len; i++) {
            font.advanceWidth[i] = GetInt16(stream);
        }

        len = GetInt32(stream);
        font.glyphWidth = new int[len];
        for (int i = 0; i < len; i++) {
            font.glyphWidth[i] = GetInt16(stream);
        }

        len = GetInt32(stream);
        font.unicodeToGID = new int[len];
        for (int i = 0; i < len; i++) {
            font.unicodeToGID[i] = GetInt16(stream);
        }

        font.cff = (inputStream.ReadByte() == 'Y') ? true : false;
        font.uncompressed_size = GetInt32(inputStream);
        font.compressed_size = GetInt32(inputStream);

        EmbedFontFile(objects, font, inputStream);
        AddFontDescriptorObject(objects, font);
        AddCIDFontDictionaryObject(objects, font);
        AddToUnicodeCMapObject(objects, font);

        // Type0 Font Dictionary
        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Type");
        dict.Add("/Font");
        dict.Add("/Subtype");
        dict.Add("/Type0");
        dict.Add("/BaseFont");
        dict.Add("/" + font.name);
        dict.Add("/Encoding");
        dict.Add("/Identity-H");
        dict.Add("/DescendantFonts");
        dict.Add("[");
        dict.Add(font.GetCidFontDictObjNumber().ToString());
        dict.Add("0");
        dict.Add("R");
        dict.Add("]");
        dict.Add("/ToUnicode");
        dict.Add(font.GetToUnicodeCMapObjNumber().ToString());
        dict.Add("0");
        dict.Add("R");
        dict.Add(">>");
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        font.objNumber = obj.number;
    }


    private static int AddMetadataObject(
            SortedDictionary<Int32, PDFobj> objects, Font font) {

        StringBuilder sb = new StringBuilder();
        sb.Append("<?xpacket begin='\uFEFF' id=\"W5M0MpCehiHzreSzNTczkc9d\"?>\n");
        sb.Append("<x:xmpmeta xmlns:x=\"adobe:ns:meta/\">\n");
        sb.Append("<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\n");
        sb.Append("<rdf:Description rdf:about=\"\" xmlns:xmpRights=\"http://ns.adobe.com/xap/1.0/rights/\">\n");
        sb.Append("<xmpRights:UsageTerms>\n");
        sb.Append("<rdf:Alt>\n");
        sb.Append("<rdf:li xml:lang=\"x-default\">\n");
        sb.Append(font.info);
        sb.Append("</rdf:li>\n");
        sb.Append("</rdf:Alt>\n");
        sb.Append("</xmpRights:UsageTerms>\n");
        sb.Append("</rdf:Description>\n");
        sb.Append("</rdf:RDF>\n");
        sb.Append("</x:xmpmeta>\n");
        sb.Append("<?xpacket end=\"w\"?>");

        byte[] xml = (new System.Text.UTF8Encoding()).GetBytes(sb.ToString());

        // This is the metadata object
        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Type");
        dict.Add("/Metadata");
        dict.Add("/Subtype");
        dict.Add("/XML");
        dict.Add("/Length");
        dict.Add(xml.Length.ToString());
        dict.Add(">>");
        obj.SetStream(xml);
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);

        return obj.number;
    }


    private static void EmbedFontFile(
            SortedDictionary<Int32, PDFobj> objects,
            Font font,
            Stream inputStream) {

        int metadataObjNumber = AddMetadataObject(objects, font);

        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Metadata");
        dict.Add(metadataObjNumber.ToString());
        dict.Add("0");
        dict.Add("R");
        dict.Add("/Filter");
        dict.Add("/FlateDecode");
        dict.Add("/Length");
        dict.Add(font.compressed_size.ToString());
        if (font.cff) {
            dict.Add("/Subtype");
            dict.Add("/CIDFontType0C");
        }
        else {
            dict.Add("/Length1");
            dict.Add(font.uncompressed_size.ToString());
        }
        dict.Add(">>");
        MemoryStream buf2 = new MemoryStream();
        byte[] buf = new byte[16384];
        int len;
        while ((len = inputStream.Read(buf, 0, buf.Length)) > 0) {
            buf2.Write(buf, 0, len);
        }
        inputStream.Close();
        obj.SetStream(buf2.ToArray());
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        font.fileObjNumber = obj.number;
    }


    private static void AddFontDescriptorObject(
            SortedDictionary<Int32, PDFobj> objects, Font font) {

        float factor = 1000f / font.unitsPerEm;

        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Type");
        dict.Add("/FontDescriptor");
        dict.Add("/FontName");
        dict.Add("/" + font.name);
        dict.Add("/FontFile" + (font.cff ? "3" : "2"));
        dict.Add(font.fileObjNumber.ToString());
        dict.Add("0");
        dict.Add("R");
        dict.Add("/Flags");
        dict.Add("32");
        dict.Add("/FontBBox");
        dict.Add("[");
        dict.Add(((Int32) (font.bBoxLLx * factor)).ToString());
        dict.Add(((Int32) (font.bBoxLLy * factor)).ToString());
        dict.Add(((Int32) (font.bBoxURx * factor)).ToString());
        dict.Add(((Int32) (font.bBoxURy * factor)).ToString());
        dict.Add("]");
        dict.Add("/Ascent");
        dict.Add(((Int32) (font.ascent * factor)).ToString());
        dict.Add("/Descent");
        dict.Add(((Int32) (font.descent * factor)).ToString());
        dict.Add("/ItalicAngle");
        dict.Add("0");
        dict.Add("/CapHeight");
        dict.Add(((Int32) (font.capHeight * factor)).ToString());
        dict.Add("/StemV");
        dict.Add("79");
        dict.Add(">>");
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        font.SetFontDescriptorObjNumber(obj.number);
    }


    private static void AddToUnicodeCMapObject(
            SortedDictionary<Int32, PDFobj> objects, Font font) {

        StringBuilder sb = new StringBuilder();

        sb.Append("/CIDInit /ProcSet findresource begin\n");
        sb.Append("12 dict begin\n");
        sb.Append("begincmap\n");
        sb.Append("/CIDSystemInfo <</Registry (Adobe) /Ordering (Identity) /Supplement 0>> def\n");
        sb.Append("/CMapName /Adobe-Identity def\n");
        sb.Append("/CMapType 2 def\n");

        sb.Append("1 begincodespacerange\n");
        sb.Append("<0000> <FFFF>\n");
        sb.Append("endcodespacerange\n");

        List<String> list = new List<String>();
        StringBuilder buf = new StringBuilder();
        for (int cid = 0; cid <= 0xffff; cid++) {
            int gid = font.unicodeToGID[cid];
            if (gid > 0) {
                buf.Append('<');
                buf.Append(ToHexString(gid));
                buf.Append("> <");
                buf.Append(ToHexString(cid));
                buf.Append(">\n");
                list.Add(buf.ToString());
                buf.Length = 0;
                if (list.Count == 100) {
                    WriteListToBuffer(list, sb);
                }
            }
        }
        if (list.Count > 0) {
            WriteListToBuffer(list, sb);
        }
        sb.Append("endcmap\n");
        sb.Append("CMapName currentdict /CMap defineresource pop\n");
        sb.Append("end\nend");

        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Length");
        dict.Add(sb.Length.ToString());
        dict.Add(">>");
        obj.SetStream((new System.Text.UTF8Encoding()).GetBytes(sb.ToString()));
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        font.SetToUnicodeCMapObjNumber(obj.number);
    }


    private static void AddCIDFontDictionaryObject(
            SortedDictionary<Int32, PDFobj> objects, Font font) {

        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Type");
        dict.Add("/Font");
        dict.Add("/Subtype");
        dict.Add("/CIDFontType" + (font.cff ? "0" : "2"));
        dict.Add("/BaseFont");
        dict.Add("/" + font.name);
        dict.Add("/CIDSystemInfo");
        dict.Add("<<");
        dict.Add("/Registry");
        dict.Add("(Adobe)");
        dict.Add("/Ordering");
        dict.Add("(Identity)");
        dict.Add("/Supplement");
        dict.Add("0");
        dict.Add(">>");
        dict.Add("/FontDescriptor");
        dict.Add(font.GetFontDescriptorObjNumber().ToString());
        dict.Add("0");
        dict.Add("R");
        dict.Add("/DW");
        dict.Add(((Int32)
                ((1000f / font.unitsPerEm) * font.advanceWidth[0])).ToString());
        dict.Add("/W");
        dict.Add("[");
        dict.Add("0");
        dict.Add("[");
        for (int i = 0; i < font.advanceWidth.Length; i++) {
            dict.Add(((int)
                    ((1000f / font.unitsPerEm) * font.advanceWidth[i])).ToString());
        }
        dict.Add("]");
        dict.Add("]");
        dict.Add("/CIDToGIDMap");
        dict.Add("/Identity");
        dict.Add(">>");
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        font.SetCidFontDictObjNumber(obj.number);
    }


    private static String ToHexString(int code) {
        String str = Convert.ToString(code, 16);
        if (str.Length == 1) {
            return "000" + str;
        }
        else if (str.Length == 2) {
            return "00" + str;
        }
        else if (str.Length == 3) {
            return "0" + str;
        }
        return str;
    }


    private static void WriteListToBuffer(List<String> list, StringBuilder sb) {
        sb.Append(list.Count);
        sb.Append(" beginbfchar\n");
        foreach (String str in list) {
            sb.Append(str);
        }
        sb.Append("endbfchar\n");
        list.Clear();
    }


    private static int GetInt16(Stream stream) {
        return stream.ReadByte() << 8 | stream.ReadByte();
    }


    private static int GetInt24(Stream stream) {
        return stream.ReadByte() << 16 |
                stream.ReadByte() << 8 | stream.ReadByte();
    }


    private static int GetInt32(Stream stream) {
        return stream.ReadByte() << 24 | stream.ReadByte() << 16 |
                stream.ReadByte() << 8 | stream.ReadByte();
    }


    private static Int32 MaxKey(SortedDictionary<Int32, PDFobj>.KeyCollection keys) {
        Int32 maxKey = 0;
        foreach (Int32 key in keys) {
            if (key > maxKey) {
                maxKey = key;
            }
        }
        return maxKey;
    }

}   // End of FastFont2.cs
}   // End of namespace PDFjet.NET
