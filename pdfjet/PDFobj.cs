/**
 *  PDFobj.cs
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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace PDFjet.NET {
/**
 *  Used to create Java or .NET objects that represent the objects in PDF document.
 *  See the PDF specification for more information.
 *
 */
public class PDFobj {

    internal int offset;           // The object offset
    internal int number;           // The object number
    internal List<String> dict;
    internal int stream_offset;
    internal byte[] stream;        // The compressed stream
    internal byte[] data;          // The decompressed data
    internal int gsNumber = -1;


    /**
     *  Used to create Java or .NET objects that represent the objects in PDF document.
     *  See the PDF specification for more information.
     *  Also see Example_19.
     *
     *  @param offset the object offset in the offsets table.
     */
    public PDFobj(int offset) {
        this.offset = offset;
        this.dict = new List<String>();
    }


    internal PDFobj() {
        this.dict = new List<String>();
    }


    public int GetNumber() {
        return this.number;
    }


    public List<String> GetDict() {
        return this.dict;
    }


    public void SetDict(List<String> dict) {
        this.dict = dict;
    }


    public byte[] GetData() {
        return this.data;
    }


    internal void SetStream(byte[] pdf, int length) {
        stream = new byte[length];
        Array.Copy(pdf, this.stream_offset, stream, 0, length);
    }


    internal void SetStream(byte[] stream) {
        this.stream = stream;
    }


    internal void SetNumber(int number) {
        this.number = number;
    }


    /**
     *  Returns the parameter value given the specified key.
     *
     *  @param key the specified key.
     *
     *  @return the value.
     */
    public String GetValue(String key) {
        for (int i = 0; i < dict.Count; i++) {
            String token = dict[i];
            if (token.Equals(key)) {
                if (dict[i + 1].Equals("<<")) {
                    StringBuilder buffer = new StringBuilder();
                    buffer.Append("<<");
                    buffer.Append(" ");
                    i += 2;
                    while (!dict[i].Equals(">>")) {
                        buffer.Append(dict[i]);
                        buffer.Append(" ");
                        i += 1;
                    }
                    buffer.Append(">>");
                    return buffer.ToString();
                }
                if (dict[i + 1].Equals("[")) {
                    StringBuilder buffer = new StringBuilder();
                    buffer.Append("[");
                    buffer.Append(" ");
                    i += 2;
                    while (!dict[i].Equals("]")) {
                        buffer.Append(dict[i]);
                        buffer.Append(" ");
                        i += 1;
                    }
                    buffer.Append("]");
                    return buffer.ToString();
                }
                return dict[i + 1];
            }
        }
        return "";
    }


    internal List<Int32> GetObjectNumbers(String key) {
        List<Int32> numbers = new List<Int32>();
        for (int i = 0; i < dict.Count; i++) {
            String token = dict[i];
            if (token.Equals(key)) {
                String str = dict[++i];
                if (str.Equals("[")) {
                    while (true) {
                        str = dict[++i];
                        if (str.Equals("]")) {
                            break;
                        }
                        numbers.Add(Int32.Parse(str));
                        ++i;    // 0
                        ++i;    // R
                    }
                }
                else {
                    numbers.Add(Int32.Parse(str));
                }
                break;
            }
        }
        return numbers;
    }


    public void AddContentObject(int number) {
        int index = -1;
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Contents")) {
                String str = dict[++i];
                if (str.Equals("[")) {
                    while (true) {
                        str = dict[++i];
                        if (str.Equals("]")) {
                            index = i;
                            break;
                        }
                        ++i;    // 0
                        ++i;    // R
                    }
                }
                break;
            }
        }
        dict.Insert(index, "R");
        dict.Insert(index, "0");
        dict.Insert(index, number.ToString());
    }


    public float[] GetPageSize() {
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/MediaBox")) {
                return new float[] {
                        Convert.ToSingle(dict[i + 4]),
                        Convert.ToSingle(dict[i + 5]) };
            }
        }
        return Letter.PORTRAIT;
    }


    internal int GetLength(List<PDFobj> objects) {
        for (int i = 0; i < dict.Count; i++) {
            String token = dict[i];
            if (token.Equals("/Length")) {
                int number = Int32.Parse(dict[i + 1]);
                if (dict[i + 2].Equals("0") &&
                        dict[i + 3].Equals("R")) {
                    return GetLength(objects, number);
                }
                else {
                    return number;
                }
            }
        }
        return 0;
    }


    internal int GetLength(List<PDFobj> objects, int number) {
        foreach (PDFobj obj in objects) {
            if (obj.number == number) {
                return Int32.Parse(obj.dict[3]);
            }
        }
        return 0;
    }


    public PDFobj GetContentsObject(SortedDictionary<Int32, PDFobj> objects) {
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Contents")) {
                if (dict[i + 1].Equals("[")) {
                    return objects[Int32.Parse(dict[i + 2])];
                }
                return objects[Int32.Parse(dict[i + 1])];
            }
        }
        return null;
    }


    public PDFobj GetResourcesObject(SortedDictionary<Int32, PDFobj> objects) {
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Resources")) {
                String token = dict[i + 1];
                if (token.Equals("<<")) {
                    PDFobj obj = new PDFobj();
                    obj.dict.Add("0");
                    obj.dict.Add("0");
                    obj.dict.Add("obj");
                    obj.dict.Add(token);
                    int level = 1;
                    i++;
                    while (i < dict.Count && level > 0) {
                        token = dict[i];
                        obj.dict.Add(token);
                        if (token.Equals("<<")) {
                            level++;
                        }
                        else if (token.Equals(">>")) {
                            level--;
                        }
                        i++;
                    }
                    return obj;
                }
                return objects[Int32.Parse(token)];
            }
        }
        return null;
    }


    public Font AddFontResource(CoreFont coreFont, SortedDictionary<Int32, PDFobj> objects) {
        Font font = new Font(coreFont);
        font.fontID = font.name.Replace('-', '_').ToUpper();

        PDFobj obj = new PDFobj();

        int maxObjNumber = -1;
        foreach (int number in objects.Keys) {
            if (number > maxObjNumber) { maxObjNumber = number; }
        }
        obj.number = maxObjNumber + 1;

        obj.dict.Add("<<");
        obj.dict.Add("/Type");
        obj.dict.Add("/Font");
        obj.dict.Add("/Subtype");
        obj.dict.Add("/Type1");
        obj.dict.Add("/BaseFont");
        obj.dict.Add("/" + font.name);
        if (!font.name.Equals("Symbol") && !font.name.Equals("ZapfDingbats")) {
            obj.dict.Add("/Encoding");
            obj.dict.Add("/WinAnsiEncoding");
        }
        obj.dict.Add(">>");

        objects.Add(obj.number, obj);

        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Resources")) {
                String token = dict[++i];
                if (token.Equals("<<")) {                   // Direct resources object
                    AddFontResource(this, objects, font.fontID, obj.number);
                }
                else if (Char.IsDigit(token[0])) {          // Indirect resources object
                    AddFontResource(objects[Int32.Parse(token)], objects, font.fontID, obj.number);
                }
            }
        }

        return font;
    }


    private void AddFontResource(
            PDFobj obj, SortedDictionary<Int32, PDFobj> objects, String fontID, int number) {

        bool fonts = false;
        for (int i = 0; i < obj.dict.Count; i++) {
            if (obj.dict[i].Equals("/Font")) {
                fonts = true;
            }
        }
        if (!fonts) {
            for (int i = 0; i < obj.dict.Count; i++) {
                if (obj.dict[i].Equals("/Resources")) {
                    obj.dict.Insert(i + 2, "/Font");
                    obj.dict.Insert(i + 3, "<<");
                    obj.dict.Insert(i + 4, ">>");
                    break;
                }
            }
        }

        for (int i = 0; i < obj.dict.Count; i++) {
            if (obj.dict[i].Equals("/Font")) {
                String token = obj.dict[i + 1];
                if (token.Equals("<<")) {
                    obj.dict.Insert(i + 2, "/" + fontID);
                    obj.dict.Insert(i + 3, number.ToString());
                    obj.dict.Insert(i + 4, "0");
                    obj.dict.Insert(i + 5, "R");
                    return;
                }
                else if (Char.IsDigit(token[0])) {
                    PDFobj o2 = objects[Int32.Parse(token)];
                    for (int j = 0; j < o2.dict.Count; j++) {
                        if (o2.dict[j].Equals("<<")) {
                            o2.dict.Insert(j + 1, "/" + fontID);
                            o2.dict.Insert(j + 2, number.ToString());
                            o2.dict.Insert(j + 3, "0");
                            o2.dict.Insert(j + 4, "R");
                            return;
                        }
                    }
                }
            }
        }
    }


    private void InsertNewObject(
            List<String> dict, String[] list, String type) {
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals(type)) {
                dict.InsertRange(i + 2, list);
                return;
            }
        }
        if (dict[3].Equals("<<")) {
            dict.InsertRange(4, list);
            return;
        }
    }


    private void AddResource(
            String type, PDFobj obj, SortedDictionary<Int32, PDFobj> objects, Int32 objNumber) {
        String tag = type.Equals("/Font") ? "/F" : "/Im";
        String number = objNumber.ToString();
        String[] list = {tag + number, number, "0", "R"};
        for (int i = 0; i < obj.dict.Count; i++) {
            if (obj.dict[i].Equals(type)) {
                String token = obj.dict[i + 1];
                if (token.Equals("<<")) {
                    InsertNewObject(obj.dict, list, type);
                }
                else {
                    InsertNewObject(objects[Int32.Parse(token)].dict, list, type);
                }
                return;
            }
        }

        // Handle the case where the page originally does not have any font resources.
        String[] array = {type, "<<", tag + number, number, "0", "R", ">>"};
        for (int i = 0; i < obj.dict.Count; i++) {
            if (obj.dict[i].Equals("/Resources")) {
                obj.dict.InsertRange(i + 2, array);
                return;
            }
        }
        for (int i = 0; i < obj.dict.Count; i++) {
            if (obj.dict[i].Equals("<<")) {
                obj.dict.InsertRange(i + 1, array);
                return;
            }
        }
    }


    public void AddResource(Image image, SortedDictionary<Int32, PDFobj> objects) {
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Resources")) {
                String token = dict[i + 1];
                if (token.Equals("<<")) {       // Direct resources object
                    AddResource("/XObject", this, objects, image.objNumber);
                }
                else {                          // Indirect resources object
                    AddResource("/XObject", objects[Int32.Parse(token)], objects, image.objNumber);
                }
                return;
            }
        }
    }


    public void AddResource(Font font, SortedDictionary<Int32, PDFobj> objects) {
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Resources")) {
                String token = dict[i + 1];
                if (token.Equals("<<")) {       // Direct resources object
                    AddResource("/Font", this, objects, font.objNumber);
                }
                else {                          // Indirect resources object
                    AddResource("/Font", objects[Int32.Parse(token)], objects, font.objNumber);
                }
                return;
            }
        }
    }


    public void AddContent(
            byte[] content, SortedDictionary<Int32, PDFobj> objects) {
        PDFobj obj = new PDFobj();
        int maxObjNumber = -1;
        foreach (int number in objects.Keys) {
            if (number > maxObjNumber) { maxObjNumber = number; }
        }
        obj.SetNumber(maxObjNumber + 1);
        obj.SetStream(content);
        objects.Add(obj.GetNumber(), obj);

        String objNumber = obj.number.ToString();
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Contents")) {
                i += 1;
                String token = dict[i];
                if (token.Equals("[")) {
                    // Array of content objects
                    while (true) {
                        i += 1;
                        token = dict[i];
                        if (token.Equals("]")) {
                            dict.Insert(i, "R");
                            dict.Insert(i, "0");
                            dict.Insert(i, objNumber);
                            return;
                        }
                        i += 2;     // Skip the 0 and R
                    }
                }
                else {
                    // Single content object
                    PDFobj obj2 = objects[Int32.Parse(token)];
                    if (obj2.data == null && obj2.stream == null) {
                        // This is not a stream object!
                        for (int j = 0; j < obj2.dict.Count; j++) {
                            if (obj2.dict[j].Equals("]")) {
                                obj2.dict.Insert(j, "R");
                                obj2.dict.Insert(j, "0");
                                obj2.dict.Insert(j, objNumber);
                                return;
                            }
                        }
                    }
                    dict.Insert(i, "[");
                    dict.Insert(i + 4, "]");
                    dict.Insert(i + 4, "R");
                    dict.Insert(i + 4, "0");
                    dict.Insert(i + 4, objNumber);
                    return;
                }
            }
        }
    }


    /**
     * Adds new content object before the existing content objects.
     * The original code was provided by Stefan Ostermann author of ScribMaster and HandWrite Pro.
     * Additional code to handle PDFs with indirect array of stream objects was written by EDragoev.
     *
     * @param content
     * @param objects
     */
    public void AddPrefixContent(
            byte[] content, SortedDictionary<Int32, PDFobj> objects) {
        PDFobj obj = new PDFobj();
        int maxObjNumber = -1;
        foreach (int number in objects.Keys) {
            if (number > maxObjNumber) { maxObjNumber = number; }
        }
        obj.SetNumber(maxObjNumber + 1);
        obj.SetStream(content);
        objects[obj.GetNumber()] = obj;

        String objNumber = obj.number.ToString();
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Contents")) {
                i += 1;
                String token = dict[i];
                if (token.Equals("[")) {
                    // Array of content object streams
                    i += 1;
                    dict.Insert(i, "R");
                    dict.Insert(i, "0");
                    dict.Insert(i, objNumber);
                    return;
                }
                else {
                    // Single content object
                    PDFobj obj2 = objects[Int32.Parse(token)];
                    if (obj2.data == null && obj2.stream == null) {
                        // This is not a stream object!
                        for (int j = 0; j < obj2.dict.Count; j++) {
                            if (obj2.dict[j].Equals("[")) {
                                j += 1;
                                obj2.dict.Insert(j, "R");
                                obj2.dict.Insert(j, "0");
                                obj2.dict.Insert(j, objNumber);
                                return;
                            }
                        }
                    }
                    dict.Insert(i, "[");
                    dict.Insert(i + 4, "]");
                    i += 1;
                    dict.Insert(i, "R");
                    dict.Insert(i, "0");
                    dict.Insert(i, objNumber);
                    return;
                }
            }
        }
    }


    private int GetMaxGSNumber(PDFobj obj) {
        List<Int32> numbers = new List<Int32>();
        foreach (String token in obj.dict) {
            if (token.StartsWith("/GS")) {
                numbers.Add(Int32.Parse(token.Substring(3)));
            }
        }
        if (numbers.Count == 0) {
            return 0;
        }
        int maxGSNumber = -1;
        foreach (Int32 number in numbers) {
            if (number > maxGSNumber) {
                maxGSNumber = number;
            }
        }
        return maxGSNumber;
    }


    public void SetGraphicsState(
            GraphicsState gs, SortedDictionary<Int32, PDFobj> objects) {
        PDFobj obj = null;
        int index = -1;
        for (int i = 0; i < dict.Count; i++) {
            if (dict[i].Equals("/Resources")) {
                String token = dict[i + 1];
                if (token.Equals("<<")) {
                    obj = this;
                    index = i + 2;
                }
                else {
                    obj = objects[Int32.Parse(token)];
                    for (int j = 0; j < obj.dict.Count; j++) {
                        if (obj.dict[j].Equals("<<")) {
                            index = j + 1;
                            break;
                        }
                    }
                }
                break;
            }
        }

        gsNumber = GetMaxGSNumber(obj);
        if (gsNumber == 0) {                        // No existing ExtGState dictionary
            obj.dict.Insert(index, "/ExtGState");   // Add ExtGState dictionary
            obj.dict.Insert(++index, "<<");
        }
        else {
            while (index < obj.dict.Count) {
                String token = obj.dict[index];
                if (token.Equals("/ExtGState")) {
                    index += 1;
                    break;
                }
                index += 1;
            }
        }
        obj.dict.Insert(++index, "/GS" + (gsNumber + 1).ToString());
        obj.dict.Insert(++index, "<<");
        obj.dict.Insert(++index, "/CA");
        obj.dict.Insert(++index, gs.Get_CA().ToString());
        obj.dict.Insert(++index, "/ca");
        obj.dict.Insert(++index, gs.Get_ca().ToString());
        obj.dict.Insert(++index, ">>");
        if (gsNumber == 0) {
            obj.dict.Insert(++index, ">>");
        }

        StringBuilder buf = new StringBuilder();
        buf.Append("q\n");
        buf.Append("/GS" + (gsNumber + 1).ToString() + " gs\n");
        AddPrefixContent(Encoding.ASCII.GetBytes(buf.ToString()), objects);
    }

}
}   // End of namespace PDFjet.NET
