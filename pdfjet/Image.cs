/**
 *  Image.cs
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
using System.Collections.Generic;
using System.IO.Compression;


namespace PDFjet.NET {
/**
 *  Used to create image objects and draw them on a page.
 *  The image type can be one of the following: ImageType.JPG, ImageType.PNG, ImageType.BMP or ImageType.JET
 *
 *  Please see Example_03 and Example_24.
 */
public class Image : IDrawable {

    internal int objNumber;

    internal float x = 0f;  // Position of the image on the page
    internal float y = 0f;
    internal float w;       // Image width
    internal float h;       // Image height

    internal String uri;
    internal String key;

    private float box_x;
    private float box_y;

    private int degrees = 0;
    private bool flipUpsideDown = false;

    private String language = null;
    private String altDescription = Single.space;
    private String actualText = Single.space;


    /**
     *  The main constructor for the Image class.
     *
     *  @param pdf the page to draw this image on.
     *  @param inputStream the input stream to read the image from.
     *  @param imageType ImageType.JPG, ImageType.PNG or ImageType.BMP.
     *
     */
    public Image(PDF pdf, Stream inputStream, int imageType) {
        byte[] data;
        if (imageType == ImageType.JPG) {
            JPGImage jpg = new JPGImage(inputStream);
            data = jpg.GetData();
            w = jpg.GetWidth();
            h = jpg.GetHeight();
            if (jpg.GetColorComponents() == 1) {
                AddImage(pdf, data, null, imageType, "DeviceGray", 8);
            }
            else if (jpg.GetColorComponents() == 3) {
                AddImage(pdf, data, null, imageType, "DeviceRGB", 8);
            }
            else if (jpg.GetColorComponents() == 4) {
                AddImage(pdf, data, null, imageType, "DeviceCMYK", 8);
            }
        }
        else if (imageType == ImageType.PNG) {
            PNGImage png = new PNGImage(inputStream);
            data = png.GetData();
            w = png.GetWidth();
            h = png.GetHeight();
            if (png.GetColorType() == 0) {
                AddImage(pdf, data, null, imageType, "DeviceGray", png.GetBitDepth());
            }
            else {
                if (png.GetBitDepth() == 16) {
                    AddImage(pdf, data, null, imageType, "DeviceRGB", 16);
                }
                else {
                    AddImage(pdf, data, png.GetAlpha(), imageType, "DeviceRGB", 8);
                }
            }
        }
        else if (imageType == ImageType.BMP) {
            BMPImage bmp = new BMPImage(inputStream);
            data = bmp.GetData();
            w = bmp.GetWidth();
            h = bmp.GetHeight();
            AddImage(pdf, data, null, imageType, "DeviceRGB", 8);
        }
        else if (imageType == ImageType.JET) {
            AddImage(pdf, inputStream);
        }

        inputStream.Dispose();
    }


    public Image(PDF pdf, PDFobj obj) {
        // Console.WriteLine(obj.GetDict());
        w = float.Parse(obj.GetValue("/Width"));
        h = float.Parse(obj.GetValue("/Height"));
        pdf.Newobj();
        pdf.Append("<<\n");
        pdf.Append("/Type /XObject\n");
        pdf.Append("/Subtype /Image\n");
        pdf.Append("/Filter ");
        pdf.Append(obj.GetValue("/Filter"));
        pdf.Append("\n");
        pdf.Append("/Width ");
        pdf.Append(w);
        pdf.Append('\n');
        pdf.Append("/Height ");
        pdf.Append(h);
        pdf.Append('\n');
        String colorSpace = obj.GetValue("/ColorSpace");
        if (!colorSpace.Equals("")) {
            pdf.Append("/ColorSpace ");
            pdf.Append(colorSpace);
            pdf.Append("\n");
        }
        pdf.Append("/BitsPerComponent ");
        pdf.Append(obj.GetValue("/BitsPerComponent"));
        pdf.Append("\n");
        String decodeParms = obj.GetValue("/DecodeParms");
        if (!decodeParms.Equals("")) {
            pdf.Append("/DecodeParms ");
            pdf.Append(decodeParms);
            pdf.Append("\n");
        }
        String imageMask = obj.GetValue("/ImageMask");
        if (!imageMask.Equals("")) {
            pdf.Append("/ImageMask ");
            pdf.Append(imageMask);
            pdf.Append("\n");
        }
        pdf.Append("/Length ");
        pdf.Append(obj.stream.Length);
        pdf.Append('\n');
        pdf.Append(">>\n");
        pdf.Append("stream\n");
        pdf.Append(obj.stream, 0, obj.stream.Length);
        pdf.Append("\nendstream\n");
        pdf.Endobj();
        pdf.images.Add(this);
        objNumber = pdf.objNumber;
    }


    /**
     *  Sets the position of this image on the page to (x, y).
     *
     *  @param x the x coordinate of the top left corner of the image.
     *  @param y the y coordinate of the top left corner of the image.
     */
    public Image SetPosition(double x, double y) {
        return SetPosition((float) x, (float) y);
    }


    /**
     *  Sets the position of this image on the page to (x, y).
     *
     *  @param x the x coordinate of the top left corner of the image.
     *  @param y the y coordinate of the top left corner of the image.
     */
    public Image SetPosition(float x, float y) {
        return SetLocation(x, y);
    }


    /**
     *  Sets the location of this image on the page to (x, y).
     *
     *  @param x the x coordinate of the top left corner of the image.
     *  @param y the y coordinate of the top left corner of the image.
     */
    public Image SetLocation(float x, float y) {
        this.x = x;
        this.y = y;
        return this;
    }


    /**
     *  Scales this image by the specified factor.
     *
     *  @param factor the factor used to scale the image.
     */
    public Image ScaleBy(double factor) {
        return this.ScaleBy((float) factor, (float) factor);
    }


    /**
     *  Scales this image by the specified factor.
     *
     *  @param factor the factor used to scale the image.
     */
    public Image ScaleBy(float factor) {
        return this.ScaleBy(factor, factor);
    }


    /**
     *  Scales this image by the specified width and height factor.
     *  <p><i>Author:</i> <strong>Pieter Libin</strong>, pieter@emweb.be</p>
     *
     *  @param widthFactor the factor used to scale the width of the image
     *  @param heightFactor the factor used to scale the height of the image
     */
    public Image ScaleBy(float widthFactor, float heightFactor) {
        this.w *= widthFactor;
        this.h *= heightFactor;
        return this;
    }


    /**
     *  Places this image in the specified box.
     *
     *  @param box the specified box.
     */
    public void PlaceIn(Box box) {
        box_x = box.x;
        box_y = box.y;
    }


    /**
     *  Sets the URI for the "click box" action.
     *
     *  @param uri the URI
     */
    public void SetURIAction(String uri) {
        this.uri = uri;
    }


    /**
     *  Sets the destination key for the action.
     *
     *  @param key the destination name.
     */
    public void SetGoToAction(String key) {
        this.key = key;
    }


    /**
     *  Sets the rotate90 flag.
     *  When the flag is true the image is rotated 90 degrees clockwise.
     *
     *  @param rotate90 the flag.
     */
    public void SetRotateCW90(bool rotate90) {
        if (rotate90) {
            this.degrees = 90;
        }
        else {
            this.degrees = 0;
        }
    }


    /**
     *  Sets the image rotation to the specified number of degrees.
     *
     *  @param degrees the number of degrees.
     */
    public void SetRotate(int degrees) {
        this.degrees = degrees;
    }


    /**
     *  Sets the alternate description of this image.
     *
     *  @param altDescription the alternate description of the image.
     *  @return this Image.
     */
    public Image SetAltDescription(String altDescription) {
        this.altDescription = altDescription;
        return this;
    }


    /**
     *  Sets the actual text for this image.
     *
     *  @param actualText the actual text for the image.
     *  @return this Image.
     */
    public Image SetActualText(String actualText) {
        this.actualText = actualText;
        return this;
    }


    /**
     *  Draws this image on the specified page.
     *
     *  @param page the page to draw on.
     *  @return x and y coordinates of the bottom right corner of this component.
     *  @throws Exception
     */
    public float[] DrawOn(Page page) {
        page.AddBMC(StructElem.SPAN, language, altDescription, actualText);

        x += box_x;
        y += box_y;
        page.Append("q\n");

        if (degrees == 0) {
            page.Append(w);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(h);
            page.Append(' ');
            page.Append(x);
            page.Append(' ');
            page.Append(page.height - (y + h));
            page.Append(" cm\n");
        }
        else if (degrees == 90) {
            page.Append(h);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(w);
            page.Append(' ');
            page.Append(x);
            page.Append(' ');
            page.Append(page.height - y);
            page.Append(" cm\n");
            page.Append("0 -1 1 0 0 0 cm\n");
        }
        else if (degrees == 180) {
            page.Append(w);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(h);
            page.Append(' ');
            page.Append(x + w);
            page.Append(' ');
            page.Append(page.height - y);
            page.Append(" cm\n");
            page.Append("-1 0 0 -1 0 0 cm\n");
        }
        else if (degrees == 270) {
            page.Append(h);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(0f);
            page.Append(' ');
            page.Append(w);
            page.Append(' ');
            page.Append(x + h);
            page.Append(' ');
            page.Append(page.height - (y + w));
            page.Append(" cm\n");
            page.Append("0 1 -1 0 0 0 cm\n");
        }

        if (flipUpsideDown) {
            page.Append("1 0 0 -1 0 0 cm\n");
        }

        page.Append("/Im");
        page.Append(objNumber);
        page.Append(" Do\n");
        page.Append("Q\n");

        page.AddEMC();

        if (uri != null || key != null) {
            page.AddAnnotation(new Annotation(
                    uri,
                    key,    // The destination name
                    x,
                    page.height - y,
                    x + w,
                    page.height - (y + h),
                    language,
                    altDescription,
                    actualText));
        }

        return new float[] {x + w, y + h};
    }


    /**
     *  Returns the width of this image when drawn on the page.
     *  The scaling is take into account.
     *
     *  @return w - the width of this image.
     */
    public float GetWidth() {
        return this.w;
    }


    /**
     *  Returns the height of this image when drawn on the page.
     *  The scaling is take into account.
     *
     *  @return h - the height of this image.
     */
    public float GetHeight() {
        return this.h;
    }


    private void AddSoftMask(
            PDF pdf,
            byte[] data,
            String colorSpace,
            int bitsPerComponent) {
        pdf.Newobj();
        pdf.Append("<<\n");
        pdf.Append("/Type /XObject\n");
        pdf.Append("/Subtype /Image\n");
        pdf.Append("/Filter /FlateDecode\n");
        pdf.Append("/Width ");
        pdf.Append((int) w);
        pdf.Append('\n');
        pdf.Append("/Height ");
        pdf.Append((int) h);
        pdf.Append('\n');
        pdf.Append("/ColorSpace /");
        pdf.Append(colorSpace);
        pdf.Append('\n');
        pdf.Append("/BitsPerComponent ");
        pdf.Append(bitsPerComponent);
        pdf.Append('\n');
        pdf.Append("/Length ");
        pdf.Append(data.Length);
        pdf.Append('\n');
        pdf.Append(">>\n");
        pdf.Append("stream\n");
        pdf.Append(data, 0, data.Length);
        pdf.Append("\nendstream\n");
        pdf.Endobj();
        objNumber = pdf.objNumber;
    }


    private void AddImage(
            PDF pdf,
            byte[] data,
            byte[] alpha,
            int imageType,
            String colorSpace,
            int bitsPerComponent) {
        if (alpha != null) {
            AddSoftMask(pdf, alpha, "DeviceGray", 8);
        }
        pdf.Newobj();
        pdf.Append("<<\n");
        pdf.Append("/Type /XObject\n");
        pdf.Append("/Subtype /Image\n");
        if (imageType == ImageType.JPG) {
            pdf.Append("/Filter /DCTDecode\n");
        }
        else if (imageType == ImageType.PNG || imageType == ImageType.BMP) {
            pdf.Append("/Filter /FlateDecode\n");
            if (alpha != null) {
                pdf.Append("/SMask ");
                pdf.Append(objNumber);
                pdf.Append(" 0 R\n");
            }
        }
        pdf.Append("/Width ");
        pdf.Append((int) w);
        pdf.Append('\n');
        pdf.Append("/Height ");
        pdf.Append((int) h);
        pdf.Append('\n');
        pdf.Append("/ColorSpace /");
        pdf.Append(colorSpace);
        pdf.Append('\n');
        pdf.Append("/BitsPerComponent ");
        pdf.Append(bitsPerComponent);
        pdf.Append('\n');
        if (colorSpace.Equals("DeviceCMYK")) {
            // If the image was created with Photoshop - invert the colors:
            pdf.Append("/Decode [1.0 0.0 1.0 0.0 1.0 0.0 1.0 0.0]\n");
        }
        pdf.Append("/Length ");
        pdf.Append(data.Length);
        pdf.Append('\n');
        pdf.Append(">>\n");
        pdf.Append("stream\n");
        pdf.Append(data, 0, data.Length);
        pdf.Append("\nendstream\n");
        pdf.Endobj();
        pdf.images.Add(this);
        objNumber = pdf.objNumber;
    }


    private void AddImage(PDF pdf, Stream inputStream) {

        w = GetInt(inputStream);                // Width
        h = GetInt(inputStream);                // Height
        byte c = (byte) inputStream.ReadByte(); // Color Space
        byte a = (byte) inputStream.ReadByte(); // Alpha

        if (a != 0) {
            pdf.Newobj();
            pdf.Append("<<\n");
            pdf.Append("/Type /XObject\n");
            pdf.Append("/Subtype /Image\n");
            pdf.Append("/Filter /FlateDecode\n");
            pdf.Append("/Width ");
            pdf.Append(w);
            pdf.Append('\n');
            pdf.Append("/Height ");
            pdf.Append(h);
            pdf.Append('\n');
            pdf.Append("/ColorSpace /DeviceGray\n");
            pdf.Append("/BitsPerComponent 8\n");
            int length = GetInt(inputStream);
            pdf.Append("/Length ");
            pdf.Append(length);
            pdf.Append('\n');
            pdf.Append(">>\n");
            pdf.Append("stream\n");
            byte[] buf1 = new byte[length];
            inputStream.Read(buf1, 0, length);
            pdf.Append(buf1, 0, length);
            pdf.Append("\nendstream\n");
            pdf.Endobj();
            objNumber = pdf.objNumber;
        }

        pdf.Newobj();
        pdf.Append("<<\n");
        pdf.Append("/Type /XObject\n");
        pdf.Append("/Subtype /Image\n");
        pdf.Append("/Filter /FlateDecode\n");
        if (a != 0) {
            pdf.Append("/SMask ");
            pdf.Append(objNumber);
            pdf.Append(" 0 R\n");
        }
        pdf.Append("/Width ");
        pdf.Append(w);
        pdf.Append('\n');
        pdf.Append("/Height ");
        pdf.Append(h);
        pdf.Append('\n');
        pdf.Append("/ColorSpace /");
        if (c == 1) {
            pdf.Append("DeviceGray");
        }
        else if (c == 3 || c == 6) {
            pdf.Append("DeviceRGB");
        }
        pdf.Append('\n');
        pdf.Append("/BitsPerComponent 8\n");
        pdf.Append("/Length ");
        pdf.Append(GetInt(inputStream));
        pdf.Append('\n');
        pdf.Append(">>\n");
        pdf.Append("stream\n");
        byte[] buf2 = new byte[2048];
        int count;
        while ((count = inputStream.Read(buf2, 0, buf2.Length)) > 0) {
            pdf.Append(buf2, 0, count);
        }
        pdf.Append("\nendstream\n");
        pdf.Endobj();
        pdf.images.Add(this);
        objNumber = pdf.objNumber;
    }


    private int GetInt(Stream inputStream) {
        byte[] buf = new byte[4];
        inputStream.Read(buf, 0, 4);
        int val = 0;
        val |= buf[0] & 0xff;
        val <<= 8;
        val |= buf[1] & 0xff;
        val <<= 8;
        val |= buf[2] & 0xff;
        val <<= 8;
        val |= buf[3] & 0xff;
        return val;
    }


    /**
     *  Constructor used to attach images to existing PDF.
     *
     *  @param pdf the page to draw this image on.
     *  @param inputStream the input stream to read the image from.
     *  @param imageType ImageType.JPG, ImageType.PNG and ImageType.BMP.
     *
     */
    public Image(SortedDictionary<Int32, PDFobj> objects, Stream inputStream, int imageType) {
        byte[] data;
        if (imageType == ImageType.JPG) {
            JPGImage jpg = new JPGImage(inputStream);
            data = jpg.GetData();
            w = jpg.GetWidth();
            h = jpg.GetHeight();
            if (jpg.GetColorComponents() == 1) {
                AddImage(objects, data, null, imageType, "DeviceGray", 8);
            }
            else if (jpg.GetColorComponents() == 3) {
                AddImage(objects, data, null, imageType, "DeviceRGB", 8);
            }
            else if (jpg.GetColorComponents() == 4) {
                AddImage(objects, data, null, imageType, "DeviceCMYK", 8);
            }
        }
        else if (imageType == ImageType.PNG) {
            PNGImage png = new PNGImage(inputStream);
            data = png.GetData();
            w = png.GetWidth();
            h = png.GetHeight();
            if (png.GetColorType() == 0) {
                AddImage(objects, data, null, imageType, "DeviceGray", png.GetBitDepth());
            }
            else {
                if (png.GetBitDepth() == 16) {
                    AddImage(objects, data, null, imageType, "DeviceRGB", 16);
                }
                else {
                    AddImage(objects, data, png.GetAlpha(), imageType, "DeviceRGB", 8);
                }
            }
        }
        else if (imageType == ImageType.BMP) {
            BMPImage bmp = new BMPImage(inputStream);
            data = bmp.GetData();
            w = bmp.GetWidth();
            h = bmp.GetHeight();
            AddImage(objects, data, null, imageType, "DeviceRGB", 8);
        }
/*
        else if (imageType == ImageType.JET) {
            AddImage(pdf, inputStream);
        }
*/
        inputStream.Close();
    }


    private void AddSoftMask(
            SortedDictionary<Int32, PDFobj> objects,
            byte[] data,
            String colorSpace,
            int bitsPerComponent) {
        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Type");
        dict.Add("/XObject");
        dict.Add("/Subtype");
        dict.Add("/Image");
        dict.Add("/Filter");
        dict.Add("/FlateDecode");
        dict.Add("/Width");
        dict.Add(((int) w).ToString());
        dict.Add("/Height");
        dict.Add(((int) h).ToString());
        dict.Add("/ColorSpace");
        dict.Add("/" + colorSpace);
        dict.Add("/BitsPerComponent");
        dict.Add(bitsPerComponent.ToString());
        dict.Add("/Length");
        dict.Add(data.Length.ToString());
        dict.Add(">>");
        obj.SetStream(data);
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        objNumber = obj.number;
    }


    private void AddImage(
            SortedDictionary<Int32, PDFobj> objects,
            byte[] data,
            byte[] alpha,
            int imageType,
            String colorSpace,
            int bitsPerComponent) {
        if (alpha != null) {
            AddSoftMask(objects, alpha, "DeviceGray", 8);
        }
        PDFobj obj = new PDFobj();
        List<String> dict = obj.GetDict();
        dict.Add("<<");
        dict.Add("/Type");
        dict.Add("/XObject");
        dict.Add("/Subtype");
        dict.Add("/Image");
        if (imageType == ImageType.JPG) {
            dict.Add("/Filter");
            dict.Add("/DCTDecode");
        }
        else if (imageType == ImageType.PNG || imageType == ImageType.BMP) {
            dict.Add("/Filter");
            dict.Add("/FlateDecode");
            if (alpha != null) {
                dict.Add("/SMask");
                dict.Add(objNumber.ToString());
                dict.Add("0");
                dict.Add("R");
            }
        }
        dict.Add("/Width");
        dict.Add(((int) w).ToString());
        dict.Add("/Height");
        dict.Add(((int) h).ToString());
        dict.Add("/ColorSpace");
        dict.Add("/" + colorSpace);
        dict.Add("/BitsPerComponent");
        dict.Add(bitsPerComponent.ToString());
        if (colorSpace.Equals("DeviceCMYK")) {
            // If the image was created with Photoshop - invert the colors:
            dict.Add("/Decode");
            dict.Add("[");
            dict.Add("1.0");
            dict.Add("0.0");
            dict.Add("1.0");
            dict.Add("0.0");
            dict.Add("1.0");
            dict.Add("0.0");
            dict.Add("1.0");
            dict.Add("0.0");
            dict.Add("]");
        }
        dict.Add("/Length");
        dict.Add(data.Length.ToString());
        dict.Add(">>");
        obj.SetStream(data);
        obj.number = MaxKey(objects.Keys) + 1;
        objects.Add(obj.number, obj);
        objNumber = obj.number;
    }


    private Int32 MaxKey(SortedDictionary<Int32, PDFobj>.KeyCollection keys) {
        Int32 maxKey = 0;
        foreach (Int32 key in keys) {
            if (key > maxKey) {
                maxKey = key;
            }
        }
        return maxKey;
    }


    public void ResizeToFit(Page page, bool keepAspectRatio) {
        float page_w = page.GetWidth();
        float page_h = page.GetHeight();
        float image_w = this.GetWidth();
        float image_h = this.GetHeight();
        if (keepAspectRatio) {
            this.ScaleBy(Math.Min((page_w - x)/image_w, (page_h - y)/image_h));
        }
        else {
            this.ScaleBy((page_w - x)/image_w, (page_h - y)/image_h);
        }
    }


    public void FlipUpsideDown(bool flipUpsideDown) {
        this.flipUpsideDown = flipUpsideDown;
    }

}   // End of Image.cs
}   // End of namespace PDFjet.NET
