/**
 *  TextLine.cs
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
using System.Text;


namespace PDFjet.NET {
/**
 *  Used to create text line objects.
 *
 *
 */
public class TextLine : IDrawable {

    internal float x;
    internal float y;

    internal Font font;
    internal Font fallbackFont;
    internal String str;
    internal bool trailingSpace = true;

    private String uri;
    private String key;

    private bool underline = false;
    private bool strikeout = false;
    private String underlineTTS = "underline";
    private String strikeoutTTS = "strikeout";

    private int degrees = 0;
    private int color = Color.black;

    private float box_x;
    private float box_y;
    
    private int textEffect = Effect.NORMAL;
    private float verticalOffset = 0f;

    private String language = null;
    private String altDescription = null;
    private String actualText = null;

    private String uriLanguage = null;
    private String uriAltDescription = null;
    private String uriActualText = null;


    /**
     *  Constructor for creating text line objects.
     *
     *  @param font the font to use.
     */
    public TextLine(Font font) {
        this.font = font;
    }


    /**
     *  Constructor for creating text line objects.
     *
     *  @param font the font to use.
     *  @param text the text.
     */
    public TextLine(Font font, String text) {
        this.font = font;
        this.str = text;
        if (this.altDescription == null) {
            this.altDescription = text;
        }
        if (this.actualText == null) {
            this.actualText = text;
        }
    }


    /**
     *  Sets the text.
     *
     *  @param text the text.
     *  @return this TextLine.
     */
    public TextLine SetText(String text) {
        this.str = text;
        if (this.altDescription == null) {
            this.altDescription = text;
        }
        if (this.actualText == null) {
            this.actualText = text;
        }
        return this;
    }


    /**
     *  Returns the text.
     *
     *  @return the text.
     */
    public String GetText() {
        return str;
    }


    /**
     *  Sets the position where this text line will be drawn on the page.
     *
     *  @param x the x coordinate of the text line.
     *  @param y the y coordinate of the text line.
     *  @return this TextLine.
     */
    public TextLine SetPosition(double x, double y) {
        return SetPosition((float) x, (float) y);
    }


    /**
     *  Sets the position where this text line will be drawn on the page.
     *
     *  @param x the x coordinate of the text line.
     *  @param y the y coordinate of the text line.
     *  @return this TextLine.
     */
    public TextLine SetPosition(float x, float y) {
        return SetLocation(x, y);
    }


    /**
     *  Sets the location where this text line will be drawn on the page.
     *
     *  @param x the x coordinate of the text line.
     *  @param y the y coordinate of the text line.
     *  @return this TextLine.
     */
    public TextLine SetLocation(float x, float y) {
        this.x = x;
        this.y = y;
        return this;
    }


    /**
     *  Sets the font to use for this text line.
     *
     *  @param font the font to use.
     *  @return this TextLine.
     */
    public TextLine SetFont(Font font) {
        this.font = font;
        return this;
    }


    /**
     *  Gets the font to use for this text line.
     *
     *  @return font the font to use.
     */
    public Font GetFont() {
        return font;
    }


    /**
     *  Sets the font size to use for this text line.
     *
     *  @param fontSize the fontSize to use.
     *  @return this TextLine.
     */
    public TextLine SetFontSize(float fontSize) {
        this.font.SetSize(fontSize);
        return this;
    }


    /**
     *  Sets the fallback font.
     *
     *  @param fallbackFont the fallback font.
     *  @return this TextLine.
     */
    public TextLine SetFallbackFont(Font fallbackFont) {
        this.fallbackFont = fallbackFont;
        return this;
    }


    /**
     *  Returns the fallback font.
     *
     *  @return the fallback font.
     */
    public Font GetFallbackFont() {
        return this.fallbackFont;
    }


    /**
     *  Sets the color for this text line.
     *
     *  @param color the color specified as an integer.
     *  @return this TextLine.
     */
    public TextLine SetColor(int color) {
        this.color = color;
        return this;
    }


    /**
     *  Returns the text line color.
     *
     *  @return the text line color.
     */
    public int GetColor() {
        return color;
    }


    /**
     * Returns the y coordinate of the destination.
     * 
     * @return the y coordinate of the destination.
     */
    public float GetDestinationY() {
        return y - font.GetSize();
    }


    /**
     * Returns the y coordinate of the destination.
     * 
     * @return the y coordinate of the destination.
     */
    public float GetY() {
        return GetDestinationY();
    }


    /**
     *  Returns the width of this TextLine.
     *
     *  @return the width.
     */
    public float GetWidth() {
        if (fallbackFont == null) {
            return font.StringWidth(str);
        }
        return font.StringWidth(fallbackFont, str);
    }


    /**
     *  Returns the height of this TextLine.
     *
     *  @return the height.
     */
    public double GetHeight() {
        return font.GetHeight();
    }


    /**
     *  Sets the URI for the "click text line" action.
     *
     *  @param uri the URI
     *  @return this TextLine.
     */
    public TextLine SetURIAction(String uri) {
        this.uri = uri;
        return this;
    }


    /**
     *  Returns the action URI.
     * 
     *  @return the action URI.
     */
    public String GetURIAction() {
        return this.uri;
    }


    /**
     *  Sets the destination key for the action.
     *
     *  @param key the destination name.
     *  @return this TextLine.
     */
    public TextLine SetGoToAction(String key) {
        this.key = key;
        return this;
    }


    /**
     * Returns the GoTo action string.
     * 
     * @return the GoTo action string.
     */
    public String GetGoToAction() {
        return this.key;
    }


    /**
     *  Sets the underline variable.
     *  If the value of the underline variable is 'true' - the text is underlined.
     *
     *  @param underline the underline flag.
     *  @return this TextLine.
     */
    public TextLine SetUnderline(bool underline) {
        this.underline = underline;
        return this;
    }


    /**
     *  Returns the underline flag.
     * 
     *  @return the underline flag.
     */
    public bool GetUnderline() {
        return this.underline;
    }


    /**
     *  Sets the strike variable.
     *  If the value of the strike variable is 'true' - a strike line is drawn through the text.
     *
     *  @param strike the strike value.
     *  @return this TextLine.
     */
    public TextLine SetStrikeout(bool strike) {
        this.strikeout = strike;
        return this;
    }


    /**
     *  Returns the strikeout flag.
     * 
     *  @return the strikeout flag.
     */
    public bool GetStrikeout() {
        return this.strikeout;
    }


    /**
     *  Sets the direction in which to draw the text.
     *
     *  @param degrees the number of degrees.
     *  @return this TextLine.
     */
    public TextLine SetTextDirection(int degrees) {
        this.degrees = degrees;
        return this;
    }


    /**
     * Returns the text direction.
     * 
     * @return the text direction.
     */
    public int GetTextDirection() {
        return degrees;
    }


    /**
     *  Sets the text effect.
     * 
     *  @param textEffect Effect.NORMAL, Effect.SUBSCRIPT or Effect.SUPERSCRIPT.
     *  @return this TextLine.
     */
    public TextLine SetTextEffect(int textEffect) {
        this.textEffect = textEffect;
        return this;
    }


    /**
     *  Returns the text effect.
     * 
     *  @return the text effect.
     */
    public int GetTextEffect() {
        return textEffect;
    }


    /**
     *  Sets the vertical offset of the text.
     * 
     *  @param verticalOffset the vertical offset.
     *  @return this TextLine.
     */
    public TextLine SetVerticalOffset(float verticalOffset) {
        this.verticalOffset = verticalOffset;
        return this;
    }


    /**
     *  Returns the vertical text offset.
     * 
     *  @return the vertical text offset.
     */
    public float GetVerticalOffset() {
        return verticalOffset;
    }


    /**
     *  Sets the trailing space after this text line when used in paragraph.
     * 
     *  @param trailingSpace the trailing space.
     *  @return this TextLine.
     */
    public TextLine SetTrailingSpace(bool trailingSpace) {
        this.trailingSpace = trailingSpace;
        return this;
    }


    /**
     *  Returns the trailing space.
     * 
     *  @return the trailing space.
     */
    public bool GetTrailingSpace() {
        return trailingSpace;
    }


    public TextLine SetLanguage(String language) {
        this.language = language;
        return this;
    }


    public String GetLanguage() {
        return this.language;
    }


    /**
     *  Sets the alternate description of this text line.
     *
     *  @param altDescription the alternate description of the text line.
     *  @return this TextLine.
     */
    public TextLine SetAltDescription(String altDescription) {
        this.altDescription = altDescription;
        return this;
    }


    public String GetAltDescription() {
        return altDescription;
    }


    /**
     *  Sets the actual text for this text line.
     *
     *  @param actualText the actual text for the text line.
     *  @return this TextLine.
     */
    public TextLine SetActualText(String actualText) {
        this.actualText = actualText;
        return this;
    }


    public String GetActualText() {
        return actualText;
    }


    public TextLine SetURILanguage(String uriLanguage) {
        this.uriLanguage = uriLanguage;
        return this;
    }


    public TextLine SetURIAltDescription(String uriAltDescription) {
        this.uriAltDescription = uriAltDescription;
        return this;
    }


    public TextLine SetURIActualText(String uriActualText) {
        this.uriActualText = uriActualText;
        return this;
    }


    /**
     *  Places this text line in the specified box at position (0.0, 0.0).
     *
     *  @param box the specified box.
     *  @return this TextLine.
     */
    public TextLine PlaceIn(Box box) {
        PlaceIn(box, 0.0, 0.0);
        return this;
    }


    /**
     *  Places this text line in the box at the specified offset.
     *
     *  @param box the specified box.
     *  @param x_offset the x offset from the top left corner of the box.
     *  @param y_offset the y offset from the top left corner of the box.
     *  @return this TextLine.
     */
    public TextLine PlaceIn(
            Box box,
            double x_offset,
            double y_offset) {
        return PlaceIn(box, (float) x_offset, (float) y_offset);
    }


    /**
     *  Places this text line in the box at the specified offset.
     *
     *  @param box the specified box.
     *  @param x_offset the x offset from the top left corner of the box.
     *  @param y_offset the y offset from the top left corner of the box.
     *  @return this TextLine.
     */
    public TextLine PlaceIn(
            Box box,
            float x_offset,
            float y_offset) {
        box_x = box.x + x_offset;
        box_y = box.y + y_offset;
        return this;
    }


    /**
     *  Draws this text line on the specified page.
     *
     *  @param page the page to draw this text line on.
     */
    public float[] DrawOn(Page page) {
        return DrawOn(page, true);
    }


    /**
     *  Draws this text line on the specified page if the draw parameter is true.
     *
     *  @param page the page to draw this text line on.
     *  @param draw if draw is false - no action is performed.
     *  @return x and y coordinates of the bottom right corner of this component.
     *  @throws Exception
     */
    internal float[] DrawOn(Page page, bool draw) {
        if (page == null || !draw || str == null || str.Equals("")) {
            return new float[] {x, y};
        }

        page.SetTextDirection(degrees);

        x += box_x;
        y += box_y;

        page.SetBrushColor(color);
        page.AddBMC(StructElem.SPAN, language, altDescription, actualText);
        if (fallbackFont == null) {
            page.DrawString(font, str, x, y);
        }
        else {
            page.DrawString(font, fallbackFont, str, x, y);
        }
        page.AddEMC();

        double radians = Math.PI * degrees / 180.0;
        if (underline) {
            page.SetPenWidth(font.underlineThickness);
            page.SetPenColor(color);
            double lineLength = font.StringWidth(str);
            double x_adjust = font.underlinePosition * Math.Sin(radians);
            double y_adjust = font.underlinePosition * Math.Cos(radians);
            double x2 = x + lineLength * Math.Cos(radians);
            double y2 = y - lineLength * Math.Sin(radians);
            page.AddBMC(StructElem.SPAN, language, underlineTTS, underlineTTS);
            page.MoveTo(x + x_adjust, y + y_adjust);
            page.LineTo(x2 + x_adjust, y2 + y_adjust);
            page.StrokePath();
            page.AddEMC();
        }

        if (strikeout) {
            page.SetPenWidth(font.underlineThickness);
            page.SetPenColor(color);
            double lineLength = font.StringWidth(str);
            double x_adjust = ( font.body_height / 4.0 ) * Math.Sin(radians);
            double y_adjust = ( font.body_height / 4.0 ) * Math.Cos(radians);
            double x2 = x + lineLength * Math.Cos(radians);
            double y2 = y - lineLength * Math.Sin(radians);
            page.AddBMC(StructElem.SPAN, language, strikeoutTTS, strikeoutTTS);
            page.MoveTo(x - x_adjust, y - y_adjust);
            page.LineTo(x2 - x_adjust, y2 - y_adjust);
            page.StrokePath();
            page.AddEMC();
        }

        if (uri != null || key != null) {
            page.AddAnnotation(new Annotation(
                    uri,
                    key,    // The destination name
                    x,
                    page.height - (y - font.ascent),
                    x + font.StringWidth(str),
                    page.height - (y - font.descent),
                    uriLanguage,
                    uriAltDescription,
                    uriActualText));
        }
        page.SetTextDirection(0);

        float len = font.StringWidth(str);
        double x_max = Math.Max((double) x, x + len*Math.Cos(radians));
        double y_max = Math.Max((double) y, y - len*Math.Sin(radians));

        return new float[] {(float) x_max, (float) y_max};
    }

}   // End of TextLine.cs
}   // End of namespace PDFjet.NET
