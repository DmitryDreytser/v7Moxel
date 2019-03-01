/**
 *  TextFrame.cs
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
using System.Text.RegularExpressions;
using System.Collections.Generic;


/**
 *  Please see Example_47
 *
 */
namespace PDFjet.NET {
public class TextFrame {

    private List<Paragraph> paragraphs;
    private Font font;
    private Font fallbackFont;
    private float x;
    private float y;
    private float w;
    private float h;
    private float x_text;
    private float y_text;
    private float leading;
    private float paragraphLeading;
    private List<float[]> beginParagraphPoints;
    private List<float[]> endParagraphPoints;
    private float spaceBetweenTextLines;


    public TextFrame(List<Paragraph> paragraphs) {
        if (paragraphs != null) {
            this.paragraphs = paragraphs;
            this.font = paragraphs[0].list[0].GetFont();
            this.fallbackFont = paragraphs[0].list[0].GetFallbackFont();
            this.leading = font.GetBodyHeight();
            this.paragraphLeading = 2*leading;
            this.beginParagraphPoints = new List<float[]>();
            this.endParagraphPoints = new List<float[]>();
            this.spaceBetweenTextLines = font.StringWidth(fallbackFont, Single.space);
        }
    }


    public TextFrame SetLocation(float x, float y) {
        this.x = x;
        this.y = y;
        return this;
    }


    public TextFrame SetWidth(float w) {
        this.w = w;
        return this;
    }


    public TextFrame SetHeight(float h) {
        this.h = h;
        return this;
    }


    public TextFrame SetLeading(float leading) {
        this.leading = leading;
        return this;
    }


    public TextFrame SetParagraphLeading(float paragraphLeading) {
        this.paragraphLeading = paragraphLeading;
        return this;
    }


    public List<float[]> GetBeginParagraphPoints() {
        return this.beginParagraphPoints;
    }


    public List<float[]> GetEndParagraphPoints() {
        return this.endParagraphPoints;
    }


    public TextFrame SetSpaceBetweenTextLines(float spaceBetweenTextLines) {
        this.spaceBetweenTextLines = spaceBetweenTextLines;
        return this;
    }


    public List<Paragraph> GetParagraphs() {
        return this.paragraphs;
    }


    public TextFrame DrawOn(Page page) {
        return DrawOn(page, true);
    }


    public TextFrame DrawOn(Page page, bool draw) {
        this.x_text = x;
        this.y_text = y + font.GetAscent();

        Paragraph paragraph = null;
        for (int i = 0; i < paragraphs.Count; i++) {
            paragraph = paragraphs[i];

            StringBuilder buf = new StringBuilder();
            foreach (TextLine textLine in paragraph.list) {
                buf.Append(textLine.GetText());
                buf.Append(Single.space);
            }

            int numOfTextLines = paragraph.list.Count;
            for (int j = 0; j < numOfTextLines; j++) {
                TextLine textLine = paragraph.list[j];
                if (j == 0) {
                    beginParagraphPoints.Add(new float[] { x_text, y_text });
                }
                textLine.SetAltDescription((i == 0) ? buf.ToString() : Single.space);
                textLine.SetActualText((i == 0) ? buf.ToString() : Single.space);

                TextLine textLine2 = DrawTextLine(page, x_text, y_text, textLine, draw);
                if (!textLine2.GetText().Equals("")) {
                    List<Paragraph> theRest = new List<Paragraph>();
                    Paragraph paragraph2 = new Paragraph(textLine2);
                    j++;
                    while (j < numOfTextLines) {
                        paragraph2.Add(paragraph.list[j]);
                        j++;
                    }
                    theRest.Add(paragraph2);
                    i++;
                    while (i < paragraphs.Count) {
                        theRest.Add(paragraphs[i]);
                        i++;
                    }
                    return new TextFrame(theRest);
                }

                if (j == (numOfTextLines - 1)) {
                    endParagraphPoints.Add(new float[] { textLine2.x, textLine2.y });
                }
                x_text = textLine2.x;
                if (textLine.GetTrailingSpace()) {
                    x_text += spaceBetweenTextLines;
                }
                y_text = textLine2.y;
            }
            x_text = x;
            y_text += paragraphLeading;
        }

        TextFrame textFrame = new TextFrame(null);
        textFrame.SetLocation(x_text, y_text + font.GetDescent());
        return textFrame;
    }


    public TextLine DrawTextLine(
            Page page, float x_text, float y_text, TextLine textLine, bool draw) {

        TextLine textLine2 = null;
        Font font = textLine.GetFont();
        Font fallbackFont = textLine.GetFallbackFont();
        int color = textLine.GetColor();

        StringBuilder buf = new StringBuilder();
        String[] tokens = Regex.Split(textLine.GetText(), @"\s+");
        bool firstTextSegment = true;
        for (int i = 0; i < tokens.Length; i++) {
            String token = (i == 0) ? tokens[i] : (Single.space + tokens[i]);
            if (font.StringWidth(fallbackFont, token) < (this.w - (x_text - x))) {
                buf.Append(token);
                x_text += font.StringWidth(fallbackFont, token);
            }
            else {
                if (draw) {
                    new TextLine(font, buf.ToString())
                            .SetFallbackFont(textLine.GetFallbackFont())
                            .SetLocation(x_text - font.StringWidth(fallbackFont, buf.ToString()),
                                    y_text + textLine.GetVerticalOffset())
                            .SetColor(color)
                            .SetUnderline(textLine.GetUnderline())
                            .SetStrikeout(textLine.GetStrikeout())
                            .SetLanguage(textLine.GetLanguage())
                            .SetAltDescription(firstTextSegment ? textLine.GetAltDescription() : Single.space)
                            .SetActualText(firstTextSegment ? textLine.GetActualText() : Single.space)
                            .DrawOn(page);
                    firstTextSegment = false;
                }
                x_text = x + font.StringWidth(fallbackFont, tokens[i]);
                y_text += leading;
                buf.Length = 0;
                buf.Append(tokens[i]);

                if (y_text + font.GetDescent() > (y + h)) {
                    i++;
                    while (i < tokens.Length) {
                        buf.Append(Single.space);
                        buf.Append(tokens[i]);
                        i++;
                    }
                    textLine2 = new TextLine(font, buf.ToString());
                    textLine2.SetLocation(x, y_text);
                    return textLine2;
                }
            }
        }
        if (draw) {
            new TextLine(font, buf.ToString())
                    .SetFallbackFont(textLine.GetFallbackFont())
                    .SetLocation(x_text - font.StringWidth(fallbackFont, buf.ToString()),
                            y_text + textLine.GetVerticalOffset())
                    .SetColor(color)
                    .SetUnderline(textLine.GetUnderline())
                    .SetStrikeout(textLine.GetStrikeout())
                    .SetLanguage(textLine.GetLanguage())
                    .SetAltDescription(firstTextSegment ? textLine.GetAltDescription() : Single.space)
                    .SetActualText(firstTextSegment ? textLine.GetActualText() : Single.space)
                    .DrawOn(page);
            firstTextSegment = false;
        }

        textLine2 = new TextLine(font, "");
        textLine2.SetLocation(x_text, y_text);
        return textLine2;
    }

}   // End of TextFrame.cs
}   // End of namespace PDFjet.NET
