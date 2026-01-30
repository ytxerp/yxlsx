/*
 * MIT License
 *
 * Copyright (c) 2024 YTX
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#include "style.h"

#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_YXLSX

Style::Style(OperationMode mode)
    : AbstractOOXmlFile { mode }
{
}

void Style::ComposeXml(QIODevice* device) const
{
    QXmlStreamWriter writer(device);
    writer.writeStartDocument();

    writer.writeStartElement(QLatin1String("styleSheet"));
    writer.writeDefaultNamespace(QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));

    // -------------------
    // Fonts
    // -------------------
    writer.writeStartElement(QLatin1String("fonts"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("2"));

    // Default font (11pt)
    writer.writeStartElement(QLatin1String("font"));
    writer.writeTextElement(QLatin1String("sz"), QLatin1String("11"));
    writer.writeTextElement(QLatin1String("name"), QLatin1String("Calibri"));
    writer.writeEndElement(); // font

    // Small font (8pt)
    writer.writeStartElement(QLatin1String("font"));
    writer.writeTextElement(QLatin1String("sz"), QLatin1String("8"));
    writer.writeTextElement(QLatin1String("name"), QLatin1String("Calibri"));
    writer.writeEndElement(); // font

    writer.writeEndElement(); // fonts

    // -------------------
    // Fills
    // -------------------
    writer.writeStartElement(QLatin1String("fills"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("2"));

    // First fill must be "none"
    writer.writeStartElement(QLatin1String("fill"));
    writer.writeEmptyElement(QLatin1String("patternFill"));
    writer.writeAttribute(QLatin1String("patternType"), QLatin1String("none"));
    writer.writeEndElement(); // fill

    // Second fill must be "gray125"
    writer.writeStartElement(QLatin1String("fill"));
    writer.writeStartElement(QLatin1String("patternFill"));
    writer.writeAttribute(QLatin1String("patternType"), QLatin1String("gray125"));
    writer.writeEndElement(); // patternFill
    writer.writeEndElement(); // fill

    writer.writeEndElement(); // fills

    // -------------------
    // Borders
    // -------------------
    writer.writeStartElement(QLatin1String("borders"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("1"));
    writer.writeEmptyElement(QLatin1String("border"));
    writer.writeEndElement(); // borders

    // -------------------
    // Critical fix: Add cellStyleXfs
    // -------------------
    writer.writeStartElement(QLatin1String("cellStyleXfs"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("1"));
    writer.writeStartElement(QLatin1String("xf"));
    writer.writeAttribute(QLatin1String("numFmtId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("fontId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("fillId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("borderId"), QLatin1String("0"));
    writer.writeEndElement(); // xf
    writer.writeEndElement(); // cellStyleXfs

    // -------------------
    // CellXfs
    // -------------------
    writer.writeStartElement(QLatin1String("cellXfs"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("2"));

    // Default XF
    writer.writeStartElement(QLatin1String("xf"));
    writer.writeAttribute(QLatin1String("numFmtId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("fontId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("fillId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("borderId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("xfId"), QLatin1String("0")); // Reference to cellStyleXfs
    writer.writeEndElement(); // xf

    // XF with small font + shrinkToFit
    writer.writeStartElement(QLatin1String("xf"));
    writer.writeAttribute(QLatin1String("numFmtId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("fontId"), QLatin1String("1"));
    writer.writeAttribute(QLatin1String("fillId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("borderId"), QLatin1String("0"));
    writer.writeAttribute(QLatin1String("xfId"), QLatin1String("0")); // Reference to cellStyleXfs
    writer.writeAttribute(QLatin1String("applyAlignment"), QLatin1String("1"));

    writer.writeStartElement(QLatin1String("alignment"));
    writer.writeAttribute(QLatin1String("shrinkToFit"), QLatin1String("1"));
    writer.writeEndElement(); // alignment

    writer.writeEndElement(); // xf
    writer.writeEndElement(); // cellXfs

    writer.writeEndElement(); // styleSheet
    writer.writeEndDocument();
}

bool Style::ParseXml(QIODevice* device)
{
    Q_UNUSED(device)

    // QXmlStreamReader reader(device);
    // while (!reader.atEnd() && !reader.hasError()) {
    //     if (reader.readNextStartElement() && reader.name() == QLatin1String("styleSheet")) {
    //         // 样式解析逻辑为空，因为无需处理任何样式
    //     }
    // }
    // return !reader.hasError();

    return true;
}

QT_END_NAMESPACE_YXLSX
