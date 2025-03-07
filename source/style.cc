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

    // Start the styleSheet element with the necessary namespace
    writer.writeStartElement(QLatin1String("styleSheet"));
    writer.writeDefaultNamespace(QLatin1String("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));

    // Write <numFmts> element (Optional: Only include if you need custom number formats)
    writer.writeStartElement(QLatin1String("numFmts"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("1"));
    writer.writeStartElement(QLatin1String("numFmt"));
    writer.writeAttribute(QLatin1String("numFmtId"), QLatin1String("164")); // General format
    writer.writeAttribute(QLatin1String("formatCode"), QLatin1String("General"));
    writer.writeEndElement(); // numFmt
    writer.writeEndElement(); // numFmts

    // Write <borders> element (Defines a border style, here an empty one)
    writer.writeStartElement(QLatin1String("borders"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("1"));
    writer.writeEmptyElement(QLatin1String("border")); // Empty border definition
    writer.writeEndElement(); // borders

    // Write <cellXfs> element (Defines the text style without font and size)
    writer.writeStartElement(QLatin1String("cellXfs"));
    writer.writeAttribute(QLatin1String("count"), QLatin1String("1"));
    writer.writeStartElement(QLatin1String("xf"));
    writer.writeAttribute(QLatin1String("numFmtId"), QLatin1String("0")); // Default number format (General)
    writer.writeAttribute(QLatin1String("borderId"), QLatin1String("0")); // First (and only) border
    writer.writeEndElement(); // xf
    writer.writeEndElement(); // cellXfs

    // End the styleSheet element
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
