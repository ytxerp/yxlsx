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

#include "contenttype.h"

#include <QBuffer>
#include <QDebug>

QT_BEGIN_NAMESPACE_YXLSX

ContentType::ContentType(OperationMode mode)
    : AbstractOOXmlFile { mode }
    , package_prefix_ { QStringLiteral("application/vnd.openxmlformats-package.") }
    , document_prefix_ { QStringLiteral("application/vnd.openxmlformats-officedocument.") }

{
}

void ContentType::AddDefault(const QString& key, const QString& value) { default_hash_.insert(key, value); }

void ContentType::AddOverride(const QString& key, const QString& value) { override_hash_.insert(key, value); }

void ContentType::AddDocPropApp() { AddOverride(QStringLiteral("/docProps/app.xml"), document_prefix_ + QStringLiteral("extended-properties+xml")); }

void ContentType::AddDocPropCore() { AddOverride(QStringLiteral("/docProps/core.xml"), package_prefix_ + QStringLiteral("core-properties+xml")); }

void ContentType::AddStyles() { AddOverride(QStringLiteral("/xl/styles.xml"), document_prefix_ + QStringLiteral("spreadsheetml.styles+xml")); }

void ContentType::AddWorkbook() { AddOverride(QStringLiteral("/xl/workbook.xml"), document_prefix_ + QStringLiteral("spreadsheetml.sheet.main+xml")); }

void ContentType::AddWorksheetName(const QString& name)
{
    AddOverride(QStringLiteral("/xl/worksheets/%1.xml").arg(name), document_prefix_ + QStringLiteral("spreadsheetml.worksheet+xml"));
}

void ContentType::AddSharedString()
{
    AddOverride(QStringLiteral("/xl/sharedStrings.xml"), document_prefix_ + QStringLiteral("spreadsheetml.sharedStrings+xml"));
}

void ContentType::ClearOverride() { override_hash_.clear(); }

void ContentType::ComposeXml(QIODevice* device) const
{
    QXmlStreamWriter writer(device);

    writer.writeStartDocument(QLatin1String("1.0"), true);
    writer.writeStartElement(QLatin1String("Types"));
    writer.writeAttribute(QLatin1String("xmlns"), QLatin1String("http://schemas.openxmlformats.org/package/2006/content-types"));

    // Compose Default elements
    ComposeElement(writer, default_hash_, QStringLiteral("Default"), QStringLiteral("Extension"), QStringLiteral("ContentType"));

    // Compose Override elements
    ComposeElement(writer, override_hash_, QStringLiteral("Override"), QStringLiteral("PartName"), QStringLiteral("ContentType"));

    writer.writeEndElement(); // Types
    writer.writeEndDocument();
}

void ContentType::ComposeElement(
    QXmlStreamWriter& writer, const QHash<QString, QString>& hash, const QString& element, const QString& key_attr, const QString& value_attr) const
{
    for (auto it = hash.constBegin(); it != hash.constEnd(); ++it) {
        writer.writeStartElement(element);
        writer.writeAttribute(key_attr, it.key());
        writer.writeAttribute(value_attr, it.value());
        writer.writeEndElement();
    }
}

bool ContentType::ParseXml(QIODevice* device)
{
    default_hash_.clear();
    override_hash_.clear();

    QXmlStreamReader reader(device);
    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        const auto name { reader.name() };

        if (name == QStringLiteral("Default")) {
            ParseElement(reader.attributes(), default_hash_, QStringLiteral("Extension"), QStringLiteral("ContentType"));
        } else if (name == QStringLiteral("Override")) {
            ParseElement(reader.attributes(), override_hash_, QStringLiteral("PartName"), QStringLiteral("ContentType"));
        }
    }

    if (reader.hasError()) {
        qDebug() << reader.errorString();
    }

    return true;
}

void ContentType::ParseElement(const QXmlStreamAttributes& attrs, QHash<QString, QString>& hash, const QString& key_attr, const QString& value_attr)
{
    const QString key { attrs.value(key_attr).toString() };
    const QString value { attrs.value(value_attr).toString() };

    if (!key.isEmpty() && !value.isEmpty()) {
        hash.insert(key, value);
    }
}

QT_END_NAMESPACE_YXLSX
