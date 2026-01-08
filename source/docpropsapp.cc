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

#include "docpropsapp.h"

#include <QBuffer>
#include <QDebug>
#include <QXmlStreamReader>
#include <QXmlStreamWriter>

QT_BEGIN_NAMESPACE_YXLSX

DocPropsApp::DocPropsApp(OperationMode mode)
    : AbstractOOXmlFile { mode }
{
}

bool DocPropsApp::SetProperty(const QString& name, const QString& value)
{
    static const QSet<QString> kValidKey { QStringLiteral("Manager"), QStringLiteral("Company"), QStringLiteral("Application"), QStringLiteral("DocSecurity"),
        QStringLiteral("ScaleCrop"), QStringLiteral("LinksUpToDate"), QStringLiteral("SharedDoc"), QStringLiteral("HyperlinksChanged"),
        QStringLiteral("AppVersion") };

    // Check if the property name is valid
    if (!kValidKey.contains(name)) {
        return false;
    }

    // Check if the value is empty
    if (value.isEmpty()) {
        return false;
    }

    // Handle special cases for boolean properties
    if (name == QStringLiteral("ScaleCrop") || name == QStringLiteral("LinksUpToDate") || name == QStringLiteral("SharedDoc")
        || name == QStringLiteral("HyperlinksChanged")) {
        // If the value is not "true" or "false", return false
        if (value != QStringLiteral("true") && value != QStringLiteral("false")) {
            qWarning() << "DocPropsApp Invalid boolean value for property:" << name << value;
            return false;
        }
    }

    // Store the property name and value in the hash map
    property_hash_[name] = value;

    return true;
}

void DocPropsApp::ComposeXml(QIODevice* device) const
{
    if (!device || !device->isWritable()) {
        qWarning() << "Invalid or unwritable device.";
        return;
    }

    QXmlStreamWriter writer(device);
    auto vt { QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes") };

    writer.writeStartDocument(QLatin1String("1.0"), true);
    writer.writeStartElement(QLatin1String("Properties"));
    writer.writeDefaultNamespace(QLatin1String("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"));
    writer.writeNamespace(vt, QLatin1String("vt"));

    // Write basic properties
    auto writeProperty = [&writer, this](const QLatin1String& key, const QLatin1String& default_value) {
        auto it = property_hash_.constFind(key);
        writer.writeTextElement(key, it != property_hash_.constEnd() ? it.value() : default_value);
    };

    writeProperty(QLatin1String("Application"), QLatin1String("Microsoft Excel"));
    writeProperty(QLatin1String("DocSecurity"), QLatin1String("0"));
    writeProperty(QLatin1String("ScaleCrop"), QLatin1String("false"));

    // Write other properties
    writeProperty(QLatin1String("Manager"), QLatin1String());
    writeProperty(QLatin1String("Company"), QLatin1String());

    // Write fixed properties
    writeProperty(QLatin1String("LinksUpToDate"), QLatin1String("false"));
    writeProperty(QLatin1String("SharedDoc"), QLatin1String("false"));
    writeProperty(QLatin1String("HyperlinksChanged"), QLatin1String("false"));
    writeProperty(QLatin1String("AppVersion"), QLatin1String("12.0000"));

    writer.writeStartElement(QLatin1String("HeadingPairs"));
    writer.writeStartElement(vt, QLatin1String("vector"));
    writer.writeAttribute(QLatin1String("size"), QString::number(heading_list_.size() * 2));
    writer.writeAttribute(QLatin1String("baseType"), QLatin1String("variant"));

    for (const auto& pair : heading_list_) {
        writer.writeStartElement(vt, QLatin1String("variant"));
        writer.writeTextElement(vt, QLatin1String("lpstr"), pair.first);
        writer.writeEndElement(); // vt:variant
        writer.writeStartElement(vt, QLatin1String("variant"));
        writer.writeTextElement(vt, QLatin1String("i4"), QString::number(pair.second));
        writer.writeEndElement(); // vt:variant
    }
    writer.writeEndElement(); // vt:vector
    writer.writeEndElement(); // HeadingPairs

    writer.writeStartElement(QLatin1String("TitlesOfParts"));
    writer.writeStartElement(vt, QLatin1String("vector"));
    writer.writeAttribute(QLatin1String("size"), QString::number(title_list_.size()));
    writer.writeAttribute(QLatin1String("baseType"), QLatin1String("lpstr"));
    for (const QString& title : title_list_)
        writer.writeTextElement(vt, QLatin1String("lpstr"), title);
    writer.writeEndElement(); // vt:vector
    writer.writeEndElement(); // TitlesOfParts

    writer.writeEndElement(); // Properties
    writer.writeEndDocument();
}

bool DocPropsApp::ParseXml(QIODevice* device)
{
    QXmlStreamReader reader(device);

    while (!reader.atEnd() && !reader.hasError()) {
        if (!reader.readNextStartElement())
            continue;

        const auto name { reader.name() };

        if (name == QStringLiteral("Manager")) {
            SetProperty(QStringLiteral("manager"), reader.readElementText());
        } else if (name == QStringLiteral("Company")) {
            SetProperty(QStringLiteral("company"), reader.readElementText());
        } else if (name == QStringLiteral("Application")) {
            SetProperty(QStringLiteral("Application"), reader.readElementText());
        } else if (name == QStringLiteral("DocSecurity")) {
            SetProperty(QStringLiteral("DocSecurity"), reader.readElementText());
        } else if (name == QStringLiteral("ScaleCrop")) {
            SetProperty(QStringLiteral("ScaleCrop"), reader.readElementText());
        } else if (name == QStringLiteral("LinksUpToDate")) {
            SetProperty(QStringLiteral("LinksUpToDate"), reader.readElementText());
        } else if (name == QStringLiteral("SharedDoc")) {
            SetProperty(QStringLiteral("SharedDoc"), reader.readElementText());
        } else if (name == QStringLiteral("HyperlinksChanged")) {
            SetProperty(QStringLiteral("HyperlinksChanged"), reader.readElementText());
        } else if (name == QStringLiteral("AppVersion")) {
            SetProperty(QStringLiteral("AppVersion"), reader.readElementText());
        }
    }

    if (reader.hasError()) {
        qDebug() << "DocPropsApp Error reading doc props app file:" << reader.errorString();
        return false;
    }

    return true;
}

QT_END_NAMESPACE_YXLSX
